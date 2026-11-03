from __future__ import annotations

import argparse
import base64
import json
import mimetypes
import os
import secrets
import socket
import subprocess
import sys
import time
import traceback
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from uuid import uuid4

from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import pkcs12
from cryptography.x509 import load_der_x509_certificate


JSON_ACCEPT = "application/json;odata=verbose"
HTTP_RETRYABLE_ERRNOS = {-3}
HTTP_MAX_ATTEMPTS = max(1, int(os.getenv("SHAREPOINT_HTTP_MAX_ATTEMPTS", "4")))
HTTP_RETRY_BASE_DELAY_SECONDS = max(0.0, float(os.getenv("SHAREPOINT_HTTP_RETRY_BASE_DELAY", "1.5")))
MONTH_NAMES_PT = {
    1: "Janeiro",
    2: "Fevereiro",
    3: "Março",
    4: "Abril",
    5: "Maio",
    6: "Junho",
    7: "Julho",
    8: "Agosto",
    9: "Setembro",
    10: "Outubro",
    11: "Novembro",
    12: "Dezembro",
}

FISCAL_ROTINA_MENSAL_PENDENCIA_ITEMS = (
    "Notas fiscais de compras;",
    "Notas de serviços tomados;",
    "Notas de serviços prestados (se houver informar acesso e senha da Prefeitura);",
    "Relação de notas de entrada do sistema operacional da empresa (se houver);",
    "Demonstrativo de aluguel (se houver);",
    "Conta energia elétrica;",
)

FISCAL_ROTINA_MENSAL_SIMPLES_DEADLINE_DAY = 15
FISCAL_ROTINA_MENSAL_SIMPLES_VENCIMENTO_MARGEM_DAY = 17

RED_FILL_RGB_SUFFIX = "FF0000"
YELLOW_FILL_RGB_SUFFIX = "FFFF00"
# Azul do Excel vem como cor de tema (sem RGB): 4 = Ênfase 1, 8 = Ênfase 5.
BLUE_FILL_THEMES = {"4", "8"}


@dataclass(frozen=True)
class MonthlyRoutineDefinition:
    routine_key: str
    area_id: int
    processo_id: int | None
    competencia_month_offset: int
    deadline_month_offset: int
    deadline_day: int | None
    process_title_template: str | None = None
    process_month_offset: int | None = None
    assignee_field_name: str = "ResponsavelFolha"
    due_day_source: str = "fixed"
    sync_sheet_name: str | None = None
    schedule_day: int = 25
    # due_day_source="client": prazo = N dias úteis antes do deadline_day do cliente.
    deadline_business_days_before: int = 0
    # Datas das etapas como o portal grava nas tarefas abertas à mão (Execução -4/-3, Vencimento -2).
    portal_manual_stage_dates: bool = False


MONTHLY_ROUTINES: dict[str, MonthlyRoutineDefinition] = {
    "folha_mensal": MonthlyRoutineDefinition(
        routine_key="folha_mensal",
        area_id=3,
        processo_id=9,
        competencia_month_offset=0,
        deadline_month_offset=1,
        deadline_day=7,
    ),
    "fgts_dctfweb_mensal": MonthlyRoutineDefinition(
        routine_key="fgts_dctfweb_mensal",
        area_id=3,
        processo_id=110,
        competencia_month_offset=-1,
        deadline_month_offset=0,
        deadline_day=20,
    ),
    "dctfweb_mensal": MonthlyRoutineDefinition(
        routine_key="dctfweb_mensal",
        area_id=3,
        processo_id=103,
        competencia_month_offset=-1,
        deadline_month_offset=0,
        deadline_day=20,
    ),
    # Prazo por cliente: 2 dias úteis antes do dia do adiantamento (deadline_day = 15 ou 20),
    # no mês da execução — a regra das tarefas abertas à mão até 09/2026.
    "adiantamento_mensal": MonthlyRoutineDefinition(
        routine_key="adiantamento_mensal",
        area_id=3,
        processo_id=34,
        competencia_month_offset=0,
        deadline_month_offset=0,
        deadline_day=None,
        due_day_source="client",
        deadline_business_days_before=2,
        portal_manual_stage_dates=True,
    ),
    "fiscal_rotina_mensal": MonthlyRoutineDefinition(
        routine_key="fiscal_rotina_mensal",
        area_id=2,
        processo_id=5,
        competencia_month_offset=0,
        deadline_month_offset=1,
        deadline_day=10,
        assignee_field_name="ResponsavelFiscal",
        due_day_source="fixed",
        sync_sheet_name="ROTINA MENSAL",
    ),
    "fiscal_pis_cofins_mes_atual": MonthlyRoutineDefinition(
        routine_key="fiscal_pis_cofins_mes_atual",
        area_id=2,
        processo_id=None,
        competencia_month_offset=0,
        deadline_month_offset=1,
        deadline_day=20,
        process_title_template="Fiscal - Apuração PIS/COFINS {month_name_upper}",
        assignee_field_name="ResponsavelFiscal",
        due_day_source="fixed",
        sync_sheet_name="PIS - COFINS MÊS ATUAL",
    ),
    "fiscal_sped_contribuicoes_mes_anterior": MonthlyRoutineDefinition(
        routine_key="fiscal_sped_contribuicoes_mes_anterior",
        area_id=2,
        processo_id=None,
        competencia_month_offset=0,
        deadline_month_offset=1,
        deadline_day=10,
        process_title_template="Fiscal - SPED Contribuições {month_name_upper}",
        process_month_offset=-1,
        assignee_field_name="ResponsavelFiscal",
        due_day_source="fixed",
        sync_sheet_name="SPED CONTRIBUIÇÕES MÊS ANTERIOR",
    ),
    "fiscal_parcelamento_tributos": MonthlyRoutineDefinition(
        routine_key="fiscal_parcelamento_tributos",
        area_id=2,
        processo_id=108,
        competencia_month_offset=0,
        deadline_month_offset=1,
        deadline_day=20,
        assignee_field_name="ResponsavelFiscal",
        due_day_source="fixed",
        sync_sheet_name="PARCELAMENTO DE TRIBUTOS",
    ),
}

FISCAL_WORKBOOK_ROUTINES = (
    "fiscal_rotina_mensal",
    "fiscal_pis_cofins_mes_atual",
    "fiscal_sped_contribuicoes_mes_anterior",
    "fiscal_parcelamento_tributos",
)

WORKBOOK_HEADERS = ("Título", "CNPJ", "Regime", "Atividade", "Cliente Ativo (S/N)", "Colaborador Fiscal")


def b64url(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def iso_now() -> int:
    return int(time.time())


def iso_utc_now_string() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def is_retryable_url_error(exc: urllib.error.URLError) -> bool:
    reason = exc.reason
    if isinstance(reason, socket.gaierror):
        return reason.errno in HTTP_RETRYABLE_ERRNOS
    if isinstance(reason, TimeoutError):
        return True
    if isinstance(reason, socket.timeout):
        return True
    return False


def make_request(
    url: str,
    *,
    method: str = "GET",
    headers: dict[str, str] | None = None,
    body: bytes | None = None,
) -> tuple[int, dict[str, str], bytes]:
    request = urllib.request.Request(url, data=body, method=method)
    for key, value in (headers or {}).items():
        request.add_header(key, value)
    last_error: Exception | None = None
    for attempt in range(1, HTTP_MAX_ATTEMPTS + 1):
        try:
            with urllib.request.urlopen(request) as response:
                return response.status, dict(response.headers.items()), response.read()
        except urllib.error.HTTPError as exc:
            payload = exc.read()
            message = payload.decode("utf-8", errors="replace")
            raise RuntimeError(f"{method} {url} falhou com HTTP {exc.code}: {message}") from exc
        except urllib.error.URLError as exc:
            last_error = exc
            if attempt >= HTTP_MAX_ATTEMPTS or not is_retryable_url_error(exc):
                break
            time.sleep(HTTP_RETRY_BASE_DELAY_SECONDS * attempt)
    if last_error is not None:
        raise RuntimeError(
            f"{method} {url} falhou após {HTTP_MAX_ATTEMPTS} tentativas: {last_error}"
        ) from last_error
    raise RuntimeError(f"{method} {url} falhou por motivo desconhecido.")


def make_json_request(
    url: str,
    *,
    method: str = "GET",
    headers: dict[str, str] | None = None,
    payload: dict[str, Any] | None = None,
) -> dict[str, Any]:
    request_headers = {"Accept": JSON_ACCEPT, **(headers or {})}
    body = None
    if payload is not None:
        body = json.dumps(payload).encode("utf-8")
        request_headers.setdefault("Content-Type", "application/json;odata=verbose")
    _, _, response_body = make_request(url, method=method, headers=request_headers, body=body)
    if not response_body.strip():
        return {}
    raw = response_body.decode("utf-8", errors="replace")
    try:
        return json.loads(raw)
    except json.JSONDecodeError as exc:
        snippet = raw[:600].replace("\n", "\\n")
        raise RuntimeError(f"Resposta não JSON em {method} {url}: {snippet}") from exc


@dataclass
class SharePointSettings:
    tenant_id: str
    client_id: str
    cert_pfx_path: Path
    cert_password: str
    task_site_url: str
    task_list_guid: str

    @classmethod
    def from_env(cls) -> "SharePointSettings":
        missing = [
            name
            for name in (
                "SHAREPOINT_TENANT_ID",
                "SHAREPOINT_CLIENT_ID",
                "SHAREPOINT_CERT_PFX_PATH",
                "SHAREPOINT_CERT_PASSWORD",
                "SHAREPOINT_SITE_URL",
                "SHAREPOINT_TASK_LIST_GUID",
            )
            if not os.getenv(name)
        ]
        if missing:
            raise RuntimeError(f"Variáveis obrigatórias ausentes: {', '.join(missing)}")
        return cls(
            tenant_id=os.environ["SHAREPOINT_TENANT_ID"].strip(),
            client_id=os.environ["SHAREPOINT_CLIENT_ID"].strip(),
            cert_pfx_path=Path(os.environ["SHAREPOINT_CERT_PFX_PATH"]).expanduser(),
            cert_password=os.environ["SHAREPOINT_CERT_PASSWORD"],
            task_site_url=os.environ["SHAREPOINT_SITE_URL"].strip().rstrip("/"),
            task_list_guid=os.environ["SHAREPOINT_TASK_LIST_GUID"].strip(),
        )


class SharePointClient:
    def __init__(self, settings: SharePointSettings) -> None:
        self.settings = settings
        self._access_token: str | None = None
        self._token_expiry = 0

    def _load_certificate_material(self) -> tuple[Any, str]:
        pfx_data = self.settings.cert_pfx_path.read_bytes()
        private_key, certificate, _ = pkcs12.load_key_and_certificates(
            pfx_data,
            self.settings.cert_password.encode("utf-8"),
        )
        if private_key is None or certificate is None:
            raise RuntimeError("Não foi possível carregar chave privada e certificado do PFX.")
        cert_der = certificate.public_bytes(serialization.Encoding.DER)
        thumbprint = load_der_x509_certificate(cert_der).fingerprint(hashes.SHA1())
        return private_key, b64url(thumbprint)

    def _build_client_assertion(self) -> str:
        private_key, thumbprint = self._load_certificate_material()
        now = iso_now()
        header = {"alg": "RS256", "typ": "JWT", "x5t": thumbprint}
        payload = {
            "aud": f"https://login.microsoftonline.com/{self.settings.tenant_id}/oauth2/v2.0/token",
            "iss": self.settings.client_id,
            "sub": self.settings.client_id,
            "jti": str(uuid4()),
            "nbf": now - 60,
            "exp": now + 600,
        }
        signing_input = ".".join(
            (
                b64url(json.dumps(header, separators=(",", ":")).encode("utf-8")),
                b64url(json.dumps(payload, separators=(",", ":")).encode("utf-8")),
            )
        ).encode("ascii")
        signature = private_key.sign(signing_input, padding.PKCS1v15(), hashes.SHA256())
        return f"{signing_input.decode('ascii')}.{b64url(signature)}"

    def access_token(self, resource_site_url: str) -> str:
        now = iso_now()
        if self._access_token and now < self._token_expiry - 60:
            return self._access_token

        assertion = self._build_client_assertion()
        site_host = urllib.parse.urlparse(resource_site_url).netloc
        token_url = f"https://login.microsoftonline.com/{self.settings.tenant_id}/oauth2/v2.0/token"
        form = urllib.parse.urlencode(
            {
                "client_id": self.settings.client_id,
                "scope": f"https://{site_host}/.default",
                "grant_type": "client_credentials",
                "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
                "client_assertion": assertion,
            }
        ).encode("utf-8")
        _, _, response_body = make_request(
            token_url,
            method="POST",
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            body=form,
        )
        payload = json.loads(response_body.decode("utf-8"))
        self._access_token = payload["access_token"]
        self._token_expiry = now + int(payload.get("expires_in", 3600))
        return self._access_token

    def rest_json(
        self,
        site_url: str,
        api_path: str,
        *,
        method: str = "GET",
        payload: dict[str, Any] | None = None,
        headers: dict[str, str] | None = None,
    ) -> dict[str, Any]:
        token = self.access_token(site_url)
        base_headers = {"Authorization": f"Bearer {token}", "Accept": JSON_ACCEPT}
        if payload is not None:
            base_headers["Content-Type"] = "application/json;odata=verbose"
        if headers:
            base_headers.update(headers)
        data = payload if payload is not None else None
        return make_json_request(f"{site_url}{api_path}", method=method, headers=base_headers, payload=data)

    def upload_file(self, site_url: str, folder_server_relative_url: str, file_path: Path) -> dict[str, Any]:
        token = self.access_token(site_url)
        file_name = file_path.name.replace("'", "''")
        mime_type, _ = mimetypes.guess_type(file_path.name)
        _, _, response_body = make_request(
            f"{site_url}/_api/web/GetFolderByServerRelativeUrl('{quote_odata_url_value(folder_server_relative_url)}')/Files/add(url='{urllib.parse.quote(file_name, safe='')}',overwrite=true)",
            method="POST",
            headers={
                "Authorization": f"Bearer {token}",
                "Accept": JSON_ACCEPT,
                "Content-Type": mime_type or "application/octet-stream",
            },
            body=file_path.read_bytes(),
        )
        return json.loads(response_body.decode("utf-8"))["d"]


@dataclass
class PostgresSettings:
    url: str

    @classmethod
    def from_env(cls) -> "PostgresSettings":
        if not os.getenv("POSTGRES_URL"):
            raise RuntimeError("Variável obrigatória ausente: POSTGRES_URL")
        return cls(url=os.environ["POSTGRES_URL"].strip())


class PostgresClient:
    def __init__(self, settings: PostgresSettings) -> None:
        self.settings = settings

    def query_json(self, sql: str) -> list[dict[str, Any]]:
        raw = self.psql_scalar(sql)
        if not raw:
            return []
        return json.loads(raw)

    def psql_scalar(self, sql: str) -> str:
        # SQL via stdin: passar como argumento (-c) estoura o limite de 128 KB por
        # argumento do Linux no sync-fiscal-workbook (~700 inserts).
        completed = subprocess.run(
            ["psql", self.settings.url, "-X", "-q", "-t", "-A", "-v", "ON_ERROR_STOP=1"],
            input=sql,
            check=False,
            text=True,
            capture_output=True,
        )
        if completed.returncode != 0:
            raise RuntimeError(completed.stderr.strip() or completed.stdout.strip())
        return completed.stdout.strip()

    def execute(self, sql: str) -> None:
        self.psql_scalar(sql)


def quote_odata(value: str) -> str:
    return value.replace("'", "''")


def quote_odata_url_value(value: str) -> str:
    return urllib.parse.quote(quote_odata(value), safe="/()'")


def site_server_relative_url(site_url: str) -> str:
    parsed = urllib.parse.urlparse(site_url)
    return parsed.path.rstrip("/")


def first_result(payload: dict[str, Any]) -> dict[str, Any]:
    results = payload["d"]["results"]
    if not results:
        raise RuntimeError("Nenhum resultado retornado.")
    return results[0]


def sql_literal(value: Any) -> str:
    if value is None:
        return "null"
    return "'" + str(value).replace("'", "''") + "'"


def jsonb_literal(value: Any) -> str:
    return f"{sql_literal(json.dumps(value, ensure_ascii=False))}::jsonb"


def parse_run_date(value: str | None) -> date:
    if not value:
        return datetime.now(timezone.utc).date()
    return date.fromisoformat(value)


def add_months(base_date: date, month_offset: int) -> date:
    month_index = (base_date.year * 12 + (base_date.month - 1)) + month_offset
    year = month_index // 12
    month = (month_index % 12) + 1
    return date(year, month, 1)


def month_name_upper(month: int) -> str:
    return MONTH_NAMES_PT[month].upper()


def normalize_cnpj(value: str | None) -> str:
    return "".join(char for char in str(value or "") if char.isdigit())


def normalize_client_title(value: str | None) -> str:
    return " ".join(str(value or "").split()).casefold()


def iso_z(dt: datetime) -> str:
    return dt.astimezone(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def workbook_cell_text(cell: ET.Element, shared_strings: list[str], ns: dict[str, str]) -> str:
    cell_type = cell.attrib.get("t")
    value = cell.find("a:v", ns)
    if cell_type == "inlineStr":
        return "".join(part.text or "" for part in cell.findall(".//a:t", ns)).strip()
    if value is None or value.text is None:
        return ""
    raw = value.text
    if cell_type == "s":
        return shared_strings[int(raw)].strip()
    return raw.strip()


def load_workbook_style_highlights(archive: zipfile.ZipFile) -> dict[int, str]:
    ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    if "xl/styles.xml" not in archive.namelist():
        return {}
    styles = ET.fromstring(archive.read("xl/styles.xml"))
    fills = styles.find("a:fills", ns)
    highlight_by_fill_id: dict[int, str] = {}
    for fill_id, fill in enumerate([] if fills is None else fills.findall("a:fill", ns)):
        pattern_fill = fill.find("a:patternFill", ns)
        if pattern_fill is None:
            continue
        fg_color = pattern_fill.find("a:fgColor", ns)
        if fg_color is None:
            continue
        rgb = (fg_color.attrib.get("rgb") or "").upper()
        theme = fg_color.attrib.get("theme")
        if rgb.endswith(RED_FILL_RGB_SUFFIX):
            highlight_by_fill_id[fill_id] = "vermelho"
        elif rgb.endswith(YELLOW_FILL_RGB_SUFFIX):
            highlight_by_fill_id[fill_id] = "amarelo"
        elif theme in BLUE_FILL_THEMES:
            highlight_by_fill_id[fill_id] = "azul"

    cell_xfs = styles.find("a:cellXfs", ns)
    highlight_by_style_index: dict[int, str] = {}
    for style_index, xf in enumerate([] if cell_xfs is None else cell_xfs.findall("a:xf", ns)):
        fill_id = xf.attrib.get("fillId")
        highlight = highlight_by_fill_id.get(int(fill_id)) if fill_id is not None else None
        if highlight:
            highlight_by_style_index[style_index] = highlight
    return highlight_by_style_index


def load_workbook_rows(workbook_path: Path) -> dict[str, list[dict[str, str]]]:
    ns = {
        "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }
    rows_by_sheet: dict[str, list[dict[str, str]]] = {}
    with zipfile.ZipFile(workbook_path) as archive:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall("a:si", ns):
                shared_strings.append("".join(part.text or "" for part in item.findall(".//a:t", ns)))

        highlight_by_style_index = load_workbook_style_highlights(archive)

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relations = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in relations}
        sheets = workbook.find("a:sheets", ns)
        for sheet in ([] if sheets is None else list(sheets)):
            sheet_name = sheet.attrib["name"]
            rel_id = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            target = relmap[rel_id].lstrip("/")
            if not target.startswith("xl/"):
                target = f"xl/{target}"
            sheet_root = ET.fromstring(archive.read(target))
            raw_rows: list[list[str]] = []
            raw_row_highlights: list[str | None] = []
            for row in sheet_root.findall(".//a:sheetData/a:row", ns):
                cells = row.findall("a:c", ns)
                values = [workbook_cell_text(cell, shared_strings, ns) for cell in cells]
                raw_rows.append(values)
                row_highlight = None
                for cell in cells:
                    style_attr = cell.attrib.get("s")
                    if style_attr is not None and int(style_attr) in highlight_by_style_index:
                        row_highlight = highlight_by_style_index[int(style_attr)]
                        break
                raw_row_highlights.append(row_highlight)
            if not raw_rows:
                rows_by_sheet[sheet_name] = []
                continue
            headers = raw_rows[0]
            if tuple(headers[: len(WORKBOOK_HEADERS)]) != WORKBOOK_HEADERS:
                raise RuntimeError(
                    f"Aba '{sheet_name}' possui cabeçalho inesperado. "
                    f"Esperado: {', '.join(WORKBOOK_HEADERS)}"
                )
            sheet_rows: list[dict[str, str]] = []
            for values, row_highlight in zip(raw_rows[1:], raw_row_highlights[1:]):
                if not any(values):
                    continue
                padded = values + [""] * (len(headers) - len(values))
                row_map = {headers[index]: padded[index].strip() for index in range(len(headers))}
                row_map["_highlight"] = row_highlight
                sheet_rows.append(row_map)
            rows_by_sheet[sheet_name] = sheet_rows
    return rows_by_sheet


def easter_sunday(year: int) -> date:
    a, b, c = year % 19, year // 100, year % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = (h + l - 7 * m + 114) % 31 + 1
    return date(year, month, day)


def national_holidays(year: int) -> set[date]:
    fixed = {(1, 1), (4, 21), (5, 1), (9, 7), (10, 12), (11, 2), (11, 15), (11, 20), (12, 25)}
    holidays = {date(year, month, day) for month, day in fixed}
    holidays.add(easter_sunday(year) - timedelta(days=2))  # Sexta-feira Santa
    return holidays


def is_business_day(day: date) -> bool:
    return day.weekday() < 5 and day not in national_holidays(day.year)


def business_days_before(day: date, count: int) -> date:
    current = day
    while count > 0:
        current -= timedelta(days=1)
        if is_business_day(current):
            count -= 1
    return current


def client_day_schedule(definition: MonthlyRoutineDefinition, prazo_mes: str, deadline_day: int) -> dict[str, str]:
    """Datas da tarefa para um cliente com prazo por dia (due_day_source="client").

    Horário 03:00Z = meia-noite de Brasília, como o portal grava nas tarefas manuais.
    """
    year, month = (int(part) for part in prazo_mes.split("-"))
    prazo_day = business_days_before(date(year, month, deadline_day), definition.deadline_business_days_before)
    prazo = datetime(prazo_day.year, prazo_day.month, prazo_day.day, 3, tzinfo=timezone.utc)
    if not definition.portal_manual_stage_dates:
        return {key: iso_z(prazo) for key in ("prazo", "execucao_prazo", "execucao_margem", "vencimento", "vencimento_margem")}
    return {
        "prazo": iso_z(prazo),
        "execucao_prazo": iso_z(prazo - timedelta(days=4)),
        "execucao_margem": iso_z(prazo - timedelta(days=3)),
        "vencimento": iso_z(prazo - timedelta(days=2)),
        "vencimento_margem": iso_z(prazo),
    }


def monthly_routine_definition(routine_key: str) -> MonthlyRoutineDefinition:
    definition = MONTHLY_ROUTINES.get(routine_key)
    if definition is None:
        available = ", ".join(sorted(MONTHLY_ROUTINES))
        raise RuntimeError(f"Rotina mensal desconhecida: {routine_key}. Disponíveis: {available}")
    return definition


def monthly_routine_dates(routine_key: str, run_date: date) -> dict[str, str]:
    definition = monthly_routine_definition(routine_key)
    competencia_base = add_months(run_date, definition.competencia_month_offset)
    execucao_inicio = datetime(run_date.year, run_date.month, run_date.day, 12, tzinfo=timezone.utc)
    result = {
        "period_reference": f"{competencia_base.year:04d}-{competencia_base.month:02d}",
        "competencia": f"{MONTH_NAMES_PT[competencia_base.month]} {competencia_base.year}",
        "execucao_inicio": execucao_inicio.isoformat().replace("+00:00", "Z"),
    }
    deadline_base = add_months(run_date, definition.deadline_month_offset)
    if definition.deadline_day is not None:
        deadline = datetime(deadline_base.year, deadline_base.month, definition.deadline_day, tzinfo=timezone.utc)
        result["prazo"] = iso_z(deadline)
    elif definition.due_day_source == "client":
        result["prazo_mes"] = f"{deadline_base.year:04d}-{deadline_base.month:02d}"
    return result


def configured_stage_assignee(stage: str) -> str | None:
    value = os.getenv(f"SHAREPOINT_{stage.upper()}_ASSIGNEE", "").strip()
    return value or None


def resolve_user_info_id(client: SharePointClient, assignee: str | int | None) -> int | None:
    if assignee is None:
        return None
    raw = str(assignee).strip()
    if not raw:
        return None
    if raw.isdigit():
        return int(raw)

    normalized = raw.casefold()
    user_sources = (
        ("/_api/web/siteuserinfolist/items?$select=ID,Title,EMail&$top=5000", "ID", "EMail"),
        ("/_api/web/siteusers?$select=Id,Title,Email&$top=5000", "Id", "Email"),
    )
    for api_path, id_key, email_key in user_sources:
        response = client.rest_json(client.settings.task_site_url, api_path)
        for row in response["d"]["results"]:
            title = str(row.get("Title") or "").strip()
            email = str(row.get(email_key) or "").strip()
            if normalized in {title.casefold(), email.casefold()}:
                return int(row[id_key])

    raise RuntimeError(
        f"Usuário configurado para SharePoint não foi encontrado no site: {raw}. "
        "Use o ID interno, o título exato ou o e-mail."
    )


def resolve_stage_assignee_id(
    client: SharePointClient,
    stage: str,
    explicit_assignee: int | None = None,
) -> int | None:
    if explicit_assignee is not None:
        return explicit_assignee
    configured = configured_stage_assignee(stage)
    if configured is None:
        return None
    return resolve_user_info_id(client, configured)


def fetch_task(client: SharePointClient, task_id: int) -> dict[str, Any]:
    query = urllib.parse.urlencode(
        {
            "$filter": f"ID eq {task_id}",
            "$select": "ID,Id,Title,Competencia,Cliente/Title,Cliente/BibliotecaPublicacao,Cliente/PastaPublicacao,Processo/Title,Area/Title",
            "$expand": "Cliente,Processo,Area",
        }
    )
    api_path = f"/_api/Web/Lists(guid'{client.settings.task_list_guid}')/Items?{query}"
    return first_result(client.rest_json(client.settings.task_site_url, api_path))


def fetch_client_responsavel_folha(client: SharePointClient, cliente_id: int) -> dict[str, Any]:
    query = urllib.parse.urlencode(
        {
            "$filter": f"ID eq {cliente_id}",
            "$select": "ID,Title,ResponsavelFolha/ID,ResponsavelFolha/Title",
            "$expand": "ResponsavelFolha",
        }
    )
    api_path = f"/_api/web/lists/getbytitle('Clientes')/items?{query}"
    row = first_result(client.rest_json(client.settings.task_site_url, api_path))
    responsavel = row.get("ResponsavelFolha") or {}
    if not responsavel.get("ID"):
        raise RuntimeError(f"Cliente {cliente_id} não possui ResponsavelFolha configurado.")
    return {"id": responsavel["ID"], "title": responsavel.get("Title"), "cliente_title": row.get("Title")}


def fetch_client_responsavel_fiscal(client: SharePointClient, cliente_id: int) -> dict[str, Any]:
    query = urllib.parse.urlencode(
        {
            "$filter": f"ID eq {cliente_id}",
            "$select": "ID,Title,ResponsavelFiscal/ID,ResponsavelFiscal/Title",
            "$expand": "ResponsavelFiscal",
        }
    )
    api_path = f"/_api/web/lists/getbytitle('Clientes')/items?{query}"
    row = first_result(client.rest_json(client.settings.task_site_url, api_path))
    responsavel = row.get("ResponsavelFiscal") or {}
    if not responsavel.get("ID"):
        raise RuntimeError(f"Cliente {cliente_id} não possui ResponsavelFiscal configurado.")
    return {"id": responsavel["ID"], "title": responsavel.get("Title"), "cliente_title": row.get("Title")}


def fetch_client_by_title(client: SharePointClient, client_title: str) -> dict[str, Any]:
    query = urllib.parse.urlencode(
        {
            "$filter": f"Title eq '{quote_odata(client_title)}'",
            "$select": "ID,Id,Title,ResponsavelFolha/ID,ResponsavelFolha/Title",
            "$expand": "ResponsavelFolha",
        }
    )
    api_path = f"/_api/web/lists/getbytitle('Clientes')/items?{query}"
    rows = client.rest_json(client.settings.task_site_url, api_path)["d"]["results"]
    if not rows:
        raise RuntimeError(f"Cliente não encontrado na lista Clientes: {client_title}")
    if len(rows) > 1:
        ids = ", ".join(str(row.get("ID")) for row in rows)
        raise RuntimeError(f"Mais de um cliente encontrado com o título '{client_title}': {ids}")
    return rows[0]


def fetch_all_clients_index(client: SharePointClient) -> list[dict[str, Any]]:
    query = urllib.parse.urlencode(
        {
            "$select": "ID,Title,CNPJ,DiaFiscal,ResponsavelFiscal/ID,ResponsavelFiscal/Title,ResponsavelFolha/ID,ResponsavelFolha/Title",
            "$expand": "ResponsavelFiscal,ResponsavelFolha",
            "$top": "5000",
        }
    )
    api_path = f"/_api/web/lists/getbytitle('Clientes')/items?{query}"
    return client.rest_json(client.settings.task_site_url, api_path)["d"]["results"]


def fetch_all_processes_index(client: SharePointClient) -> list[dict[str, Any]]:
    query = urllib.parse.urlencode({"$select": "ID,Title", "$top": "5000"})
    api_path = f"/_api/web/lists/getbytitle('Processos')/items?{query}"
    return client.rest_json(client.settings.task_site_url, api_path)["d"]["results"]


def fetch_task_for_finalize(client: SharePointClient, task_id: int) -> dict[str, Any]:
    query = urllib.parse.urlencode(
        {
            "$filter": f"ID eq {task_id}",
            "$select": "*,Cliente/ID,Cliente/Title,Cliente/BibliotecaPublicacao,Cliente/PastaPublicacao,Processo/ID,Processo/Title,Area/ID,Area/Title,UsuarioAtual/ID,UsuarioAtual/Title,UsuarioAtual/EMail",
            "$expand": "Cliente,Processo,Area,UsuarioAtual",
        }
    )
    api_path = f"/_api/Web/Lists(guid'{client.settings.task_list_guid}')/Items?{query}"
    return first_result(client.rest_json(client.settings.task_site_url, api_path))


def fetch_tasks_for_fiscal_rotina_mensal_pendencia(
    client: SharePointClient,
    competencia: str,
    passo_atual: str = "Execucao",
) -> list[dict[str, Any]]:
    query = urllib.parse.urlencode(
        {
            "$filter": f"ProcessoId eq 5 and Competencia eq '{quote_odata(competencia)}' and PassoAtual eq '{quote_odata(passo_atual)}'",
            "$select": "ID,Id,Title,Competencia,PassoAtual,ExecucaoInicio,ExecucaoPrazo,PendenciaCliente,ClienteId",
            "$top": "5000",
        }
    )
    api_path = f"/_api/Web/Lists(guid'{client.settings.task_list_guid}')/Items?{query}"
    return client.rest_json(client.settings.task_site_url, api_path)["d"]["results"]


def fetch_clients_ativar_pendencia_fiscal_flags(
    client: SharePointClient,
    cliente_ids: list[int],
) -> dict[int, bool]:
    flags: dict[int, bool] = {}
    unique_ids = sorted(set(int(cid) for cid in cliente_ids))
    chunk_size = 50
    for start in range(0, len(unique_ids), chunk_size):
        chunk = unique_ids[start : start + chunk_size]
        query = urllib.parse.urlencode(
            {
                "$filter": " or ".join(f"ID eq {cid}" for cid in chunk),
                "$select": "ID,AtivarSistemaPendenciaFiscal",
                "$top": "5000",
            }
        )
        api_path = f"/_api/web/lists/getbytitle('Clientes')/items?{query}"
        rows = client.rest_json(client.settings.task_site_url, api_path)["d"]["results"]
        for row in rows:
            flags[int(row["ID"])] = bool(row.get("AtivarSistemaPendenciaFiscal"))
    return flags


def resolve_library(client: SharePointClient, client_site_url: str, expected_title: str) -> dict[str, Any]:
    query = urllib.parse.urlencode(
        {
            "$select": "Id,Title,ListItemEntityTypeFullName,BaseTemplate,RootFolder/ServerRelativeUrl,RootFolder/Name",
            "$expand": "RootFolder",
        }
    )
    response = client.rest_json(client_site_url, f"/_api/web/lists?{query}")
    site_root = site_server_relative_url(client_site_url)
    matches = [
        row
        for row in response["d"]["results"]
        if row.get("BaseTemplate") == 101
        and row.get("Title") == expected_title
        and row["RootFolder"]["ServerRelativeUrl"].startswith(site_root + "/")
    ]
    if not matches:
        raise RuntimeError(f"Nenhuma biblioteca encontrada com título '{expected_title}' em {client_site_url}.")
    if len(matches) > 1:
        paths = ", ".join(match["RootFolder"]["ServerRelativeUrl"] for match in matches)
        raise RuntimeError(f"Mais de uma biblioteca encontrada com título '{expected_title}': {paths}")
    return matches[0]


def fetch_uploaded_item(
    client: SharePointClient,
    client_site_url: str,
    file_server_relative_url: str,
) -> dict[str, Any]:
    query = urllib.parse.urlencode({"$select": "ID,Id,__metadata/type"})
    api_path = f"/_api/web/GetFileByServerRelativeUrl('{quote_odata_url_value(file_server_relative_url)}')/ListItemAllFields?{query}"
    return client.rest_json(client_site_url, api_path)["d"]


def update_file_metadata(
    client: SharePointClient,
    client_site_url: str,
    file_server_relative_url: str,
    item_type: str,
    metadata: dict[str, Any],
) -> None:
    api_path = f"/_api/web/GetFileByServerRelativeUrl('{quote_odata_url_value(file_server_relative_url)}')/ListItemAllFields"
    payload = {"__metadata": {"type": item_type}, **metadata}
    client.rest_json(
        client_site_url,
        api_path,
        method="POST",
        payload=payload,
        headers={"X-HTTP-Method": "MERGE", "IF-MATCH": "*"},
    )


def update_task_item(
    client: SharePointClient,
    task_id: int,
    item_type: str,
    payload: dict[str, Any],
) -> None:
    api_path = f"/_api/web/lists/getbytitle('Tarefas')/items({task_id})"
    client.rest_json(
        client.settings.task_site_url,
        api_path,
        method="POST",
        payload={"__metadata": {"type": item_type}, **payload},
        headers={"X-HTTP-Method": "MERGE", "IF-MATCH": "*"},
    )


def create_task_item(client: SharePointClient, payload: dict[str, Any]) -> dict[str, Any]:
    api_path = "/_api/web/lists/getbytitle('Tarefas')/items"
    return client.rest_json(
        client.settings.task_site_url,
        api_path,
        method="POST",
        payload={"__metadata": {"type": "SP.Data.TarefasListItem"}, **payload},
    )["d"]


def build_finalize_payload(
    task: dict[str, Any],
    comment_html: str | None = None,
    usuario_atual_id: int | None = None,
) -> dict[str, Any]:
    now = iso_utc_now_string()
    execucao_usuario_atual_id = task.get("UsuarioAtual", {}).get("ID") or task.get("UsuarioAtualId")
    payload = {
        "Id": task["Id"],
        "ID": task["ID"],
        "Title": "Revisão e Validação",
        "Action": "Aprovar",
        "Descricao": task.get("Descricao"),
        "Competencia": task.get("Competencia"),
        "CheckListOption": task.get("CheckListOption"),
        "CheckListItems": task.get("CheckListItems"),
        "Comments": comment_html or task.get("Comments") or "",
        "PassoAnterior": "Execucao",
        "PassoAtual": "RevisaoEValidacao",
        "ExecucaoInicio": task.get("ExecucaoInicio"),
        "ExecucaoFim": now,
        "ExecucaoPrazo": task.get("ExecucaoPrazo"),
        "ExecucaoMargem": task.get("ExecucaoMargem"),
        "RevisaoEValidacaoInicio": now,
        "RevisaoEValidacaoFim": task.get("RevisaoEValidacaoFim"),
        "RevisaoEValidacaoPrazo": task.get("RevisaoEValidacaoPrazo"),
        "RevisaoEValidacaoMargem": task.get("RevisaoEValidacaoMargem"),
        "PublicacaoInicio": task.get("PublicacaoInicio"),
        "PublicacaoFim": task.get("PublicacaoFim"),
        "PublicacaoPrazo": task.get("PublicacaoPrazo"),
        "PublicacaoMargem": task.get("PublicacaoMargem"),
        "FechamentoInicio": task.get("FechamentoInicio"),
        "FechamentoFim": task.get("FechamentoFim"),
        "FechamentoPrazo": task.get("FechamentoPrazo"),
        "FechamentoMargem": task.get("FechamentoMargem"),
        "Prazo": task.get("Prazo"),
        "PendenciaCliente": task.get("PendenciaCliente"),
        "MotivoDaPendencia": task.get("MotivoDaPendencia"),
        "PrazoPendencia": task.get("PrazoPendencia"),
        "MotivoDaPendenciaHistorico": task.get("MotivoDaPendenciaHistorico"),
        "NotificacaoCliente": task.get("NotificacaoCliente") or "",
        "ExecucaoUserAprovId": execucao_usuario_atual_id,
        "UsuarioAtualId": usuario_atual_id or execucao_usuario_atual_id,
        "Vencimento": task.get("RevisaoEValidacaoPrazo") or task.get("Vencimento"),
        "VencimentoMargem": task.get("RevisaoEValidacaoMargem") or task.get("VencimentoMargem"),
    }
    return payload


def build_create_payload(args: argparse.Namespace, usuario_atual_id: int | None = None) -> dict[str, Any]:
    resolved_usuario_atual_id = usuario_atual_id or args.usuario_atual_id or 300349
    return {
        "Title": "Execução",
        "Action": "Aprovar",
        "AreaId": str(args.area_id),
        "ProcessoId": str(args.processo_id),
        "ClienteId": str(args.cliente_id),
        "Competencia": args.competencia,
        "Descricao": args.descricao,
        "Prazo": args.prazo,
        "Origem": args.origem,
        "PassoAtual": "Execucao",
        "UsuarioAtualId": resolved_usuario_atual_id,
        "ExecucaoInicio": args.execucao_inicio,
        "ExecucaoPrazo": args.execucao_prazo,
        "ExecucaoMargem": args.execucao_margem,
        "Vencimento": args.vencimento or args.execucao_prazo,
        "VencimentoMargem": args.vencimento_margem or args.execucao_margem,
        "CheckListOption": args.checklist_option,
        "CheckListItems": args.checklist_items,
        "NotificacaoCliente": args.notificacao_cliente or "",
    }


def build_create_payload_from_request(row: dict[str, Any], usuario_atual_id: int | None = None) -> dict[str, Any]:
    resolved_usuario_atual_id = usuario_atual_id or row.get("usuario_atual_id") or 300349
    return {
        "Title": "Execução",
        "Action": "Aprovar",
        "AreaId": str(row["area_id"]),
        "ProcessoId": str(row["processo_id"]),
        "ClienteId": str(row["cliente_id"]),
        "Competencia": row["competencia"],
        "Descricao": row["descricao"],
        "Prazo": row["prazo"],
        "Origem": row.get("origem") or "Manual",
        "PassoAtual": "Execucao",
        "UsuarioAtualId": resolved_usuario_atual_id,
        "ExecucaoInicio": row["execucao_inicio"],
        "ExecucaoPrazo": row["execucao_prazo"],
        "ExecucaoMargem": row["execucao_margem"],
        "Vencimento": row.get("vencimento") or row["execucao_prazo"],
        "VencimentoMargem": row.get("vencimento_margem") or row["execucao_margem"],
        "CheckListOption": row.get("checklist_option") or False,
        "CheckListItems": json.dumps(row.get("checklist_items") or [], ensure_ascii=False),
        "NotificacaoCliente": row.get("notificacao_cliente") or "",
    }


def derive_client_site_url(task: dict[str, Any], task_site_url: str) -> str:
    tenant_root = task_site_url.split("/sites/", 1)[0]
    client_title = task["Cliente"]["Title"].strip()
    if not client_title:
        raise RuntimeError("A tarefa não retornou o nome do cliente.")
    slug = client_title.replace(" ", "")
    return f"{tenant_root}/sites/Empresa/{slug}"


def attach_document_to_task(args: argparse.Namespace) -> None:
    settings = SharePointSettings.from_env()
    client = SharePointClient(settings)

    file_path = Path(args.file).expanduser().resolve()
    if not file_path.is_file():
        raise RuntimeError(f"Arquivo não encontrado: {file_path}")

    task = fetch_task(client, args.task_id)
    client_site_url = (args.client_site_url or derive_client_site_url(task, settings.task_site_url)).rstrip("/")
    library_title = args.library_title or task["Cliente"]["BibliotecaPublicacao"]
    if not library_title:
        raise RuntimeError("A tarefa não informou BibliotecaPublicacao e nenhuma biblioteca foi passada manualmente.")

    library = resolve_library(client, client_site_url, library_title)
    folder_server_relative_url = library["RootFolder"]["ServerRelativeUrl"]
    publication_folder = args.folder or task["Cliente"].get("PastaPublicacao")
    if publication_folder:
        folder_server_relative_url = f"{folder_server_relative_url.rstrip('/')}/{publication_folder.strip('/')}"

    upload_result = client.upload_file(client_site_url, folder_server_relative_url, file_path)
    file_server_relative_url = upload_result["ServerRelativeUrl"]
    uploaded_item = fetch_uploaded_item(client, client_site_url, file_server_relative_url)

    metadata = {
        "ItemParentID": task["ID"],
        "Processo": task["Processo"]["Title"] if task.get("Processo") else None,
        "Area": task["Area"]["Title"] if task.get("Area") else None,
        "Competencia": task.get("Competencia"),
    }
    metadata = {key: value for key, value in metadata.items() if value not in (None, "")}

    update_file_metadata(
        client,
        client_site_url,
        file_server_relative_url,
        uploaded_item["__metadata"]["type"],
        metadata,
    )

    result = {
        "task_id": task["ID"],
        "task_title": task["Title"],
        "client_site_url": client_site_url,
        "library_id": library["Id"],
        "library_title": library["Title"],
        "library_root_folder": library["RootFolder"]["ServerRelativeUrl"],
        "uploaded_file_server_relative_url": file_server_relative_url,
        "uploaded_item_id": uploaded_item["ID"],
        "metadata_applied": metadata,
    }
    print(json.dumps(result, indent=2, ensure_ascii=False))


def finalize_task(args: argparse.Namespace) -> None:
    settings = SharePointSettings.from_env()
    client = SharePointClient(settings)
    task = fetch_task_for_finalize(client, args.task_id)
    payload = build_finalize_payload(
        task,
        args.comment_html,
        usuario_atual_id=resolve_stage_assignee_id(client, "revisao"),
    )
    update_task_item(client, args.task_id, task["__metadata"]["type"], payload)
    print(
        json.dumps(
            {
                "task_id": args.task_id,
                "from_stage": "Execucao",
                "to_stage": "RevisaoEValidacao",
                "action": payload["Action"],
                "title": payload["Title"],
                "execucao_fim": payload["ExecucaoFim"],
                "revisao_inicio": payload["RevisaoEValidacaoInicio"],
                "usuario_atual_id": payload["UsuarioAtualId"],
            },
            indent=2,
            ensure_ascii=False,
        )
    )


def create_task(args: argparse.Namespace) -> None:
    settings = SharePointSettings.from_env()
    client = SharePointClient(settings)
    payload = build_create_payload(
        args,
        usuario_atual_id=resolve_stage_assignee_id(client, "execucao", args.usuario_atual_id),
    )
    created = create_task_item(client, payload)
    print(
        json.dumps(
            {
                "task_id": created["ID"],
                "title": created["Title"],
                "passo_atual": created.get("PassoAtual"),
                "cliente_id": payload["ClienteId"],
                "processo_id": payload["ProcessoId"],
                "area_id": payload["AreaId"],
                "competencia": payload["Competencia"],
                "prazo": payload["Prazo"],
            },
            indent=2,
            ensure_ascii=False,
        )
    )


def fetch_pending_task_requests(
    db: PostgresClient,
    limit: int,
    routine_key: str | None = None,
    period_reference: str | None = None,
) -> list[dict[str, Any]]:
    filters = ["status = 'pending'"]
    if routine_key:
        filters.append(f"routine_key = {sql_literal(routine_key)}")
    if period_reference:
        filters.append(f"period_reference = {sql_literal(period_reference)}")
    where_clause = " and ".join(filters)
    return db.query_json(
        f"""
        select coalesce(jsonb_agg(to_jsonb(t) order by t.created_at), '[]'::jsonb)
        from (
          select *
          from sharepoint_task_requests
          where {where_clause}
          order by created_at
          limit {int(limit)}
        ) t
        """
    )


def fiscal_request_schedule(routine_key: str, prazo_iso: str, regime: str) -> dict[str, str]:
    """Datas por cliente das rotinas fiscais (regra das tarefas geradas até 08/2026).

    Rotina Mensal: execução termina 4 dias antes do prazo; Simples Nacional tem
    prazo dia 15 e margem dia 17. Demais rotinas: tudo no prazo.
    """
    prazo = parse_iso_datetime(prazo_iso)
    if routine_key == "fiscal_rotina_mensal":
        margem = prazo
        if "simples" in regime.casefold():
            prazo = prazo.replace(day=FISCAL_ROTINA_MENSAL_SIMPLES_DEADLINE_DAY)
            margem = prazo.replace(day=FISCAL_ROTINA_MENSAL_SIMPLES_VENCIMENTO_MARGEM_DAY)
        execucao_prazo = prazo - timedelta(days=4)
    else:
        execucao_prazo = margem = prazo
    return {
        "prazo": iso_z(prazo),
        "execucao_prazo": iso_z(execucao_prazo),
        "execucao_margem": iso_z(margem),
        "vencimento": iso_z(execucao_prazo),
        "vencimento_margem": iso_z(margem),
    }


def enqueue_fiscal_workbook_routine(
    args: argparse.Namespace,
    db: PostgresClient,
    run_date: date,
    values: dict[str, str],
    requested_by: str,
) -> None:
    routine_key = args.routine_key
    definition = monthly_routine_definition(routine_key)
    processes_by_title = None
    if definition.processo_id is None:
        client = SharePointClient(SharePointSettings.from_env())
        processes_by_title = {
            str(row.get("Title") or "").strip(): row for row in fetch_all_processes_index(client)
        }
    processo_id = resolve_process_id_for_routine(definition, run_date, processes_by_title)

    clients = db.query_json(
        f"""
        select coalesce(jsonb_agg(to_jsonb(t) order by t.cliente_nome), '[]'::jsonb)
        from (
          select cliente_id, cliente_nome, area_id, notes
          from sharepoint_monthly_task_clients
          where routine_key = {sql_literal(routine_key)}
            and active = true
        ) t
        """
    )
    rows = []
    for item in clients:
        notes = item.get("notes") or {}
        if isinstance(notes, str):
            notes = json.loads(notes) if notes.strip() else {}
        schedule = fiscal_request_schedule(routine_key, values["prazo"], str(notes.get("regime") or ""))
        rows.append(
            {
                "cliente_id": int(item["cliente_id"]),
                "cliente_nome": item.get("cliente_nome"),
                "area_id": int(item["area_id"]),
                **schedule,
            }
        )

    summary_base = {
        "routine_key": routine_key,
        "period_reference": values["period_reference"],
        "competencia": values["competencia"],
        "processo_id": processo_id,
        "processo_title": routine_process_title(definition, run_date),
        "execucao_inicio": values["execucao_inicio"],
        "clients": len(rows),
        "prazos": {
            prazo: sum(1 for row in rows if row["prazo"] == prazo)
            for prazo in sorted({row["prazo"] for row in rows})
        },
    }
    if args.dry_run:
        print(json.dumps({"dry_run": True, **summary_base, "sample": rows[:5]}, indent=2, ensure_ascii=False))
        return
    if not rows:
        print(json.dumps({**summary_base, "inserted": 0, "request_ids": []}, indent=2, ensure_ascii=False))
        return

    values_sql = ",\n".join(
        "("
        + ", ".join(
            [
                sql_literal(routine_key),
                sql_literal(values["period_reference"]),
                str(row["area_id"]),
                str(processo_id),
                str(row["cliente_id"]),
                sql_literal(values["competencia"]),
                f"{sql_literal(row['prazo'])}::timestamptz",
                "''",
                f"{sql_literal(values['execucao_inicio'])}::timestamptz",
                f"{sql_literal(row['execucao_prazo'])}::timestamptz",
                f"{sql_literal(row['execucao_margem'])}::timestamptz",
                f"{sql_literal(row['vencimento'])}::timestamptz",
                f"{sql_literal(row['vencimento_margem'])}::timestamptz",
                sql_literal(requested_by),
            ]
        )
        + ")"
        for row in rows
    )
    result = db.query_json(
        f"""
        with inserted as (
          insert into sharepoint_task_requests (
            routine_key, period_reference, area_id, processo_id, cliente_id, competencia,
            prazo, descricao, execucao_inicio, execucao_prazo, execucao_margem,
            vencimento, vencimento_margem, requested_by
          )
          values
          {values_sql}
          on conflict do nothing
          returning request_id
        )
        select jsonb_build_object(
          'inserted', count(*),
          'request_ids', coalesce(jsonb_agg(request_id order by request_id), '[]'::jsonb)
        )
        from inserted
        """
    )
    summary = result if isinstance(result, dict) else {"inserted": 0, "request_ids": []}
    print(json.dumps({**summary_base, **summary}, indent=2, ensure_ascii=False))


def enqueue_monthly_routine(args: argparse.Namespace) -> None:
    db = PostgresClient(PostgresSettings.from_env())
    run_date = parse_run_date(args.run_date)
    routine_key = args.routine_key
    values = monthly_routine_dates(routine_key, run_date)
    requested_by = args.requested_by or f"n8n:{routine_key}:{values['period_reference']}"

    if routine_key in FISCAL_WORKBOOK_ROUTINES:
        enqueue_fiscal_workbook_routine(args, db, run_date, values, requested_by)
        return

    # Prazo fixo da rotina, ou calculado pelo dia de cada cliente (deadline_day) quando due_day_source = "client".
    extra_columns = ""
    extra_values = ""
    if "prazo_mes" in values:
        definition = monthly_routine_definition(routine_key)
        missing_day = db.query_json(
            f"""
            select coalesce(jsonb_agg(cliente_id order by cliente_id), '[]'::jsonb)
            from sharepoint_monthly_task_clients
            where routine_key = {sql_literal(routine_key)}
              and active = true
              and deadline_day is null
            """
        )
        if missing_day:
            raise RuntimeError(f"Clientes ativos sem deadline_day na rotina {routine_key}: {missing_day}")
        deadline_days = db.query_json(
            f"""
            select coalesce(jsonb_agg(distinct deadline_day), '[]'::jsonb)
            from sharepoint_monthly_task_clients
            where routine_key = {sql_literal(routine_key)}
              and active = true
            """
        )
        schedules = {int(day): client_day_schedule(definition, values["prazo_mes"], int(day)) for day in deadline_days}

        def by_deadline_day(field: str) -> str:
            whens = " ".join(
                f"when {day} then {sql_literal(schedule[field])}::timestamptz" for day, schedule in sorted(schedules.items())
            )
            return f"case deadline_day {whens} end"

        prazo_sql = by_deadline_day("prazo")
        execucao_prazo_sql = by_deadline_day("execucao_prazo")
        execucao_margem_sql = by_deadline_day("execucao_margem")
        extra_columns = ",\n            vencimento,\n            vencimento_margem"
        extra_values = f",\n            {by_deadline_day('vencimento')},\n            {by_deadline_day('vencimento_margem')}"
    else:
        prazo_sql = execucao_prazo_sql = execucao_margem_sql = f"{sql_literal(values['prazo'])}::timestamptz"

    if args.dry_run:
        rows = db.query_json(
            f"""
            select coalesce(
              jsonb_agg(
                jsonb_build_object(
                  'cliente_id', cliente_id,
                  'cliente_nome', cliente_nome,
                  'area_id', area_id,
                  'processo_id', processo_id,
                  'prazo', {prazo_sql}
                )
                order by cliente_nome
              ),
              '[]'::jsonb
            )
            from sharepoint_monthly_task_clients
            where routine_key = {sql_literal(routine_key)}
              and active = true
            """
        )
        print(
            json.dumps(
                {
                    "dry_run": True,
                    "routine_key": routine_key,
                    "period_reference": values["period_reference"],
                    "competencia": values["competencia"],
                    "prazo": values.get("prazo") or "por cliente",
                    "execucao_inicio": values["execucao_inicio"],
                    "clients": len(rows),
                    "sample": rows[:10],
                },
                indent=2,
                ensure_ascii=False,
            )
        )
        return

    result = db.query_json(
        f"""
        with inserted as (
          insert into sharepoint_task_requests (
            routine_key,
            period_reference,
            area_id,
            processo_id,
            cliente_id,
            competencia,
            prazo,
            descricao,
            execucao_inicio,
            execucao_prazo,
            execucao_margem,
            requested_by{extra_columns}
          )
          select
            {sql_literal(routine_key)},
            {sql_literal(values['period_reference'])},
            area_id,
            processo_id,
            cliente_id,
            {sql_literal(values['competencia'])},
            {prazo_sql},
            '',
            {sql_literal(values['execucao_inicio'])}::timestamptz,
            {execucao_prazo_sql},
            {execucao_margem_sql},
            {sql_literal(requested_by)}{extra_values}
          from sharepoint_monthly_task_clients
          where routine_key = {sql_literal(routine_key)}
            and active = true
          on conflict do nothing
          returning request_id, cliente_id
        )
        select jsonb_build_object(
          'inserted', count(*),
          'request_ids', coalesce(jsonb_agg(request_id order by request_id), '[]'::jsonb)
        )
        from inserted
        """
    )
    summary = result if isinstance(result, dict) else {"inserted": 0, "request_ids": []}
    print(
        json.dumps(
            {
                "routine_key": routine_key,
                "period_reference": values["period_reference"],
                "competencia": values["competencia"],
                "prazo": values.get("prazo") or "por cliente",
                "execucao_inicio": values["execucao_inicio"],
                **summary,
            },
            indent=2,
            ensure_ascii=False,
        )
    )


def mark_task_request_running(db: PostgresClient, request_id: str) -> None:
    db.execute(
        f"""
        update sharepoint_task_requests
        set status = 'running',
            started_at = coalesce(started_at, now()),
            updated_at = now(),
            error_message = null
        where request_id = {sql_literal(request_id)}
          and status = 'pending'
        """
    )


def mark_task_request_created(
    db: PostgresClient,
    request_id: str,
    sharepoint_task_id: int,
    response: dict[str, Any],
) -> None:
    db.execute(
        f"""
        update sharepoint_task_requests
        set status = 'created',
            sharepoint_task_id = {int(sharepoint_task_id)},
            sharepoint_response = {jsonb_literal(response)},
            finished_at = now(),
            updated_at = now(),
            error_message = null
        where request_id = {sql_literal(request_id)}
        """
    )


def mark_task_request_failed(db: PostgresClient, request_id: str, error: str) -> None:
    db.execute(
        f"""
        update sharepoint_task_requests
        set status = 'failed',
            error_message = {sql_literal(error[:4000])},
            finished_at = now(),
            updated_at = now()
        where request_id = {sql_literal(request_id)}
        """
    )


def process_pending_creates(args: argparse.Namespace) -> None:
    settings = SharePointSettings.from_env()
    client = SharePointClient(settings)
    db = PostgresClient(PostgresSettings.from_env())
    rows = fetch_pending_task_requests(db, args.limit, args.routine_key, args.period_reference)
    results = []
    fixed_execucao_assignee_id = resolve_stage_assignee_id(client, "execucao")

    for row in rows:
        request_id = row["request_id"]
        try:
            usuario_atual = None
            usuario_atual_id = fixed_execucao_assignee_id or row.get("usuario_atual_id")
            if usuario_atual_id is None and args.resolve_responsavel_folha:
                routine = MONTHLY_ROUTINES.get(row.get("routine_key") or "")
                if routine is not None and routine.assignee_field_name == "ResponsavelFiscal":
                    usuario_atual = fetch_client_responsavel_fiscal(client, int(row["cliente_id"]))
                else:
                    usuario_atual = fetch_client_responsavel_folha(client, int(row["cliente_id"]))
                usuario_atual_id = usuario_atual["id"]

            if args.dry_run:
                results.append(
                    {
                        "request_id": request_id,
                        "status": "dry_run",
                        "usuario_atual_id": usuario_atual_id,
                        "responsavel_folha": usuario_atual,
                        "payload": build_create_payload_from_request(
                            row,
                            usuario_atual_id,
                        ),
                    }
                )
                continue

            mark_task_request_running(db, request_id)
            payload = build_create_payload_from_request(row, usuario_atual_id)
            created = create_task_item(client, payload)
            response = {
                "task_id": created["ID"],
                "title": created["Title"],
                "passo_atual": created.get("PassoAtual"),
                "cliente_id": payload["ClienteId"],
                "processo_id": payload["ProcessoId"],
                "area_id": payload["AreaId"],
                "competencia": payload["Competencia"],
                "prazo": payload["Prazo"],
            }
            mark_task_request_created(db, request_id, created["ID"], response)
            results.append({"request_id": request_id, "status": "created", **response})
        except Exception as exc:
            mark_task_request_failed(db, request_id, str(exc))
            results.append({"request_id": request_id, "status": "failed", "error": str(exc)})

    print(json.dumps({"processed": len(results), "results": results}, indent=2, ensure_ascii=False))


def routine_process_title(definition: MonthlyRoutineDefinition, run_date: date) -> str | None:
    if definition.process_title_template is None:
        return None
    process_month_offset = (
        definition.process_month_offset
        if definition.process_month_offset is not None
        else definition.competencia_month_offset
    )
    process_month = add_months(run_date, process_month_offset)
    return definition.process_title_template.format(
        month_name_upper=month_name_upper(process_month.month),
        year=process_month.year,
    )


def resolve_process_id_for_routine(
    definition: MonthlyRoutineDefinition,
    run_date: date,
    processes_by_title: dict[str, dict[str, Any]] | None = None,
) -> int:
    if definition.processo_id is not None:
        return int(definition.processo_id)
    process_title = routine_process_title(definition, run_date)
    if not process_title:
        raise RuntimeError(f"Rotina {definition.routine_key} não possui processo configurado.")
    if not processes_by_title or process_title not in processes_by_title:
        raise RuntimeError(f"Processo não encontrado no SharePoint: {process_title}")
    return int(processes_by_title[process_title]["ID"])


def parse_iso_datetime(value: str) -> datetime:
    normalized = value.strip()
    if normalized.endswith("Z"):
        normalized = normalized[:-1] + "+00:00"
    return datetime.fromisoformat(normalized)


def format_pt_br_date(value: datetime) -> str:
    return value.strftime("%d/%m/%Y")


def fiscal_rotina_mensal_pendencia_payload(execucao_inicio: str, execucao_prazo: str) -> dict[str, Any]:
    execucao_inicio_dt = parse_iso_datetime(execucao_inicio)
    prazo_pendencia_dt = parse_iso_datetime(execucao_prazo) - timedelta(days=1)
    pendencia_text = "Aguardando documentação do cliente (pendência automática)"
    checklist_html = "".join(f"<li>{item}</li>" for item in FISCAL_ROTINA_MENSAL_PENDENCIA_ITEMS)
    pendencia_historico = (
        "<div class='historic-pendencia-container'>"
        "<div class='historic-pendencia-box'>"
        f"<div class='historic-pendencia-date'>{format_pt_br_date(execucao_inicio_dt)}</div>"
        "<div class='historic-pendencia-text'>"
        f"<P>{pendencia_text}</P>"
        f"<P><div><div><ul>{checklist_html}</ul></div></div></P> "
        f"Prazo: {format_pt_br_date(prazo_pendencia_dt)}"
        "</div></div></div>"
    )
    pendencia_resumo = "\n".join(
        [
            pendencia_text,
            "",
            *FISCAL_ROTINA_MENSAL_PENDENCIA_ITEMS,
            "",
            f"Prazo: {format_pt_br_date(prazo_pendencia_dt)}",
        ]
    )
    return {
        "PendenciaCliente": True,
        "MotivoDaPendencia": pendencia_resumo,
        "PrazoPendencia": iso_z(prazo_pendencia_dt),
        "MotivoDaPendenciaHistorico": pendencia_historico,
    }


def fiscal_rotina_mensal_pendencia_activation_payload(execucao_prazo: str) -> dict[str, Any]:
    execucao_prazo_dt = parse_iso_datetime(execucao_prazo)
    activation_start = datetime(execucao_prazo_dt.year, execucao_prazo_dt.month, 1, tzinfo=timezone.utc)
    return fiscal_rotina_mensal_pendencia_payload(
        iso_z(activation_start),
        execucao_prazo,
    )


def fiscal_rotina_mensal_pendencia_deactivation_payload() -> dict[str, Any]:
    return {
        "PendenciaCliente": False,
        "MotivoDaPendencia": "",
        "PrazoPendencia": None,
        "MotivoDaPendenciaHistorico": "",
        "NotificacaoCliente": "",
    }


def active_flag_from_sheet(value: str) -> bool:
    normalized = str(value or "").strip().casefold()
    return normalized in {"1", "s", "sim", "true", "x"}


def routine_key_for_sheet(sheet_name: str) -> str:
    for routine_key in FISCAL_WORKBOOK_ROUTINES:
        definition = monthly_routine_definition(routine_key)
        if definition.sync_sheet_name == sheet_name:
            return routine_key
    available = ", ".join(
        routine_definition.sync_sheet_name or ""
        for routine_definition in map(monthly_routine_definition, FISCAL_WORKBOOK_ROUTINES)
    )
    raise RuntimeError(f"Aba desconhecida no workbook fiscal: {sheet_name}. Esperadas: {available}")


def sync_fiscal_workbook(args: argparse.Namespace) -> None:
    workbook_path = Path(args.workbook).expanduser().resolve()
    if not workbook_path.is_file():
        raise RuntimeError(f"Workbook fiscal não encontrado: {workbook_path}")

    settings = SharePointSettings.from_env()
    client = SharePointClient(settings)
    db = PostgresClient(PostgresSettings.from_env())

    workbook_rows = load_workbook_rows(workbook_path)
    clients = fetch_all_clients_index(client)
    clients_by_cnpj = {normalize_cnpj(row.get("CNPJ")): row for row in clients if normalize_cnpj(row.get("CNPJ"))}
    clients_by_title = {normalize_client_title(row.get("Title")): row for row in clients}

    results: list[dict[str, Any]] = []
    statements: list[str] = ["begin;"]
    for routine_key in FISCAL_WORKBOOK_ROUTINES:
        definition = monthly_routine_definition(routine_key)
        sheet_name = definition.sync_sheet_name or ""
        if sheet_name not in workbook_rows:
            raise RuntimeError(f"Aba obrigatória ausente no workbook fiscal: {sheet_name}")
        statements.append(f"delete from sharepoint_monthly_task_clients where routine_key = {sql_literal(routine_key)};")
        rows_by_client_id: dict[int, dict[str, Any]] = {}
        for row in workbook_rows[sheet_name]:
            cnpj = normalize_cnpj(row.get("CNPJ"))
            title = row.get("Título", "").strip()
            resolved = clients_by_cnpj.get(cnpj) if cnpj else None
            if resolved is None:
                resolved = clients_by_title.get(normalize_client_title(title))
            if resolved is None:
                raise RuntimeError(
                    f"Cliente não encontrado no SharePoint para a aba '{sheet_name}': "
                    f"Título='{title}' CNPJ='{row.get('CNPJ', '').strip()}'"
                )
            highlight = row.get("_highlight")
            if highlight == "vermelho":
                active = False
            elif highlight in ("amarelo", "azul"):
                # Amarelo e azul = empresas novas: sempre geram tarefa.
                active = True
            else:
                active = active_flag_from_sheet(row.get("Cliente Ativo (S/N)", ""))
            process_id = definition.processo_id or 0
            cliente_id = int(resolved["ID"])
            regime = row.get("Regime", "").strip()
            if routine_key == "fiscal_pis_cofins_mes_atual" and "simples" in regime.casefold():
                # Simples Nacional não apura PIS/COFINS separadamente — nunca gerar tarefa
                # dessa rotina para esses clientes, mesmo que a planilha marque como ativo.
                active = False
            rows_by_client_id[cliente_id] = {
                "cliente_id": cliente_id,
                "cliente_nome": str(resolved.get("Title") or "").strip(),
                "active": active,
                "highlight": highlight,
                "processo_id_seed": int(process_id),
                "notes": {
                    "source": "fiscal_workbook",
                    "workbook_path": str(workbook_path),
                    "sheet_name": sheet_name,
                    "cnpj": cnpj,
                    "regime": regime,
                    "colaborador_fiscal_sheet": row.get("Colaborador Fiscal", "").strip(),
                    "highlight": highlight,
                },
            }
        sheet_results = list(rows_by_client_id.values())
        for item in sheet_results:
            notes = item["notes"]
            statements.append(
                "insert into sharepoint_monthly_task_clients "
                "(routine_key, cliente_id, cliente_nome, area_id, processo_id, active, notes) values "
                f"({sql_literal(routine_key)}, {item['cliente_id']}, {sql_literal(item['cliente_nome'])}, "
                f"{definition.area_id}, {item['processo_id_seed']}, {'true' if item['active'] else 'false'}, "
                f"{sql_literal(json.dumps(notes, ensure_ascii=False))});"
            )
        removed = [item for item in sheet_results if item["highlight"] == "vermelho"]
        added = [item for item in sheet_results if item["highlight"] == "amarelo"]
        added_blue = [item for item in sheet_results if item["highlight"] == "azul"]
        results.append(
            {
                "routine_key": routine_key,
                "sheet_name": sheet_name,
                "rows": len(sheet_results),
                "active": sum(1 for item in sheet_results if item["active"]),
                "removidos_vermelho": [item["cliente_nome"] for item in removed],
                "novos_amarelo": [item["cliente_nome"] for item in added],
                "novos_azul": [item["cliente_nome"] for item in added_blue],
            }
        )
    statements.append("commit;")

    if args.dry_run:
        print(json.dumps({"dry_run": True, "workbook": str(workbook_path), "results": results}, indent=2, ensure_ascii=False))
        return

    db.execute("\n".join(statements))
    print(json.dumps({"workbook": str(workbook_path), "results": results}, indent=2, ensure_ascii=False))


def fiscal_rotina_mensal_pending_competencia(run_date: date) -> str:
    competencia_base = add_months(run_date, -1)
    return f"{MONTH_NAMES_PT[competencia_base.month]} {competencia_base.year}"


def set_fiscal_rotina_mensal_pendencia(args: argparse.Namespace) -> None:
    run_date = parse_run_date(args.run_date)
    competencia = args.competencia or fiscal_rotina_mensal_pending_competencia(run_date)
    settings = SharePointSettings.from_env()
    client = SharePointClient(settings)
    tasks = fetch_tasks_for_fiscal_rotina_mensal_pendencia(client, competencia, passo_atual=args.passo_atual)

    skipped = []
    if args.mode == "activate":
        flags = fetch_clients_ativar_pendencia_fiscal_flags(client, [task["ClienteId"] for task in tasks])
        eligible_tasks = []
        for task in tasks:
            if flags.get(int(task["ClienteId"]), False):
                eligible_tasks.append(task)
            else:
                skipped.append(
                    {
                        "task_id": int(task["ID"]),
                        "cliente_id": int(task["ClienteId"]),
                        "reason": "cliente com AtivarSistemaPendenciaFiscal desativado",
                    }
                )
    else:
        eligible_tasks = tasks

    results = []
    for task in eligible_tasks:
        task_id = int(task["ID"])
        if args.mode == "activate":
            payload = fiscal_rotina_mensal_pendencia_activation_payload(str(task["ExecucaoPrazo"]))
        else:
            payload = fiscal_rotina_mensal_pendencia_deactivation_payload()
        update_task_item(client, task_id, "SP.Data.TarefasListItem", payload)
        results.append(
            {
                "task_id": task_id,
                "competencia": task.get("Competencia"),
                "passo_atual": task.get("PassoAtual"),
                "mode": args.mode,
                "pendencia_cliente": payload["PendenciaCliente"],
                "prazo_pendencia": payload.get("PrazoPendencia"),
            }
        )

    print(
        json.dumps(
            {
                "mode": args.mode,
                "competencia": competencia,
                "passo_atual": args.passo_atual,
                "updated": len(results),
                "skipped": len(skipped),
                "results": results,
                "skipped_details": skipped,
            },
            indent=2,
            ensure_ascii=False,
        )
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Opera anexos e finalização de tarefas no portal via SharePoint REST.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    create_parser = subparsers.add_parser("create", help="Cria uma nova tarefa no estágio Execucao.")
    create_parser.add_argument("--area-id", type=int, required=True, help="ID da área.")
    create_parser.add_argument("--processo-id", type=int, required=True, help="ID do processo.")
    create_parser.add_argument("--cliente-id", type=int, required=True, help="ID do cliente.")
    create_parser.add_argument("--competencia", required=True, help="Competência da tarefa.")
    create_parser.add_argument("--prazo", required=True, help="Prazo final informado na criação. Ex.: 2026-04-26T00:00:00")
    create_parser.add_argument("--descricao", required=True, help="Descrição da tarefa.")
    create_parser.add_argument("--execucao-inicio", required=True, help="Data/hora ISO de início da execução.")
    create_parser.add_argument("--execucao-prazo", required=True, help="Prazo ISO calculado para a execução.")
    create_parser.add_argument("--execucao-margem", required=True, help="Margem ISO calculada para a execução.")
    create_parser.add_argument("--vencimento", help="Vencimento ISO da tarefa. Se omitido, usa execucao-prazo.")
    create_parser.add_argument("--vencimento-margem", help="Margem ISO da tarefa. Se omitido, usa execucao-margem.")
    create_parser.add_argument("--usuario-atual-id", type=int, help="ID do usuário atual.")
    create_parser.add_argument("--origem", default="Manual", help="Origem da tarefa. Padrão: Manual.")
    create_parser.add_argument("--checklist-option", action="store_true", help="Ativa CheckListOption.")
    create_parser.add_argument("--checklist-items", default="[]", help="JSON serializado do checklist. Padrão: [].")
    create_parser.add_argument("--notificacao-cliente", default="", help="Valor inicial de NotificacaoCliente.")
    create_parser.set_defaults(handler=create_task)

    monthly_parser = subparsers.add_parser(
        "enqueue-monthly-routine",
        help="Enfileira tarefas mensais a partir da tabela sharepoint_monthly_task_clients.",
    )
    monthly_parser.add_argument("--routine-key", default="folha_mensal", help="Chave da rotina mensal.")
    monthly_parser.add_argument("--run-date", help="Data base YYYY-MM-DD. Se omitida, usa a data UTC atual.")
    monthly_parser.add_argument("--requested-by", help="Identificador de origem gravado na fila.")
    monthly_parser.add_argument("--dry-run", action="store_true", help="Mostra o lote calculado sem inserir filas.")
    monthly_parser.set_defaults(handler=enqueue_monthly_routine)

    fiscal_sync_parser = subparsers.add_parser(
        "sync-fiscal-workbook",
        help="Sincroniza as abas do workbook fiscal para sharepoint_monthly_task_clients.",
    )
    fiscal_sync_parser.add_argument("--workbook", required=True, help="Caminho do arquivo .xlsx com as abas fiscais.")
    fiscal_sync_parser.add_argument("--dry-run", action="store_true", help="Valida e mostra a carga sem gravar no Postgres.")
    fiscal_sync_parser.set_defaults(handler=sync_fiscal_workbook)

    describe_monthly_parser = subparsers.add_parser(
        "describe-monthly-routine",
        help="Calcula competencia, period_reference e prazo de uma rotina mensal sem criar filas.",
    )
    describe_monthly_parser.add_argument("--routine-key", default="folha_mensal", help="Chave da rotina mensal.")
    describe_monthly_parser.add_argument("--run-date", help="Data base YYYY-MM-DD. Se omitida, usa a data UTC atual.")
    describe_monthly_parser.set_defaults(
        handler=lambda args: print(
            json.dumps(
                {
                    "routine_key": args.routine_key,
                    **monthly_routine_dates(args.routine_key, parse_run_date(args.run_date)),
                },
                indent=2,
                ensure_ascii=False,
            )
        )
    )

    process_parser = subparsers.add_parser(
        "process-pending-creates",
        help="Processa tarefas pendentes da tabela sharepoint_task_requests.",
    )
    process_parser.add_argument("--limit", type=int, default=10, help="Quantidade máxima de pendências a processar.")
    process_parser.add_argument("--routine-key", help="Processa apenas pendências desta rotina.")
    process_parser.add_argument("--period-reference", help="Processa apenas pendências desta competência YYYY-MM.")
    process_parser.add_argument("--dry-run", action="store_true", help="Mostra os payloads sem criar tarefas nem atualizar o banco.")
    process_parser.add_argument(
        "--no-resolve-responsavel-folha",
        action="store_false",
        dest="resolve_responsavel_folha",
        help="Não resolve UsuarioAtualId pelo ResponsavelFolha do cliente.",
    )
    process_parser.set_defaults(resolve_responsavel_folha=True)
    process_parser.set_defaults(handler=process_pending_creates)

    fiscal_pendencia_parser = subparsers.add_parser(
        "set-fiscal-rotina-mensal-pendencia",
        help="Ativa ou desativa a pendência automática das tarefas Fiscal - Rotina Mensal Fiscal.",
    )
    fiscal_pendencia_parser.add_argument("--mode", choices=("activate", "deactivate"), required=True, help="Ação a aplicar nas tarefas encontradas.")
    fiscal_pendencia_parser.add_argument("--run-date", help="Data base YYYY-MM-DD. Se omitida, usa a data UTC atual.")
    fiscal_pendencia_parser.add_argument("--competencia", help="Competência textual alvo. Se omitida, deriva do run-date.")
    fiscal_pendencia_parser.add_argument("--passo-atual", default="Execucao", help="Filtra tarefas pelo PassoAtual. Padrão: Execucao.")
    fiscal_pendencia_parser.set_defaults(handler=set_fiscal_rotina_mensal_pendencia)

    attach_parser = subparsers.add_parser("attach", help="Anexa um arquivo em uma tarefa.")
    attach_parser.add_argument("--task-id", type=int, required=True, help="ID da tarefa na lista Tarefas.")
    attach_parser.add_argument("--file", required=True, help="Caminho local do arquivo a ser enviado.")
    attach_parser.add_argument(
        "--client-site-url",
        help="URL do site do cliente. Se omitido, tenta derivar a partir do nome do cliente retornado pela tarefa.",
    )
    attach_parser.add_argument("--library-title", help="Título da biblioteca de destino. Se omitido, usa BibliotecaPublicacao da tarefa.")
    attach_parser.add_argument("--folder", help="Subpasta opcional dentro da biblioteca.")
    attach_parser.set_defaults(handler=attach_document_to_task)

    finalize_parser = subparsers.add_parser("finalize", help="Finaliza a etapa Execucao e avança para RevisaoEValidacao.")
    finalize_parser.add_argument("--task-id", type=int, required=True, help="ID da tarefa na lista Tarefas.")
    finalize_parser.add_argument("--comment-html", help="HTML opcional a ser enviado em Comments.")
    finalize_parser.set_defaults(handler=finalize_task)
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        args.handler(args)
        return 0
    except Exception as exc:
        # Traceback completo para o painel de automações; a última linha continua "Erro: ...".
        traceback.print_exc(file=sys.stderr)
        print(f"Erro: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
