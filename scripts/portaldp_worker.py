#!/usr/bin/env python3
"""Worker PortalDP -> tarefa no portal de tarefas do SharePoint.

Roda no servidor que tem o certificado e o .env.sharepoint). A cada execução:
  1. pergunta ao PortalDP (API interna, token) quais pedidos estão pendentes;
  2. para cada um: cria a tarefa (ou localiza a mensal já existente), anexa os
     documentos enviados pelo agente DP ao concluir e avança para
     Revisão e Validação;
  3. devolve o resultado ao PortalDP (criada / erro / aviso).

Regras:
  - responsável da Execução = Clientes.ResponsavelFolha do cliente no portal;
  - responsável da Revisão   = payload["revisao_assignee"] (config da origem);
  - processo mensal (9): se já existe a tarefa da competência -> em Execução,
    reaproveita; além da Execução, NÃO cria nada e avisa o agente;
  - nome do arquivo no SharePoint recebe o nº do chamado, para dois chamados
    com "folha.pdf" não se sobrescreverem na biblioteca do cliente.

Uso:
  set -a; . ./.env.sharepoint; set +a
  export PORTALDP_INTERNAL_TOKEN=...   PORTALDP_URL=http://localhost:8001
  python3 scripts/portaldp_worker.py            # processa a fila
  python3 scripts/portaldp_worker.py --dry-run  # só mostra o que faria (não escreve no SharePoint, não reivindica pedidos)
"""
from __future__ import annotations

import argparse
import fcntl
import json
import os
import re
import sys
import tempfile
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

DEFAULT_PORTALDP_URL = "http://localhost:8001"
LOCK_PATH = "/tmp/portaldp_worker.lock"
STEP_LABELS = {
    "Execucao": "Execução",
    "RevisaoEValidacao": "Revisão e Validação",
    "Publicacao": "Publicação",
    "Fechamento": "Fechamento",
}


# --------------------------------------------------------------------------- regras puras
def iso_z(day: date) -> str:
    """03:00Z = meia-noite de Brasília, como o portal grava nas tarefas manuais."""
    return datetime(day.year, day.month, day.day, 3, tzinfo=timezone.utc).isoformat().replace("+00:00", "Z")


def task_dates(prazo: date) -> dict[str, str]:
    """Mesma regra das tarefas manuais do portal (medida em 30/09/2026) e do
    `client_day_schedule` do sharepoint_task_attachment.py: tudo deriva do prazo."""
    return {
        "Prazo": f"{prazo.isoformat()}T00:00:00",
        "ExecucaoPrazo": iso_z(prazo - timedelta(days=4)),
        "ExecucaoMargem": iso_z(prazo - timedelta(days=3)),
        "Vencimento": iso_z(prazo - timedelta(days=2)),
        "VencimentoMargem": iso_z(prazo),
    }


_INVALID_NAME_CHARS = re.compile(r'[\\/:*?"<>|#%~&{}\x00-\x1f]')


def safe_file_name(name: str) -> str:
    cleaned = _INVALID_NAME_CHARS.sub("_", Path(name).name).strip(" .")
    return cleaned or "arquivo"


def sharepoint_file_names(attachments: list[dict[str, Any]], ticket_id: int) -> dict[int, str]:
    """attachment id -> nome no SharePoint: '<nome> - PortalDP <chamado>.<ext>' (único dentro do pedido)."""
    names: dict[int, str] = {}
    used: set[str] = set()
    for attachment in attachments:
        original = safe_file_name(attachment["name"])
        stem, suffix = os.path.splitext(original)
        candidate = f"{stem} - PortalDP {ticket_id}{suffix}"
        counter = 2
        while candidate.casefold() in used:
            candidate = f"{stem} - PortalDP {ticket_id} ({counter}){suffix}"
            counter += 1
        used.add(candidate.casefold())
        names[attachment["id"]] = candidate
    return names


def build_create_payload(request_payload: dict[str, Any], responsavel_id: int, now: datetime | None = None) -> dict[str, Any]:
    now = now or datetime.now(timezone.utc)
    dates = task_dates(date.fromisoformat(request_payload["prazo"]))
    return {
        "Title": "Execução",
        "Action": "Aprovar",
        "AreaId": str(request_payload["area_id"]),
        "ProcessoId": str(request_payload["process_id"]),
        "ClienteId": str(request_payload["cliente_id"]),
        "Competencia": request_payload["competencia"],
        "Descricao": request_payload["descricao"],
        "Origem": "Manual",
        "PassoAtual": "Execucao",
        "UsuarioAtualId": responsavel_id,
        "ExecucaoInicio": now.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "CheckListOption": False,
        "CheckListItems": "[]",
        "NotificacaoCliente": "",
        **dates,
    }


# --------------------------------------------------------------------------- PortalDP (API interna)
class PortalDPApi:
    def __init__(self, base_url: str, token: str, timeout: int = 60) -> None:
        self.base_url = base_url.rstrip("/")
        self.token = token
        self.timeout = timeout

    def _request(self, method: str, path: str, body: dict[str, Any] | None = None) -> tuple[int, bytes]:
        data = json.dumps(body).encode("utf-8") if body is not None else (b"" if method == "POST" else None)
        headers = {"Authorization": f"Bearer {self.token}"}
        if body is not None:
            headers["Content-Type"] = "application/json"
        request = urllib.request.Request(f"{self.base_url}{path}", data=data, method=method, headers=headers)
        try:
            with urllib.request.urlopen(request, timeout=self.timeout) as response:
                return response.status, response.read()
        except urllib.error.HTTPError as exc:
            return exc.code, exc.read()

    def pending(self) -> list[dict[str, Any]]:
        status, body = self._request("GET", "/interno/sharepoint/pendentes")
        if status != 200:
            raise RuntimeError(f"PortalDP respondeu {status} ao listar pendentes (token/rede?)")
        return json.loads(body)["requests"]

    def claim(self, request_id: int) -> dict[str, Any] | None:
        status, body = self._request("POST", f"/interno/sharepoint/{request_id}/iniciar")
        if status == 409:
            return None
        if status != 200:
            raise RuntimeError(f"PortalDP respondeu {status} ao iniciar o pedido {request_id}")
        return json.loads(body)

    def download(self, request_id: int, attachment_id: int) -> bytes:
        status, body = self._request("GET", f"/interno/sharepoint/{request_id}/anexos/{attachment_id}")
        if status != 200:
            raise RuntimeError(f"PortalDP respondeu {status} ao baixar o anexo {attachment_id}")
        return body

    def result(self, request_id: int, status: str, task_id: int | None = None, message: str = "") -> None:
        code, body = self._request(
            "POST", f"/interno/sharepoint/{request_id}/resultado",
            {"status": status, "task_id": task_id, "message": message},
        )
        if code != 200:
            raise RuntimeError(f"PortalDP respondeu {code} ao gravar o resultado do pedido {request_id}: {body[:200]!r}")


# --------------------------------------------------------------------------- SharePoint
class SharePointOps:
    """Operações no portal de tarefas, reaproveitando sharepoint_task_attachment.py."""

    def __init__(self) -> None:
        import sharepoint_task_attachment as sta

        self.sta = sta
        self.client = sta.SharePointClient(sta.SharePointSettings.from_env())

    def find_monthly_task(self, process_id: int, cliente_id: int, competencia: str) -> dict[str, Any] | None:
        query = urllib.parse.urlencode({
            "$filter": f"ProcessoId eq {int(process_id)} and ClienteId eq {int(cliente_id)} and Competencia eq '{competencia}'",
            "$select": "Id,ID,PassoAtual,Competencia",
            "$orderby": "Id desc",
            "$top": "5",
        })
        response = self.client.rest_json(
            self.client.settings.task_site_url, f"/_api/web/lists/getbytitle('Tarefas')/items?{query}"
        )
        rows = response["d"]["results"]
        return rows[0] if rows else None

    def task_step(self, task_id: int) -> str:
        query = urllib.parse.urlencode({"$select": "Id,PassoAtual"})
        response = self.client.rest_json(
            self.client.settings.task_site_url, f"/_api/web/lists/getbytitle('Tarefas')/items({int(task_id)})?{query}"
        )
        return response["d"]["PassoAtual"]

    def responsavel_folha_id(self, cliente_id: int) -> int:
        return int(self.sta.fetch_client_responsavel_folha(self.client, int(cliente_id))["id"])

    def create_task(self, payload: dict[str, Any]) -> int:
        return int(self.sta.create_task_item(self.client, payload)["ID"])

    def attach(self, task_id: int, client_site_url: str, file_path: Path) -> str:
        sta = self.sta
        task = sta.fetch_task(self.client, task_id)
        library_title = task["Cliente"]["BibliotecaPublicacao"]
        if not library_title:
            raise RuntimeError("A tarefa não informou BibliotecaPublicacao para o cliente.")
        library = sta.resolve_library(self.client, client_site_url.rstrip("/"), library_title)
        folder = library["RootFolder"]["ServerRelativeUrl"]
        publication_folder = task["Cliente"].get("PastaPublicacao")
        if publication_folder:
            folder = f"{folder.rstrip('/')}/{publication_folder.strip('/')}"
        uploaded = self.client.upload_file(client_site_url.rstrip("/"), folder, file_path)
        file_url = uploaded["ServerRelativeUrl"]
        item = sta.fetch_uploaded_item(self.client, client_site_url.rstrip("/"), file_url)
        metadata = {
            "ItemParentID": task["ID"],
            "Processo": task["Processo"]["Title"] if task.get("Processo") else None,
            "Area": task["Area"]["Title"] if task.get("Area") else None,
            "Competencia": task.get("Competencia"),
        }
        metadata = {key: value for key, value in metadata.items() if value not in (None, "")}
        sta.update_file_metadata(self.client, client_site_url.rstrip("/"), file_url, item["__metadata"]["type"], metadata)
        return file_url

    def finalize(self, task_id: int, revisao_assignee: str) -> None:
        sta = self.sta
        task = sta.fetch_task_for_finalize(self.client, task_id)
        usuario_id = sta.resolve_user_info_id(self.client, revisao_assignee)
        payload = sta.build_finalize_payload(task, None, usuario_atual_id=usuario_id)
        sta.update_task_item(self.client, task_id, task["__metadata"]["type"], payload)


# --------------------------------------------------------------------------- fluxo de um pedido
def process_request(request: dict[str, Any], api: PortalDPApi, sp: SharePointOps, work_dir: Path,
                    now: datetime | None = None) -> tuple[str, int | None, str]:
    """Executa um pedido já reivindicado. Retorna (status, task_id, mensagem)."""
    payload = request["payload"]
    request_id = request["id"]
    task_id: int | None = request.get("task_id")
    created_here = False
    adopted = False

    if not task_id and payload.get("check_existing"):
        existing = sp.find_monthly_task(payload["process_id"], payload["cliente_id"], payload["competencia"])
        if existing:
            step = existing.get("PassoAtual")
            if step != "Execucao":
                return (
                    "aviso", int(existing["Id"]),
                    f"Já existe a tarefa #{existing['Id']} da rotina mensal de {payload['competencia']}, "
                    f"em {STEP_LABELS.get(step, step)}. Nada foi criado nem enviado: anexe os documentos "
                    "na tarefa manualmente, se necessário.",
                )
            task_id = int(existing["Id"])
            adopted = True

    if not task_id:
        responsavel = sp.responsavel_folha_id(payload["cliente_id"])
        task_id = sp.create_task(build_create_payload(payload, responsavel, now))
        created_here = True
        api.result(request_id, "enviando", task_id, "Tarefa criada; enviando os documentos.")

    step = sp.task_step(task_id)
    files_sent = 0
    if step == "Execucao":
        names = sharepoint_file_names(request["attachments"], request["ticket_id"])
        for attachment in request["attachments"]:
            data = api.download(request_id, attachment["id"])
            target = work_dir / names[attachment["id"]]
            target.write_bytes(data)
            sp.attach(task_id, payload["client_site_url"], target)
            files_sent += 1
        sp.finalize(task_id, payload["revisao_assignee"])
    elif step != "RevisaoEValidacao":
        return (
            "aviso", task_id,
            f"A tarefa #{task_id} já está em {STEP_LABELS.get(step, step)}; nada foi alterado.",
        )

    origem = "criada" if created_here else ("localizada (rotina mensal)" if adopted else "retomada")
    return (
        "criada", task_id,
        f"Tarefa #{task_id} {origem}; {files_sent} documento(s) anexado(s); enviada para Revisão e Validação.",
    )


def run_once(api: PortalDPApi, sp: SharePointOps, dry_run: bool = False) -> int:
    pending = api.pending()
    print(f"{len(pending)} pedido(s) pendente(s).")
    failures = 0
    for request in pending:
        rid = request["id"]
        if dry_run:
            payload = request["payload"]
            print(json.dumps({
                "pedido": rid, "chamado": request["ticket_id"], "processo": payload.get("process_id"),
                "cliente": payload.get("cliente_id"), "competencia": payload.get("competencia"),
                "datas": task_dates(date.fromisoformat(payload["prazo"])),
                "revisao": payload.get("revisao_assignee"), "mensal_verificar_existente": payload.get("check_existing"),
                "arquivos": list(sharepoint_file_names(request["attachments"], request["ticket_id"]).values()),
                "descricao": payload.get("descricao"),
            }, ensure_ascii=False, indent=2))
            continue
        claimed = api.claim(rid)
        if not claimed:
            print(f"pedido {rid}: já reivindicado por outra execução.")
            continue
        task_id = claimed.get("task_id")
        try:
            with tempfile.TemporaryDirectory(prefix="portaldp-") as tmp:
                status, task_id, message = process_request(claimed, api, sp, Path(tmp))
        except Exception as exc:  # noqa: BLE001 - qualquer falha vira "erro" reenviável no PortalDP
            failures += 1
            status, message = "erro", f"{type(exc).__name__}: {str(exc)[:300]}"
            print(f"pedido {rid}: {message}", file=sys.stderr)
        api.result(rid, status, task_id, message)
        print(f"pedido {rid}: {status} - {message}")
    return failures


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--dry-run", action="store_true", help="Só mostra o que faria (não reivindica pedidos nem escreve no SharePoint).")
    args = parser.parse_args()

    token = os.environ.get("PORTALDP_INTERNAL_TOKEN", "").strip()
    if not token:
        print("PORTALDP_INTERNAL_TOKEN não configurado.", file=sys.stderr)
        return 2
    lock = open(LOCK_PATH, "w")
    try:
        fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except BlockingIOError:
        print("Outra execução do worker está em andamento.")
        return 0
    api = PortalDPApi(os.environ.get("PORTALDP_URL", DEFAULT_PORTALDP_URL), token)
    sp = None if args.dry_run else SharePointOps()
    failures = run_once(api, sp, dry_run=args.dry_run)  # type: ignore[arg-type]
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
