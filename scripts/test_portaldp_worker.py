"""Testes do worker PortalDP -> SharePoint (sem rede: PortalDP e SharePoint simulados).

Rodar: python3 -m unittest scripts.test_portaldp_worker -v   (na raiz do repo)
"""
import sys
import tempfile
import unittest
from datetime import date, datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import portaldp_worker as w  # noqa: E402


class FakeApi:
    def __init__(self, files=None):
        self.results = []
        self.files = files or {}

    def download(self, request_id, attachment_id):
        return self.files.get(attachment_id, b"DADOS")

    def result(self, request_id, status, task_id=None, message=""):
        self.results.append((request_id, status, task_id, message))


class FakeSp:
    def __init__(self, existing=None, step="Execucao", fail_on=None):
        self.existing = existing
        self.step = step
        self.fail_on = fail_on
        self.calls = []
        self.uploaded = []

    def find_monthly_task(self, process_id, cliente_id, competencia):
        self.calls.append(("find", process_id, cliente_id, competencia))
        return self.existing

    def responsavel_folha_id(self, cliente_id):
        self.calls.append(("responsavel", cliente_id))
        return 300123

    def create_task(self, payload):
        self.calls.append(("create", payload))
        return 84600

    def task_step(self, task_id):
        return self.step

    def attach(self, task_id, site, path):
        if self.fail_on == "attach":
            raise RuntimeError("SharePoint 503")
        self.uploaded.append((task_id, site, path.name, path.read_bytes()))
        self.calls.append(("attach", task_id, path.name))

    def finalize(self, task_id, assignee):
        self.calls.append(("finalize", task_id, assignee))


def make_request(**over):
    base = {
        "id": 7, "ticket_id": 55, "status": "enviando", "task_id": None, "attempts": 1, "message": None,
        "payload": {
            "ticket_id": 55, "request_category": "Admissão", "process_id": 29, "area_id": 3, "cliente_id": 273,
            "client_site_url": "https://empresa.sharepoint.com/sites/Empresa/gama",
            "revisao_assignee": "31", "competencia": "Setembro 2026", "prazo": "2026-10-03",
            "descricao": "PortalDP #55 — Admissão\nMatriz\nSolicitante: Lucas\nTítulo: X", "check_existing": False,
        },
        "attachments": [{"id": 11, "name": "contrato.pdf", "size": 5, "content_type": "application/pdf"}],
    }
    base.update(over)
    return base


class DatesAndNamesTest(unittest.TestCase):
    def test_task_dates_bate_com_tarefas_manuais_do_portal(self):
        # Tarefas reais (portal, 29/09/2026): Prazo 29/09 -> ExecPrazo 25/09, Margem 26/09, Venc 27/09, VencMargem 29/09
        d = w.task_dates(date(2026, 9, 29))
        self.assertEqual(d["Prazo"], "2026-09-29T00:00:00")
        self.assertEqual(d["ExecucaoPrazo"], "2026-09-25T03:00:00Z")
        self.assertEqual(d["ExecucaoMargem"], "2026-09-26T03:00:00Z")
        self.assertEqual(d["Vencimento"], "2026-09-27T03:00:00Z")
        self.assertEqual(d["VencimentoMargem"], "2026-09-29T03:00:00Z")
        # Prazo 02/10 -> 28/09, 29/09, 30/09, 02/10 (atravessa o mes)
        d = w.task_dates(date(2026, 10, 2))
        self.assertEqual(
            (d["ExecucaoPrazo"], d["ExecucaoMargem"], d["Vencimento"], d["VencimentoMargem"]),
            ("2026-09-28T03:00:00Z", "2026-09-29T03:00:00Z", "2026-09-30T03:00:00Z", "2026-10-02T03:00:00Z"),
        )

    def test_nomes_de_arquivo_seguros_e_unicos(self):
        atts = [{"id": 1, "name": "folha.pdf"}, {"id": 2, "name": "folha.pdf"}, {"id": 3, "name": "a#b%c/d.xlsx"}]
        names = w.sharepoint_file_names(atts, 55)
        self.assertEqual(names[1], "folha - PortalDP 55.pdf")
        self.assertEqual(names[2], "folha - PortalDP 55 (2).pdf")
        self.assertEqual(names[3], "d - PortalDP 55.xlsx")  # Path.name descarta o caminho
        self.assertEqual(w.safe_file_name('a#b%c*.pdf'), "a_b_c_.pdf")

    def test_payload_de_criacao(self):
        payload = w.build_create_payload(make_request()["payload"], 300123, datetime(2026, 9, 30, 15, 0, tzinfo=timezone.utc))
        self.assertEqual(payload["Title"], "Execução")
        self.assertEqual((payload["AreaId"], payload["ProcessoId"], payload["ClienteId"]), ("3", "29", "273"))
        self.assertEqual(payload["UsuarioAtualId"], 300123)
        self.assertEqual(payload["PassoAtual"], "Execucao")
        self.assertEqual(payload["Competencia"], "Setembro 2026")
        self.assertEqual(payload["ExecucaoInicio"], "2026-09-30T15:00:00Z")
        self.assertEqual(payload["Prazo"], "2026-10-03T00:00:00")


class ProcessRequestTest(unittest.TestCase):
    def run_req(self, sp, request=None, api=None):
        api = api or FakeApi()
        with tempfile.TemporaryDirectory() as tmp:
            return w.process_request(request or make_request(), api, sp, Path(tmp)), api

    def test_fluxo_completo_cria_anexa_e_avanca(self):
        sp = FakeSp()
        (status, task_id, message), api = self.run_req(sp)
        self.assertEqual((status, task_id), ("criada", 84600))
        self.assertIn("Revisão e Validação", message)
        ordem = [c[0] for c in sp.calls]
        self.assertEqual(ordem, ["responsavel", "create", "attach", "finalize"])
        self.assertEqual(sp.calls[-1], ("finalize", 84600, "31"))
        self.assertEqual(sp.uploaded[0][2], "contrato - PortalDP 55.pdf")
        self.assertEqual(sp.calls[1][1]["UsuarioAtualId"], 300123)  # Execução = ResponsavelFolha
        # task_id foi devolvido logo após criar (para retomada em caso de falha depois)
        self.assertEqual(api.results[0][:3], (7, "enviando", 84600))

    def test_processo_mensal_existente_em_execucao_e_reaproveitado(self):
        req = make_request()
        req["payload"].update(process_id=9, check_existing=True, competencia="Setembro 2026")
        sp = FakeSp(existing={"Id": 83000, "PassoAtual": "Execucao"})
        (status, task_id, message), _ = self.run_req(sp, req)
        self.assertEqual((status, task_id), ("criada", 83000))
        self.assertIn("localizada", message)
        self.assertNotIn("create", [c[0] for c in sp.calls])
        self.assertIn(("finalize", 83000, "31"), sp.calls)

    def test_processo_mensal_alem_da_execucao_avisa_e_nao_mexe(self):
        req = make_request()
        req["payload"].update(process_id=9, check_existing=True)
        sp = FakeSp(existing={"Id": 83727, "PassoAtual": "Fechamento"})
        (status, task_id, message), _ = self.run_req(sp, req)
        self.assertEqual((status, task_id), ("aviso", 83727))
        self.assertIn("#83727", message)
        self.assertIn("Fechamento", message)
        self.assertEqual([c[0] for c in sp.calls], ["find"])  # nada criado, anexado ou finalizado

    def test_processo_mensal_inexistente_cria(self):
        req = make_request()
        req["payload"].update(process_id=9, check_existing=True)
        sp = FakeSp(existing=None)
        (status, task_id, _), _ = self.run_req(sp, req)
        self.assertEqual((status, task_id), ("criada", 84600))
        self.assertIn("create", [c[0] for c in sp.calls])

    def test_processo_eventual_nao_consulta_existente(self):
        sp = FakeSp(existing={"Id": 1, "PassoAtual": "Fechamento"})
        self.run_req(sp)
        self.assertNotIn("find", [c[0] for c in sp.calls])

    def test_retomada_com_task_id_nao_cria_de_novo(self):
        sp = FakeSp()
        (status, task_id, message), _ = self.run_req(sp, make_request(task_id=84600))
        self.assertEqual((status, task_id), ("criada", 84600))
        self.assertNotIn("create", [c[0] for c in sp.calls])
        self.assertIn("retomada", message)

    def test_retomada_com_tarefa_ja_em_revisao_so_confirma(self):
        sp = FakeSp(step="RevisaoEValidacao")
        (status, task_id, _), _ = self.run_req(sp, make_request(task_id=84600))
        self.assertEqual((status, task_id), ("criada", 84600))
        self.assertEqual(sp.calls, [])  # nao reenvia nem refinaliza

    def test_tarefa_em_etapa_inesperada_avisa(self):
        sp = FakeSp(step="Fechamento")
        (status, _, message), _ = self.run_req(sp, make_request(task_id=84600))
        self.assertEqual(status, "aviso")
        self.assertIn("Fechamento", message)

    def test_varios_anexos_com_mesmo_nome_nao_se_sobrescrevem(self):
        req = make_request()
        req["attachments"] = [
            {"id": 1, "name": "folha.pdf", "size": 1, "content_type": ""},
            {"id": 2, "name": "folha.pdf", "size": 1, "content_type": ""},
        ]
        sp = FakeSp()
        self.run_req(sp, req, FakeApi(files={1: b"A", 2: b"B"}))
        self.assertEqual([(u[2], u[3]) for u in sp.uploaded],
                         [("folha - PortalDP 55.pdf", b"A"), ("folha - PortalDP 55 (2).pdf", b"B")])


class RunOnceTest(unittest.TestCase):
    class Api(FakeApi):
        def __init__(self, pending, claim_ok=True):
            super().__init__()
            self._pending = pending
            self._claim_ok = claim_ok

        def pending(self):
            return self._pending

        def claim(self, rid):
            return next((r for r in self._pending if r["id"] == rid), None) if self._claim_ok else None

    def test_falha_vira_erro_com_task_id_preservado(self):
        api = self.Api([make_request()])
        sp = FakeSp(fail_on="attach")
        failures = w.run_once(api, sp)
        self.assertEqual(failures, 1)
        rid, status, task_id, message = api.results[-1]
        self.assertEqual((rid, status, task_id), (7, "erro", None))  # claimed.task_id era None; a 1a chamada 'enviando' gravou 84600
        self.assertIn("SharePoint 503", message)
        self.assertEqual(api.results[0][:3], (7, "enviando", 84600))

    def test_dry_run_nao_reivindica_nem_escreve(self):
        api = self.Api([make_request()])
        sp = FakeSp()
        self.assertEqual(w.run_once(api, sp, dry_run=True), 0)
        self.assertEqual(api.results, [])
        self.assertEqual(sp.calls, [])

    def test_pedido_ja_reivindicado_por_outra_execucao_e_ignorado(self):
        api = self.Api([make_request()], claim_ok=False)
        sp = FakeSp()
        self.assertEqual(w.run_once(api, sp), 0)
        self.assertEqual(sp.calls, [])
        self.assertEqual(api.results, [])


if __name__ == "__main__":
    unittest.main()
