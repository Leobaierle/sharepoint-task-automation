#!/usr/bin/env bash
set -euo pipefail

# Marcadores de etapa (stderr): o painel de automações lê a última "@@ etapa:" para dizer onde parou.
ETAPA="início"
etapa() { ETAPA="$1"; echo "@@ etapa: $ETAPA" >&2; }
trap 'rc=$?; echo "@@ falhou: $ETAPA (exit $rc)" >&2' ERR

cd "$(dirname "$0")/.."

if [ -f ".env.sharepoint" ]; then
  set -a
  # shellcheck disable=SC1091
  . ".env.sharepoint"
  set +a
fi

WORKBOOK_PATH="${WORKBOOK_PATH:-./GERAÇÃO DE AGENDA - FISCAL.xlsx}"
LIMIT_VALUE="${LIMIT:-500}"
RUN_DATE_ARG=()
if [ -n "${RUN_DATE:-}" ]; then
  RUN_DATE_ARG=(--run-date "$RUN_DATE")
fi

RUN_DAY="$(python3 - <<'PY'
import os
from datetime import datetime, timezone
run_date = os.getenv("RUN_DATE")
if run_date:
    print(int(run_date.split("-")[2]))
else:
    print(datetime.now(timezone.utc).day)
PY
)"

if [ "$RUN_DAY" -eq 1 ]; then
  etapa "ativar pendência da Rotina Mensal"
  ACTIVATE_ARGS=(set-fiscal-rotina-mensal-pendencia --mode activate "${RUN_DATE_ARG[@]}")
  python3 sharepoint_task_attachment.py "${ACTIVATE_ARGS[@]}"
  exit 0
fi

SYNC_ARGS=(sync-fiscal-workbook --workbook "$WORKBOOK_PATH")
if [ "${DRY_RUN:-0}" = "1" ]; then
  SYNC_ARGS+=(--dry-run)
fi
etapa "ler a planilha GERAÇÃO DE AGENDA"
python3 sharepoint_task_attachment.py "${SYNC_ARGS[@]}"

ROUTINES=(
  fiscal_rotina_mensal
  fiscal_pis_cofins_mes_atual
  fiscal_sped_contribuicoes_mes_anterior
  fiscal_parcelamento_tributos
)

for routine_key in "${ROUTINES[@]}"; do
  etapa "calcular a competência de $routine_key"
  ENQUEUE_ARGS=(enqueue-monthly-routine --routine-key "$routine_key" "${RUN_DATE_ARG[@]}")
  PROCESS_PERIOD_REFERENCE="$(python3 sharepoint_task_attachment.py describe-monthly-routine --routine-key "$routine_key" "${RUN_DATE_ARG[@]}" | python3 -c 'import json,sys; print(json.load(sys.stdin)["period_reference"])')"
  PROCESS_ARGS=(
    process-pending-creates
    --routine-key "$routine_key"
    --period-reference "$PROCESS_PERIOD_REFERENCE"
    --limit "$LIMIT_VALUE"
  )

  if [ "${DRY_RUN:-0}" = "1" ]; then
    ENQUEUE_ARGS+=(--dry-run)
    PROCESS_ARGS+=(--dry-run)
  fi

  etapa "enfileirar $routine_key ($PROCESS_PERIOD_REFERENCE)"
  python3 sharepoint_task_attachment.py "${ENQUEUE_ARGS[@]}"
  etapa "criar as tarefas de $routine_key ($PROCESS_PERIOD_REFERENCE)"
  python3 sharepoint_task_attachment.py "${PROCESS_ARGS[@]}"
done
