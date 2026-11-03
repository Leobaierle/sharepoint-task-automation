#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")/.."

if [ -f ".env.sharepoint" ]; then
  set -a
  # shellcheck disable=SC1091
  . ".env.sharepoint"
  set +a
fi

ROUTINE_KEY=""
while [ $# -gt 0 ]; do
  case "$1" in
    --routine-key)
      ROUTINE_KEY="$2"
      shift 2
      ;;
    *)
      echo "Argumento desconhecido: $1" >&2
      exit 1
      ;;
  esac
done

if [ -z "$ROUTINE_KEY" ]; then
  echo "Uso: run_fiscal_routine_manual.sh --routine-key <chave>" >&2
  exit 1
fi

case "$ROUTINE_KEY" in
  fiscal_rotina_mensal|fiscal_pis_cofins_mes_atual|fiscal_sped_contribuicoes_mes_anterior|fiscal_parcelamento_tributos)
    ;;
  *)
    echo "Rotina fiscal desconhecida: $ROUTINE_KEY" >&2
    exit 1
    ;;
esac

WORKBOOK_PATH="${WORKBOOK_PATH:-./GERAÇÃO DE AGENDA - FISCAL.xlsx}"
LIMIT_VALUE="${LIMIT:-500}"

# Disparo manual sempre usa a data de hoje: mesma competencia que o
# agendamento automatico do dia 25 usaria se rodasse agora no mesmo mes.
SYNC_ARGS=(sync-fiscal-workbook --workbook "$WORKBOOK_PATH")
if [ "${DRY_RUN:-0}" = "1" ]; then
  SYNC_ARGS+=(--dry-run)
fi
python3 sharepoint_task_attachment.py "${SYNC_ARGS[@]}"

ENQUEUE_ARGS=(enqueue-monthly-routine --routine-key "$ROUTINE_KEY")
PROCESS_PERIOD_REFERENCE="$(python3 sharepoint_task_attachment.py describe-monthly-routine --routine-key "$ROUTINE_KEY" | python3 -c 'import json,sys; print(json.load(sys.stdin)["period_reference"])')"
PROCESS_ARGS=(
  process-pending-creates
  --routine-key "$ROUTINE_KEY"
  --period-reference "$PROCESS_PERIOD_REFERENCE"
  --limit "$LIMIT_VALUE"
)

if [ "${DRY_RUN:-0}" = "1" ]; then
  ENQUEUE_ARGS+=(--dry-run)
  PROCESS_ARGS+=(--dry-run)
fi

python3 sharepoint_task_attachment.py "${ENQUEUE_ARGS[@]}"
python3 sharepoint_task_attachment.py "${PROCESS_ARGS[@]}"
