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

RUN_DATE_ARG=()
if [ -n "${RUN_DATE:-}" ]; then
  RUN_DATE_ARG=(--run-date "$RUN_DATE")
fi

etapa "calcular a competência de adiantamento_mensal"
ENQUEUE_ARGS=(enqueue-monthly-routine --routine-key adiantamento_mensal "${RUN_DATE_ARG[@]}")
PROCESS_ARGS=(process-pending-creates --routine-key adiantamento_mensal --limit "${LIMIT:-200}")

if [ -n "${RUN_DATE:-}" ]; then
  PROCESS_ARGS+=(--period-reference "$(date -u -d "$RUN_DATE" +%Y-%m)")
else
  PROCESS_ARGS+=(--period-reference "$(date -u +%Y-%m)")
fi

if [ "${DRY_RUN:-0}" = "1" ]; then
  ENQUEUE_ARGS+=(--dry-run)
  PROCESS_ARGS+=(--dry-run)
fi

etapa "enfileirar adiantamento_mensal"
python3 sharepoint_task_attachment.py "${ENQUEUE_ARGS[@]}"
etapa "criar as tarefas de adiantamento_mensal"
python3 sharepoint_task_attachment.py "${PROCESS_ARGS[@]}"
