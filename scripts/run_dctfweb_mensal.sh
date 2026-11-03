#!/usr/bin/env bash
set -euo pipefail

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

ENQUEUE_ARGS=(enqueue-monthly-routine --routine-key dctfweb_mensal "${RUN_DATE_ARG[@]}")
PROCESS_PERIOD_REFERENCE="$(python3 sharepoint_task_attachment.py describe-monthly-routine --routine-key dctfweb_mensal "${RUN_DATE_ARG[@]}" | python3 -c 'import json,sys; print(json.load(sys.stdin)["period_reference"])')"
PROCESS_ARGS=(
  process-pending-creates
  --routine-key dctfweb_mensal
  --period-reference "$PROCESS_PERIOD_REFERENCE"
  --limit "${LIMIT:-200}"
)

if [ "${DRY_RUN:-0}" = "1" ]; then
  ENQUEUE_ARGS+=(--dry-run)
  PROCESS_ARGS+=(--dry-run)
fi

python3 sharepoint_task_attachment.py "${ENQUEUE_ARGS[@]}"
python3 sharepoint_task_attachment.py "${PROCESS_ARGS[@]}"
