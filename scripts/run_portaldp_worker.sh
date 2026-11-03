#!/usr/bin/env bash
# Worker PortalDP -> tarefa no portal de tarefas. Chamado pelo timer (portaldp-worker.timer) a cada minuto.
# Marcadores de etapa (stderr) no mesmo padrão dos outros runners.
set -euo pipefail

ETAPA="início"
etapa() { ETAPA="$1"; echo "@@ etapa: $ETAPA" >&2; }
trap 'rc=$?; echo "@@ falhou: $ETAPA (exit $rc)" >&2' ERR

cd "$(dirname "$0")/.."

for env_file in .env.sharepoint .env.portaldp; do
  if [ -f "$env_file" ]; then
    set -a
    # shellcheck disable=SC1090
    . "$env_file"
    set +a
  fi
done

etapa "processar pedidos pendentes do PortalDP"
exec python3 scripts/portaldp_worker.py "$@"
