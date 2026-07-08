#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

if ! command -v clasp >/dev/null 2>&1; then
  echo "clasp не установлен. Выполните: npm install -g @google/clasp && clasp login"
  exit 1
fi

ENV_FILE="$ROOT/scripts/deploy.env"
if [[ ! -f "$ENV_FILE" ]]; then
  echo "Нет файла scripts/deploy.env"
  echo "Скопируйте: cp scripts/deploy.env.example scripts/deploy.env"
  exit 1
fi

# shellcheck source=/dev/null
source "$ENV_FILE"

if [[ -z "${DEV_SCRIPT_ID:-}" ]]; then
  echo "Заполните DEV_SCRIPT_ID в scripts/deploy.env"
  exit 1
fi

echo "→ clasp push в DEV (scriptId: ${DEV_SCRIPT_ID})"
clasp push --force --projectId "$DEV_SCRIPT_ID"

cat <<'EOF'

Готово (код в DEV-проекте обновлён).

Дальше вручную в редакторе DEV Apps Script:
  1. Развернуть → Новая версия
  2. Открыть DEV URL веб-приложения и проверить

EOF
