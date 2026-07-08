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

if [[ -z "${PROD_SCRIPT_ID:-}" ]]; then
  echo "Заполните PROD_SCRIPT_ID в scripts/deploy.env"
  exit 1
fi

echo "Внимание: пуш в PROD затронет боевой Apps Script."
read -r -p "Продолжить? Введите yes: " CONFIRM
if [[ "$CONFIRM" != "yes" ]]; then
  echo "Отменено."
  exit 0
fi

echo "→ clasp push в PROD (scriptId: ${PROD_SCRIPT_ID})"
clasp push --force --projectId "$PROD_SCRIPT_ID"

cat <<'EOF'

Готово (код в PROD-проекте обновлён).

Дальше вручную в редакторе PROD Apps Script:
  1. Развернуть → Новая версия
  2. Развернуть → Управление развертываниями → выбрать боевое развертывание
  3. Указать новую версию (не «Последняя») → Сохранить
  4. Кратко проверить PROD URL

EOF
