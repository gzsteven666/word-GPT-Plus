#!/usr/bin/env bash
set -euo pipefail

MANIFEST_URL="${WORD_GPT_MANIFEST_URL:-https://raw.githubusercontent.com/gzsteven666/word-GPT-Plus/master/release/instant-use/manifest.xml}"
WORD_WEF_DIR="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
DESTINATION="$WORD_WEF_DIR/GPT-Plus-Steven.xml"
TEMP_FILE="$(mktemp)"
trap 'rm -f "$TEMP_FILE"' EXIT

printf 'Downloading GPT Plus Steven manifest...\n'
curl --fail --location --silent --show-error "$MANIFEST_URL" --output "$TEMP_FILE"

if ! grep -q '<OfficeApp' "$TEMP_FILE" || ! grep -q 'GPT Plus Steven' "$TEMP_FILE"; then
  printf 'Error: the downloaded file is not the expected Office add-in manifest.\n' >&2
  exit 1
fi

mkdir -p "$WORD_WEF_DIR"
if [[ -f "$DESTINATION" ]]; then
  cp "$DESTINATION" "$DESTINATION.backup"
fi
cp "$TEMP_FILE" "$DESTINATION"
chmod 600 "$DESTINATION"

printf '\nInstalled manifest:\n  %s\n' "$DESTINATION"
if pgrep -x 'Microsoft Word' >/dev/null 2>&1; then
  printf '\nMicrosoft Word is running. Save your documents, quit Word completely, and reopen it.\n'
else
  printf '\nOpen Microsoft Word and a .docx document.\n'
fi
printf 'Then choose Home > Add-ins > GPT Plus Steven.\n'
printf 'Future app updates load automatically from GitHub Pages when the add-in is reopened.\n'
