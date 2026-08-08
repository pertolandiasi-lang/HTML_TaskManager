#!/usr/bin/env bash
#
# Pubblica il frontend su Cloudflare Pages (progetto: taskflow).
#   ./deploy.sh
#
# Perché esiste invece di lanciare wrangler sulla cartella: qui dentro ci sono
# oauth.json, token.json, trigger_token.json e il backend Python. Sono in
# .gitignore, ma un deploy diretto della cartella se ne infischia e li
# pubblicherebbe su internet. Questo script copia una ALLOWLIST di file in una
# cartella temporanea e pubblica soltanto quella: aggiungere un file al sito
# deve essere una scelta esplicita, non un effetto collaterale.

set -euo pipefail
cd "$(dirname "$0")"

DIST="$(mktemp -d)"
trap 'rm -rf "$DIST"' EXIT

cp index.html sw.js _headers "$DIST/"

# Firma della build. Il rilevatore di aggiornamenti in index.html la rilegge
# ogni minuto: se cambia, mostra il banner "nuova versione". Serve perché
# Cloudflare Pages sulla root non manda ETag e il vecchio controllo sugli
# header lì non vedrebbe mai niente.
shasum -a 256 index.html | cut -c1-12 > "$DIST/version.txt"

echo "File pubblicati : $(cd "$DIST" && ls | tr '\n' ' ')"
echo "Versione        : $(cat "$DIST/version.txt")"

wrangler pages deploy "$DIST" \
  --project-name taskflow \
  --branch main \
  --commit-dirty=true
