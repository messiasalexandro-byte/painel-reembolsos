#!/bin/sh
# Publica o painel no Apps Script mantendo a mesma URL.
# Requer: npx @google/clasp login (uma vez, com a conta dona do script).
set -e
cd "$(dirname "$0")"
cp ../index.html Index.html
npx -y @google/clasp@3.4.1 push -f
npx -y @google/clasp@3.4.1 update-deployment AKfycbxgwqnN5kGArZmYM-zbGyV_h4AjfkzqITj2ZTfVP1PfpZMDMu2Ag8FDnlDpEGiWbXVy --description "painel $(date +%Y-%m-%d)"
