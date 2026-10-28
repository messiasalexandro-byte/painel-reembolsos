#!/bin/sh
# Envia o painel para a implantação de teste (/dev), sem mexer no /exec.
# Requer: npx @google/clasp login (uma vez, com a conta dona do script).
set -e
cd "$(dirname "$0")"
cp ../index.html Index.html
npx -y @google/clasp@3.4.1 push -f
