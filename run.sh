#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

# Ensure we have a Linux/WSL-style venv with bin/activate
if [ ! -f .venv/bin/activate ]; then
	echo "(Re)creating virtualenv..."
	rm -rf .venv
	python3 -m venv .venv
fi

source .venv/bin/activate
pip install -r requirements.txt

if [ ! -f .env ] && [ -f env.example ]; then
	cp env.example .env
fi

exec python3 -m streamlit run app.py


