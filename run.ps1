Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Set-Location $PSScriptRoot

if (-not (Test-Path .venv)) {
	py -3 -m venv .venv
}

.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt

if (-not (Test-Path .env) -and (Test-Path env.example)) {
	Copy-Item env.example .env
}

python -m streamlit run app.py


