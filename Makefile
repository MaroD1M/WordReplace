.PHONY: venv install run test lint

venv:
	python3 -m venv .venv

install:
	.venv/bin/python -m pip install -r requirements.txt
	.venv/bin/python -m pip install pytest

run:
	.venv/bin/streamlit run app/main.py

test:
	.venv/bin/python -m pytest -q tests

lint:
	python3 -m py_compile app/main.py app/core_utils.py app/services.py
