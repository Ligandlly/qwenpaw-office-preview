.PHONY: test backend-test frontend-test frontend-build lint validate

test: backend-test frontend-test

backend-test:
	.venv/bin/python -m pytest -q

frontend-test:
	cd frontend && npm run test

frontend-build:
	cd frontend && npm run typecheck && npm run build

lint:
	uvx ruff check . --exclude .venv,frontend/node_modules,frontend

validate:
	qwenpaw plugin validate .
