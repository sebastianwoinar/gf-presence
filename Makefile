PYTHON := .venv/bin/python3

.PHONY: all generate generate-current upload help

all: generate upload

help:
	@echo "Usage:"
	@echo "  make generate                  Generate report for NEXT month (default)"
	@echo "  make generate MONTH=YYYY-MM    Generate report for a specific month"
	@echo "  make generate-current          Generate report for the CURRENT month"
	@echo "  make upload FILE=path/to/file  Upload a generated report to Google Drive"
	@echo "  make all MONTH=YYYY-MM FILE=.. Generate then upload"
	@echo ""
	@echo "Examples:"
	@echo "  make generate-current"
	@echo "  make generate MONTH=2026-07"

generate:
	$(PYTHON) generate_praesenzplanung.py $(MONTH)

generate-current:
	$(PYTHON) generate_praesenzplanung.py $$(date +%Y-%m)

upload:
	$(PYTHON) upload_to_gdrive.py $(FILE)
