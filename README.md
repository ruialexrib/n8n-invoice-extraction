<div align="center">

# n8n Invoice Extraction

### Local AI-assisted extraction of Portuguese utility invoices

[![n8n](https://img.shields.io/badge/n8n-Workflow%20Automation-EA4B71?logo=n8n&logoColor=white)](https://n8n.io/)
[![Ollama](https://img.shields.io/badge/Ollama-Local%20LLM-black)](https://ollama.com/)
[![Docker](https://img.shields.io/badge/Docker-Compose-2496ED?logo=docker&logoColor=white)](https://www.docker.com/)
[![Excel](https://img.shields.io/badge/Output-Excel-217346?logo=microsoftexcel&logoColor=white)](#outputs)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

**Invoice processing · Local LLM · Structured extraction · Excel register · Word reporting**

</div>

---

## About

**n8n Invoice Extraction** is an MVP for processing Portuguese telecommunications, electricity, and water invoices. It extracts customer and payment information using a locally hosted quantised **AMALIA** model through Ollama, validates the structured result, and maintains a local Excel register.

Valid extractions are also used to generate a consolidated Word expense report. The workflow is designed to keep invoice processing and model inference on local infrastructure.

---

## Workflow

```text
data/inbox
    │
    ▼
PDF Text Extraction
    │
    ▼
AMALIA via Ollama
    │
    ▼
Structured Validation
   / \
  /   \
 ▼     ▼
Valid  Invalid
 │       │
 ▼       ▼
Excel   Error report
 │      + PDF copy
 ▼
Word Expense Report
 │
 ▼
data/processed
```

The scheduled trigger checks `data/inbox` every minute.

---

## Technology Stack

| Technology | Purpose |
| --- | --- |
| **n8n** | Workflow orchestration and scheduling |
| **AMALIA-9B** | Invoice information extraction |
| **Ollama** | Local LLM execution |
| **Docker Compose** | Local runtime |
| **Excel** | Cumulative structured invoice register |
| **Word** | Consolidated expense report |
| **JSON Schema** | Extraction contract and validation |

---

## Repository Structure

```text
.
├── config/
│   └── extraction.json
├── data/
│   ├── inbox/       # Invoices awaiting processing
│   ├── processed/   # Successfully processed invoices
│   ├── error/       # Files requiring review
│   └── output/      # Generated Excel and Word files
├── docs/
│   ├── architecture.md
│   └── configuration.md
├── prompts/
│   └── invoice-extractor.md
├── schemas/
│   └── invoice.schema.json
├── workflows/
│   └── invoice-extraction.json
├── .env.example
└── docker-compose.yml
```

---

## Quick Start

1. Start Ollama and verify the model with `ollama list`.
2. If required, install it:

```bash
ollama pull hf.co/ruialexrib/AMALIA-9B-0626-SFT-GGUF:Q3_K_M
```

3. Copy `.env.example` to `.env` and change `N8N_ENCRYPTION_KEY`.
4. Start the stack:

```bash
docker compose up -d --build
```

5. Open n8n at `http://localhost:5678` and import `workflows/invoice-extraction.json`.
6. Configure the Ollama credential with `http://host.docker.internal:11434`.
7. Run the workflow manually once, then publish/activate it.
8. Place a searchable PDF invoice in `data/inbox`.

> The container maps `./data` to `/files`. Use `/files/...` paths inside n8n nodes.

---

## Configuration

The workflow reads configuration files at the beginning of every execution:

| File | Purpose |
| --- | --- |
| `config/extraction.json` | Model, paths, workbook settings, currency, and provider mappings |
| `prompts/invoice-extractor.md` | LLM extraction instructions |
| `schemas/invoice.schema.json` | Expected structured extraction contract |

Changes take effect on the next execution without reimporting the workflow.

---

## Extracted Information

The Excel register contains the main operational fields, including processing timestamp, source filename, provider, supplier, invoice number, issue date, customer information, tax ID, account or contract number, address, payment due date, payment method, Multibanco entity and reference, IBAN, total amount, currency, confidence, and warnings.

Uncertain values must be returned as `null` and described in `warnings`; the workflow is designed not to fabricate missing information.

---

## Outputs

For valid extractions, the workflow creates or updates:

```text
data/output/invoices.xlsx
data/output/relatorio-despesas.docx
```

The Word report is generated in Portuguese (Portugal) and contains a consolidated summary, totals by category, processed expenses, payment information, and extraction warnings.

Invalid model output produces:

```text
data/error/<invoice-name>.error.json
data/error/<invoice-name>.pdf
```

The source PDF remains in `data/inbox`. Excel deduplication uses source filename plus invoice number, but confirmed source files should still be removed from the inbox to avoid unnecessary repeated model executions.

---

## Current Scope & Limitations

- Input PDFs must contain searchable text.
- Current provider categories are telecommunications, electricity, and water.
- Scanned PDFs require an OCR stage that is not implemented by default.
- Output files should not remain open while the workflow overwrites them.
- The workflow is an MVP and should be reviewed before production use.

---

## Security

Invoices can contain personal and confidential information such as tax IDs, addresses, IBANs, and payment references. Do not commit invoices, generated results, secrets, `.env`, or the `.n8n` directory.

Before processing real data, define appropriate retention, access-control, encryption, backup, and deployment policies. Keep Ollama on the local machine or a trusted network and do not expose port `11434` publicly without authentication and transport security.

---

## License

This project is licensed under the [MIT License](LICENSE).
