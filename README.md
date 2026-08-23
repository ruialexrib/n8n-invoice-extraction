# n8n Invoice Extraction

An MVP that processes Portuguese invoices (NOS, EDP, and water utilities), extracts customer and payment data with a locally hosted, quantized Amália model through Ollama, and appends validated results to a local Excel register.

## Intended workflow

Every minute: `data/inbox` → text extraction → Amália → validation. Invalid extractions write an error report and copy the PDF to `error`. Valid extractions copy the PDF to `processed` while the workflow updates Excel and then generates the Word expense report.

The initial workflow is available at [`workflows/invoice-extraction.json`](workflows/invoice-extraction.json). The **Invoice Extraction Agent** is connected to the n8n **Ollama Chat Model** node and defaults to the locally installed `hf.co/ruialexrib/AMALIA-9B-0626-SFT-GGUF:Q3_K_M` model.

## Quick start

1. Start Ollama and verify that the model is available with `ollama list`.
2. If needed, install it with `ollama pull hf.co/ruialexrib/AMALIA-9B-0626-SFT-GGUF:Q3_K_M`.
3. Copy `.env.example` to `.env` and change `N8N_ENCRYPTION_KEY`.
4. Run `docker compose up -d --build`. This starts n8n and the internal Word report service.
5. Open `http://localhost:5678` and complete the initial setup.
6. Import `workflows/invoice-extraction.json`.
7. Open the **Amalia via Ollama** node and create an Ollama credential with base URL `http://host.docker.internal:11434`. No API key is required for the default local Ollama setup.
8. Run the workflow manually once to validate the credentials and paths.
9. Publish/activate the workflow. The **Every minute** trigger will then check `data/inbox` automatically; the browser does not need to remain open.
10. Place a PDF invoice in `data/inbox`.
11. Open `data/output/invoices.xlsx` and `data/output/relatorio-despesas.docx`. Both are created automatically on the first valid extraction and updated thereafter. A PDF whose extraction passes validation is copied to `data/processed`; invalid output and its PDF copy are written to `data/error`.

> The container maps `./data` to `/files`. Always use `/files/...` paths inside n8n nodes.

> `localhost` inside the n8n container refers to the container itself. The supplied Compose configuration maps `host.docker.internal` so n8n can reach Ollama running on the host machine.

## Local model configuration

The workflow reads the model identifier from `config/extraction.json`. The Ollama base URL remains managed in the n8n credential store and should be `http://host.docker.internal:11434` for the supplied Docker setup.

## File-based configuration

The workflow reads these files at the beginning of every execution:

- `config/extraction.json` — model, input glob, output paths, workbook settings, currency, and provider mappings.
- `prompts/invoice-extractor.md` — agent instructions.
- `schemas/invoice.schema.json` — expected extraction contract supplied to the agent.

Changes take effect on the next run without reimporting the workflow. See [`docs/configuration.md`](docs/configuration.md) for supported fields and limitations.

## Repository structure

```text
.
├── config/
│   └── extraction.json
├── data/
│   ├── inbox/       # invoices awaiting processing (not versioned)
│   ├── processed/   # successfully processed invoices
│   ├── error/       # files requiring review
│   └── output/      # cumulative Excel invoice register
├── docs/
│   └── architecture.md
├── prompts/
│   └── invoice-extractor.md
├── schemas/
│   └── invoice.schema.json
├── workflows/
│   └── invoice-extraction.json
├── .env.example
└── docker-compose.yml
```

## MVP scope

- Input: PDFs containing searchable text.
- Providers: NOS, EDP, and municipal water utilities.
- Output: one row per invoice in `data/output/invoices.xlsx`.
- Monetary values: decimal numbers without currency symbols.
- Dates: ISO 8601 (`YYYY-MM-DD`).
- Uncertain fields: use `null` and add an entry to `warnings`; never fabricate values.

## Extracted columns

The register intentionally contains only the principal operational fields:

- Processing timestamp and source filename
- Provider, supplier, invoice number, and issue date
- Customer name, tax ID, account/contract number, and address
- Payment due date, method, Multibanco entity and reference, and IBAN
- Total amount, currency, extraction confidence, and warnings

For an extraction that passes validation, the workflow starts three branches: it reads the existing register, supplies the new row for the register merge, and copies the PDF to `data/processed`. The copy can therefore finish before the Excel and Word outputs. A later Excel or report-generation failure does not move that PDF to `data/error`.

It also creates or rebuilds `data/output/relatorio-despesas.docx` in Portuguese (Portugal). The report contains a consolidated summary, totals by category, processed expenses, payment details, and extraction warnings.

If the model returns invalid JSON or fails the core validation rules, it writes:

```text
data/error/<invoice-name>.error.json
data/error/<invoice-name>.pdf
```

The source PDF is intentionally retained in `data/inbox`. This avoids destructive file operations while the workflow is being tested. The register uses source filename plus invoice number as its deduplication key, so re-running the same invoice does not add a duplicate row.

The scheduled trigger runs every minute. Because source PDFs remain in `data/inbox`, they are read again on later executions even though the Excel deduplication prevents duplicate rows. Remove confirmed source PDFs from `data/inbox` to avoid unnecessary model executions.

Do not keep the Excel or Word output open while the workflow runs because the files are overwritten after processing. The workflow creates missing output files automatically. All `.xlsx` and `.docx` files are ignored by Git because they contain personal and payment data.

Scanned PDFs require OCR. The architecture reserves a step for it, but the exact implementation depends on the selected service, such as Azure Document Intelligence, Google Document AI, or local OCR.

## Security

- Do not commit invoices, generated results, secrets, or the `.n8n` directory.
- Excel workbooks, Word reports, runtime PDFs, error reports, and inspection files are ignored by Git.
- Treat tax IDs, addresses, Multibanco references, IBANs, and invoice data as personal or confidential data.
- Define retention and access-control policies before processing real data.
- Pin the n8n image to a specific version before deploying to production.
- Keep Ollama bound to the local machine or a trusted network; do not expose port `11434` publicly without authentication and transport security.

## Troubleshooting

See [`docs/troubleshooting.md`](docs/troubleshooting.md) for the known filesystem, environment-variable, Ollama, Excel-locking, and PDF-text errors.

## License

This project is licensed under the MIT License. See [`LICENSE`](LICENSE) for details.
