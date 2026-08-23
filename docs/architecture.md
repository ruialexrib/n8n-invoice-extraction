# Architecture

## Pipeline

1. **Configuration bootstrap** — each execution reads the runtime JSON, Markdown prompt, and JSON Schema from read-only container mounts.
2. **Ingestion** — a Schedule Trigger runs every minute and reads searchable PDF files using the configured input glob. A Manual Trigger remains available for testing.
3. **Text extraction** — n8n extracts embedded text from each PDF. Scanned or image-only documents require a future OCR stage.
4. **Local AI extraction** — the configured Ollama model uses the external prompt and schema to extract supplier, invoice, customer, and payment fields.
5. **Normalization** — Code nodes apply configured provider mappings and normalize Multibanco fields and confidence scores.
6. **Validation** — the workflow checks the provider type, total, currency, confidence, warnings, and core object structure.
7. **Valid-result branches** — after validation, n8n copies the PDF to the configured processed directory while it reads the existing register and supplies the new row for the register merge. These branches run independently, so the PDF copy can complete before the output files.
8. **Excel registration** — if the configured register exists, its rows are merged with the new row. Otherwise a new workbook is created automatically. Duplicate invoices are removed before writing.
9. **Word reporting** — after writing Excel, n8n sends all deduplicated register rows to an internal service that creates or rebuilds the Portuguese expense report.
10. **Invalid-result routing** — an invalid extraction writes its diagnostic JSON and then copies the PDF to the configured error directory.

The original PDF remains in `data/inbox` during the MVP to avoid destructive file operations.

## Extracted data contract

The operational output is intentionally narrow:

- Processing timestamp and source filename
- Provider and supplier
- Invoice number and issue date
- Customer name, tax ID, account/contract number, and address
- Payment due date and method
- Multibanco entity and reference
- IBAN
- Total amount and currency
- Extraction confidence and warnings

Identifiers are stored as text to preserve leading zeros. Dates use ISO 8601 (`YYYY-MM-DD`) and amounts are decimal numbers.

## Local components

- **n8n** runs in Docker and mounts `./data` at `/files`.
- **Ollama** runs on the Windows host and is reached from n8n at `http://host.docker.internal:11434`.
- **Amália** defaults to `hf.co/ruialexrib/AMALIA-9B-0626-SFT-GGUF:Q3_K_M`.
- **Excel register** is a local `.xlsx` file rebuilt by n8n after each successful extraction.
- **Word report service** is a private Docker service using `python-docx`; it writes the configured `.docx` under `/files` and is not exposed on a host port.

## Idempotency

The Excel register uses `Source File + Invoice Number` as its current deduplication key. Reprocessing the same invoice does not append another row. A future production version should use a SHA-256 content hash and define whether corrected invoices replace or version earlier records.

## Repository and runtime data

The repository does not version Excel or Word outputs. The workflow creates both on the first valid extraction. All `.xlsx` and `.docx` files, invoices, processed copies, error reports, and diagnostic files are ignored by Git.

This separation prevents real customer names, tax IDs, addresses, Multibanco references, and IBANs from being included in a commit.

## Recommended evolution

- Add OCR for scanned PDFs and image files.
- Validate normalized results against the JSON Schema with a dedicated validator.
- Add SHA-256-based deduplication and a controlled retention policy.
- Add a processed-file ledger or move confirmed source files out of the inbox to avoid repeated model executions on the one-minute schedule.
- Add anonymized regression fixtures for NOS, EDP, and water invoices.
- Add workbook locking/retry handling for concurrent executions.
- Move the operational register to PostgreSQL or an ERP when multi-user access or stronger auditability is required.
