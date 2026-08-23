# Troubleshooting

## Word report is not created

Rebuild and start the internal report service:

```powershell
docker compose up -d --build
docker compose ps
docker compose logs report-service
```

The service must be healthy before n8n starts. Confirm that `report_service_url` is `http://report-service:8080/report` and that `report_path` is under `/files`.

## Word report cannot be overwritten

Close `data/output/relatorio-despesas.docx` in Microsoft Word while the workflow runs. Word may lock the file and prevent the report service from replacing it.

## Configuration file cannot be read

Recreate the n8n container after pulling the file-based configuration changes so the new read-only mounts are applied:

```powershell
docker compose down
docker compose up -d
```

Inside the container, the expected files are `/config/runtime/extraction.json`, `/config/prompts/invoice-extractor.md`, and `/config/schemas/invoice.schema.json`.

## `Missing required configuration field`

The **Parse extraction config** node validates `config/extraction.json`. Restore the missing field using `docs/configuration.md` as reference. Invalid JSON stops the workflow deliberately before any invoice is processed.

## `Access to the file is not allowed`

n8n 2.x restricts filesystem access by default. The supplied Compose file sets:

```text
N8N_RESTRICT_FILE_ACCESS_TO=/files;/config
```

`/files` contains writable runtime data. `/config` grants access to the read-only runtime configuration, prompt, and schema mounts.

After changing the Compose file, recreate the container:

```powershell
docker compose down
docker compose up -d
```

Confirm that the files are visible inside the container:

```powershell
docker compose exec n8n ls -la /files/inbox
```

## `access to env vars denied`

n8n blocks environment-variable access in node expressions by default. The workflow therefore stores the Amália model identifier directly in the **Amalia via Ollama** node. Do not replace it with `{{ $env.OLLAMA_MODEL }}` unless the n8n security policy is deliberately changed.

## Ollama connection failure

Use this base URL in the n8n Ollama credential:

```text
http://host.docker.internal:11434
```

Check the host service and installed model:

```powershell
ollama list
Invoke-RestMethod http://127.0.0.1:11434/api/tags
```

## `The file path was expected but the given path is a directory`

The write node received a directory because the filename expression evaluated to an empty value. The current **Copy PDF to processed** node reads the filename from `$binary.data.fileName`. Reimport the latest workflow or use:

```javascript
{{ '/files/processed/' + $binary.data.fileName.replace(/[^a-zA-Z0-9._-]/g, '_') }}
```

## Excel register not found

The latest workflow treats a missing register as the first execution and creates `data/output/invoices.xlsx` automatically. Reimport the current `workflows/invoice-extraction.json` if an older imported workflow still stops at **Read invoice register**.

## Excel register cannot be written

Close `data/output/invoices.xlsx` in Microsoft Excel before running the workflow. Excel may lock the file and prevent n8n from overwriting it.

## Extraction goes to `data/error`

Open the corresponding `<invoice-name>.error.json`. It contains validation errors and the raw model response. Common causes include an unreadable PDF, a missing total, or invalid JSON from the model.

## Scanned PDF returns no useful text

The current workflow only handles PDFs with embedded searchable text. Scanned or image-only invoices need an OCR stage before the Amália agent.

## `invoices.xlsx.inspect.ndjson`

This is a development-time workbook inspection report. It is not used by n8n and can be deleted safely. Runtime inspection files are ignored by Git.
