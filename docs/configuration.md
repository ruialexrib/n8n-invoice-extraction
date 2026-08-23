# Configuration

## Runtime settings

`config/extraction.json` is read and validated at the start of every execution.

| Field | Purpose |
| --- | --- |
| `ollama_model` | Exact Ollama model identifier |
| `input_glob` | Absolute container path/glob for invoice PDFs |
| `register_path` | Absolute container path of the Excel register |
| `register_file_name` | Filename embedded in the generated workbook |
| `worksheet_name` | Worksheet name created by n8n |
| `report_service_url` | Internal Docker URL of the Word report endpoint |
| `report_path` | Absolute container path of the generated `.docx` |
| `report_title` | Portuguese title displayed in the report |
| `report_language` | Word document language, normally `pt-PT` |
| `processed_directory` | Destination for successful PDF copies |
| `error_directory` | Destination for failed PDFs and diagnostic JSON |
| `currency` | Fallback applied to register rows with no currency; the current extraction contract and validation require `EUR` |
| `provider_labels` | Model/legacy provider-label mappings |

Writable paths must remain under `/files`, the data mount authorized by `N8N_RESTRICT_FILE_ACCESS_TO`. The current Compose configuration also authorizes `/config` so n8n can read the three read-only configuration mounts.

## Prompt and schema

`prompts/invoice-extractor.md` becomes the agent system message on every execution. `schemas/invoice.schema.json` is appended to that message as the expected response contract. Changes to either file take effect on the next run.

Keep the schema and the workflow normalization logic compatible when adding or removing fields.

The current schema fixes `amounts.currency` to `EUR`, and the normalization and validation logic enforce the same value. Supporting another currency requires coordinated changes to the schema and workflow; changing only `currency` in `config/extraction.json` is not sufficient.

## Read-only mounts

```yaml
- ./config:/config/runtime:ro
- ./prompts:/config/prompts:ro
- ./schemas:/config/schemas:ro
```

The n8n container can read but cannot modify these repository files.

## Settings that cannot be runtime files

- The one-minute Schedule Trigger interval is part of the workflow definition because it is evaluated before execution begins.
- Ollama credentials and base URL remain in the n8n credential store.
- Word layout and Portuguese report structure are implemented by `report-service/app.py`.
- Docker mounts, filesystem permissions, timezone, and encryption key remain in Compose/`.env`.

After changing Compose mounts or environment settings, recreate the container. Runtime JSON, prompt, and schema changes require no restart or workflow reimport.
