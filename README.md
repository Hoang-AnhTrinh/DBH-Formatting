# DBH-Formatting

DBH-Formatting is a Codex skill for building a local, JSON-first workflow around Japanese structural PDFs and Excel templates. It classifies drawing pages, selectively renders pages that need deeper inspection, prepares reviewed project data, and writes confirmed values into an Excel template while preserving OOXML controls.

Current version: **1.0.2**

## Install in Codex

Add this repository as a plugin marketplace:

```powershell
codex plugin marketplace add Hoang-AnhTrinh/DBH-Formatting --ref main
```

Restart the ChatGPT desktop app, open the Plugins Directory, select the **DBH-Formatting** marketplace, and install the plugin. Start a new conversation before using the skill for the first time.

You can then invoke it explicitly with:

```text
$dbh-formatting
```

### Install only the skill

If you do not want the plugin wrapper, ask Codex to install the skill from:

```text
https://github.com/Hoang-AnhTrinh/DBH-Formatting/tree/main/plugins/dbh-formatting/skills/dbh-formatting
```

## Runtime requirements

- Windows PowerShell for the bundled pipeline runner.
- Python 3 with `pypdf`.
- Poppler `pdftoppm` for rendering selected PDF pages. Text extraction still works without it, but images will not be rendered.
- Node.js and `@oai/artifact-tool` only when the optional Excel-template inspection stage is enabled.

Install the Python dependency in the Python environment used by the runner:

```powershell
python -m pip install -r plugins/dbh-formatting/skills/dbh-formatting/scripts/requirements.txt
```

Copy `references/pipeline_config.example.json` to a project-owned location, replace its placeholder paths, and run:

```powershell
powershell -ExecutionPolicy Bypass -File "<skill-folder>\scripts\run_pipeline.ps1" -ConfigPath "<path-to-config.json>"
```

The skill never overwrites source PDFs or Excel templates. Keep customer documents, generated workbooks, local paths, credentials, and project-specific configurations out of the repository.

## Repository layout

```text
.
|-- .agents/plugins/marketplace.json
`-- plugins/dbh-formatting/
    |-- plugin.json
    `-- skills/dbh-formatting/
        |-- SKILL.md
        |-- agents/openai.yaml
        |-- references/
        `-- scripts/
```

## Tiếng Việt

Đây là bản phân phối công khai của skill DBH-Formatting. Repo không chứa bản vẽ, Excel mẫu của khách hàng, đường dẫn máy cá nhân hay dữ liệu dự án nội bộ. Trước khi chạy, hãy sao chép file cấu hình mẫu ra thư mục dự án và thay các đường dẫn placeholder bằng file của chính dự án đó.

## License

MIT License. See [LICENSE](LICENSE).

