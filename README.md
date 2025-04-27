# Usage

It is recommended to use virtual environment because handling COM object needs quite fine version handling.

```bash
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

Now you can run this

```bash
python main.py "<folder-path-of-hwpx-file>"
```

This command reads all files with valid extensions(.hwp, .hwpx)
and convert those to pdfs,
and extracts the first pages of them,
and merge them.
