# Azure Oral Assessment Coding App

This Tkinter desktop app loads `.docx` interview transcripts, sends them to an
Azure OpenAI deployment for coding, and color-codes the transcript according to
the following categories:

1. **Background & Context**
2. **Feasibility & Practical Implementation**
3. **Validity & Learning Assurance**
4. **Disciplinary Relevance**
5. **Student Engagement & Observations**
6. **Reflection & Improvement**
7. **Sustainability & Future Use**
8. **Additional Insights**

The highlighted document preserves the original transcript text, only adding
color coding to the relevant passages.

## Requirements

* Python 3.10+
* `pip install -r requirements.txt`
* Azure OpenAI resource with a deployed model such as GPT-4o mini

### Recommended: Use a virtual environment

Isolate the app dependencies to avoid conflicts with other projects on your
machine:

```bash
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\\Scripts\\activate
pip install -r requirements.txt
```

After activating, run the commands below in the same shell so they use the
virtual environment's Python and packages.

Set the following environment variables before launching the UI. On Windows
Command Prompt, use `set VAR=value`; in PowerShell, use `$env:VAR="value"`.

```bash
# macOS/Linux
Set the following environment variables before launching the UI:

```
export AZURE_OPENAI_API_KEY=...        # Required
export AZURE_OPENAI_ENDPOINT=...       # Required
export AZURE_OPENAI_DEPLOYMENT=...     # Required (deployment/model name)
export AZURE_OPENAI_API_VERSION=2024-02-15-preview  # Optional override

# Windows (Command Prompt)
set AZURE_OPENAI_API_KEY=...
set AZURE_OPENAI_ENDPOINT=...
set AZURE_OPENAI_DEPLOYMENT=...
set AZURE_OPENAI_API_VERSION=2024-02-15-preview

# Windows (PowerShell)
$env:AZURE_OPENAI_API_KEY="..."
$env:AZURE_OPENAI_ENDPOINT="..."
$env:AZURE_OPENAI_DEPLOYMENT="..."
$env:AZURE_OPENAI_API_VERSION="2024-02-15-preview"
```

### Using a `.env` file

You can keep your Azure values in a local `.env` file instead of exporting
variables each time. Create a file named `.env` next to `app.py` with:

```
AZURE_OPENAI_API_KEY=...
AZURE_OPENAI_ENDPOINT=...
AZURE_OPENAI_DEPLOYMENT=...
AZURE_OPENAI_API_VERSION=2024-02-15-preview
```

The app automatically loads this file on startup. Never commit your real keys;
add `.env` to `.gitignore` if you store it locally.

```

## Running the app

```
python app.py
```

Steps:

1. Choose the input transcript (`.docx`).
2. Choose where to save the highlighted output file.
3. Click **Process Transcript**. Large transcripts are automatically chunked.
4. The status window displays model progress, and the coded file is saved to the
   selected location when finished.

If no matches are returned, the original transcript is saved without changes.
