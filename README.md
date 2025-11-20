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

Set the following environment variables before launching the UI:

```
export AZURE_OPENAI_API_KEY=...        # Required
export AZURE_OPENAI_ENDPOINT=...       # Required
export AZURE_OPENAI_DEPLOYMENT=...     # Required (deployment/model name)
export AZURE_OPENAI_API_VERSION=2024-02-15-preview  # Optional override
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
