# iAIR PDF Tools - Streamlit App

This is a simple internal Streamlit portal for Excel-to-PDF automation.

Currently included tools:

1. Assignment PDF Generator
2. Assessment PDF Generator

Future placeholders are already shown for:

- Project PDF Generator
- Question Bank Generator

## Folder Structure

```text
iair_pdf_tools_streamlit/
├── app.py
├── requirements.txt
├── packages.txt
├── runtime.txt
├── .streamlit/
│   ├── config.toml
│   └── secrets.example.toml
└── tools/
    ├── assignment/
    │   ├── generate_pdfs.py
    │   ├── assignment_template.html
    │   ├── Assignment_Input_Template.xlsx
    │   └── assets/iair_logo.png
    └── assessment/
        ├── generate_assessments.py
        ├── assessment_template.html
        ├── Assessment_Input_Template.xlsx
        └── assets/iair_logo.png
```

## Deploy on Streamlit Community Cloud

1. Create a new GitHub repository.
2. Upload all files from this folder to the repository root.
3. Go to Streamlit Community Cloud.
4. Create a new app from your GitHub repository.
5. Set the main file path as:

```text
app.py
```

6. In Streamlit Cloud app settings, add this secret:

```toml
APP_PASSWORD = "your-secure-password"
```

7. Deploy the app.

## How to Use

1. Open the Streamlit app URL.
2. Enter the password.
3. Select Assignment or Assessment generator.
4. Download the Excel template if needed.
5. Fill the Excel template.
6. Upload the completed `.xlsx` file.
7. Click `Generate PDF ZIP`.
8. Download the generated ZIP file.

## Important Notes

- Only `.xlsx` uploads are allowed.
- Maximum upload size is configured as 25 MB.
- Generated files are temporary and are not stored permanently.
- The app uses Playwright + Chromium for PDF generation.
- `packages.txt` installs system Chromium for Streamlit Cloud.

## Local Run

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

For local password protection, either set an environment variable:

```bash
export APP_PASSWORD="your-password"
```

or create `.streamlit/secrets.toml`:

```toml
APP_PASSWORD = "your-password"
```

Do not commit `.streamlit/secrets.toml` to GitHub.
