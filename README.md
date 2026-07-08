# iAIR PDF Tools - Streamlit App

This repository contains one internal Streamlit app for Excel-to-PDF generation.

## Active tools

1. Assignment PDF Generator
   - Expected Excel sheet: `Assignments`
   - Template: `tools/assignment/Assignment_Input_Template.xlsx`

2. Assessment PDF Generator
   - Expected Excel sheet: `Assessments`
   - Template: `tools/assessment/Assessment_Input_Template.xlsx`

3. Project PDF Generator
   - Expected Excel sheet: `Projects`
   - Template: `tools/project/Project_Input_Template.xlsx`

4. Question Bank Generator
   - Status: Coming soon placeholder

## Streamlit Cloud setup

1. Upload this repository to GitHub.
2. Create a new app on Streamlit Community Cloud.
3. Select:
   - Branch: `main` or `master`
   - Main file: `app.py`
   - Python version: `3.12`
4. Add this in Streamlit secrets:

```toml
APP_PASSWORD = "your-secure-password"
```

5. Deploy or reboot the app.

## Important notes

- First PDF generation may take extra time because Playwright Chromium may be prepared.
- Use the correct Excel template for the selected tool. The app validates the expected sheet before generating PDFs.
- Generated files are created in a temporary folder and returned as a ZIP download.

## Local run

```bash
pip install -r requirements.txt
python -m playwright install chromium
streamlit run app.py
```
