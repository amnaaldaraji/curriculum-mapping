# Curriculum Mapping Tool

A web-based survey for teachers to map their courses against SDG-related competences. Built for the biomedical programmes at Karolinska Institutet.

Teachers fill in their course details, rate how each learning outcome is addressed, and submit. Responses are saved automatically to an Excel sheet via Power Automate.

## How it works

1. Teacher opens the tool and accepts a GDPR consent form
2. They fill in their course name, programme, and term
3. For each of the 10 SDG-related learning outcomes, they rate their course (Introductory, Advanced, or Masters level) and describe their teaching activities and assessments
4. On submission, the data is sent to a Power Automate HTTP flow that writes it to a shared Excel file

## Built with

- Plain HTML, CSS and JavaScript, no frameworks
- Google Fonts (EB Garamond + DM Sans)
- jsPDF for PDF export
- Power Automate for backend data collection
- Google Apps Script (Code.gs) for an earlier Sheets integration

## Structure

| File | Description |
|------|-------------|
| `index.html` | The entire app, single file |
| `Code.gs` | Google Apps Script for Sheets integration |
