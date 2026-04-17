# Select by G — Supplier Presentation Generator

A Flask web app that generates branded PPTX presentations for hotel furniture suppliers.

## Features

- 6-step form: Identity, History, Projects, Values & References, Photos, Generate
- Upload supplier logo (replaces template logo on all slides, aspect ratio preserved)
- 2 supplier photos placed directly on the resume slide (slide 1)
- Hotel reference photos each get their own slide (slide 2, 3, …)
- Output saved directly to the Google Drive folder
- Windows double-click launcher (`Launch App.bat`)

## Requirements

- Python 3.8+
- `flask`, `lxml`, `Pillow` (auto-installed by the launcher)

## Quick Start (Windows)

1. Double-click **`Launch App.bat`**
2. The browser opens automatically at `http://localhost:5001`
3. Fill in the form and click **Generate Presentation**

## Manual Start

```bash
pip install flask lxml Pillow
python app.py
```

Then open `http://localhost:5001` in your browser.

## Template

The app expects the PPTX template (`Présentation Fournisseurs Template EN.pptx`) to be present in the `Select by G Group Présentations` Google Drive folder. The generated file is saved to the same folder.
