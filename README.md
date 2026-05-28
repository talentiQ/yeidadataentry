# YEIDA Flask Automation Agent

This Flask app lets you upload all images for **one YEIDA applicant** at a time. It extracts the details with the OpenAI API, applies your fixed rules, and appends the applicant to the next row in the master Excel file. It also regenerates a master XML file after every upload.

## What it does

- Upload multiple images for one applicant in one form submission.
- Extracts application number, applicant name, plot size, address, bank details, PAN, mobile number, receipt details, and handwritten cover-page values.
- Uses the **handwritten / blue-circled receipt number first**.
- Marks employment as **SALARIED** by default.
- Uses **Account Type 31** by default, with dropdown override for 10 or 29.
- Calculates YEIDA asset cost and amount financed from plot size/category.
- Saves every applicant to:
  - `data/YEIDA_MASTER.xlsx`
  - `data/YEIDA_MASTER.xml`
  - `data/YEIDA_MASTER.json`

## Setup

```bash
cd yeida_flask_agent
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env
```

Open `.env` and add your API key:

```env
OPENAI_API_KEY=your_api_key_here
```

## Run

```bash
python app.py
```

Then open:

```text
http://127.0.0.1:5000
```

## How to use

1. Upload all images for one applicant together.
2. Choose Account Type only if it is not default 31.
3. Use Receipt No Override only when the handwritten/circled receipt number is unclear.
4. Click **Extract and add next row**.
5. Download Excel/XML from the result page.
6. Upload the next applicant; it will append to the next row.

## Important notes

- If Aadhaar is not visible, the app leaves ID proof blank.
- If Aadhaar is visible, the prompt asks the model to output only the last digit.
- If category is SC/ST, pricing uses the SC/ST YEIDA table. Otherwise it uses the general category table.
- The app does not delete old rows. It always appends.

## Files

```text
app.py              Flask routes and upload flow
agent.py            OpenAI image extraction + business rules
storage.py          Excel/XML/JSON storage
requirements.txt    Python dependencies
templates/          HTML pages
data/               Generated Excel/XML/JSON files
uploads/            Uploaded image batches
```
