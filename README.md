# InternshipContactor

## Overview

A simple Python project to automate sending resumes to multiple companies using a touch of AI. This project demonstrates how to integrate OpenAI's API and improve Python skills while managing job applications.

---

## Requirements

Your `.env` file needs to contain the following:
1. `OPENAI_API_KEY={Your OpenAI API Key}`
2. `EMAIL_USER={Your Google Email}`
3. `EMAIL_PASSWORD={Google App Password}`

To get your Google email password, please visit:
[Generate Google App Password](https://myaccount.google.com/apppasswords)

---

## Setup Instructions

### Update the Following in the Code
Search for `#TODO` in the code and update:
1. **Your Name**: Enter your full name where required.
2. **Path to Excel File**: Provide the path to your Excel file containing the following columns:
   - `Company_Name`
   - `Email`
3. **Path to Your CV**: Specify the path to your CV (in PDF format).  
   **Note:** This project works only with single-page CVs.

---

## Poppler Installation

If you encounter difficulties with `poppler`, install it using the following command:

```bash
choco install poppler
