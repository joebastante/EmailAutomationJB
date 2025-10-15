# Outlook Email Generator Using OpenAI

I've created a Python program that will process your Excel workbook and generate personalized emails[cite: 2]. Here are the key features:

---

## What the program does:

* Opens the **Excel workbook** with contact information [cite: 4]
* Reads the **"Working" worksheet** starting from row 2 [cite: 5]
* For each row, extracts: **name** (Column A), **email type** (Column B), **email address** (Column C), and **context** (Column D) [cite: 6]
* Separates **first name** from the full name [cite: 7]
* Calls **OpenAI's API** (using **gpt-5-mini** model) to generate personalized email text [cite: 8]
* Opens a new email window with the recipient, subject **"Checking in"**, and generated body [cite: 9]
* After opening the new Outlook message window it moves to the next row [cite: 10]

---

## Setup Requirements:

* Install required packages:
    ```bash
    pip install openpyxl openai
    ```
* Set your **OpenAI API key** as an environment variable:
    * Windows: `set OPENAI_API_KEY=your-api-key` [cite: 16]
    * Mac/Linux: `export OPENAI_API_KEY='your-api-key'` [cite: 17]
* Update the **workbook path** in the code (line 113) to match your file location [cite: 18]

---

## Important Notes:

* The model name is **"gpt-5-mini"** - this is OpenAI's current efficient model [cite: 20]
* The program is **OS-agnostic** and will work on Windows, macOS, and Linux [cite: 21]
* On Windows, it attempts to use Outlook directly, but falls back to the default mail client if needed [cite: 22]
* The program **pauses between emails** so you can review each one before proceeding [cite: 23]
* Emails are **NOT sent automatically** - you'll review and send them manually [cite: 24]