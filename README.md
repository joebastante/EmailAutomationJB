# Outlook Email Generator Using OpenAI

I've created a Python program that will process your Excel workbook and generate personalized emails opening each in a new Outlook email window. Here are the key features:

---

## What the program does:

* Opens the **Excel workbook** with contact information
* Reads the **"Working" worksheet** starting from row 2 
* For each row, extracts: **name** (Column A), **email type** (Column B), **email address** (Column C), and **context** (Column D) 
* Separates **first name** from the full name 
* Calls **OpenAI's API** (using **gpt-5-mini** model) to generate personalized email text 
* Opens a new email window with the recipient, subject **"Checking in"**, and generated body
* After opening the new Outlook message window it moves to the next row 

---

## Setup Requirements:

* Install required packages:
    ```bash
    pip install openpyxl openai
    ```
* Set your **OpenAI API key** as an environment variable:
    * Windows: `set OPENAI_API_KEY=your-api-key` 
    * Mac/Linux: `export OPENAI_API_KEY='your-api-key'` 
* Update the **workbook path** in the code (line 113) to match your file location 

---

## Important Notes:

* The model name is **"gpt-5-mini"** - this is OpenAI's current efficient model 
* The program is **OS-agnostic** and will work on Windows, macOS, and Linux 
* On Windows, it attempts to use Outlook directly, but falls back to the default mail client if needed 
* The program **pauses between emails** so you can review each one before proceeding 
* Emails are **NOT sent automatically** - you'll review and send them manually
