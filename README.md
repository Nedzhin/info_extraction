# PDF Data Extraction and Analysis Script

This Python script is designed to extract specific information from a PDF file, analyze the extracted data, and write the results to an Excel file. It utilizes various libraries, including PyPDF2, openpyxl, and the OpenAI API.

## Overview

The script performs the following tasks:

1. **Finding Page Interval:** It locates the page interval containing information about a special fund within the PDF file. The page interval is determined by specifying the starting and ending headings.

2. **Getting Full Text and Information:** It extracts the full text within the specified page interval and identifies whether the content is related to "Article 8" or "Article 9" by searching for specific keywords.

3. **Using the OpenAI API:** It uses the OpenAI API to generate a response based on the extracted text and a prompt requesting key bullet points related to the ESG strategy.

4. **Excel File Creation:** It creates an Excel file and populates it with relevant data, including fund information, ESG bullet points, and source page details.

## Prerequisites

Before running the script, make sure you have the following dependencies installed:

- `PyPDF2`
- `openpyxl`
- `openai`

You can install them using pip:

```bash
pip install PyPDF2 openpyxl openai
