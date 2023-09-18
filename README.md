# Note
To see your own results after running the script -> change the excell name when you save the excell file. It is in the bottom of the code

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
```

## Usage

1. Open a terminal or command prompt.

2. Navigate to the directory containing the script.

3. Run the script with the following command, replacing 'your_api_key_here' with your actual OpenAI API key:

```bash
python extract.py  'your_api_key_here'
```

4. The script will process the PDF, extract information, make an API request, and create an Excel file with the results.


## Output 

The script will generate an Excel file named my_sample_reply.xlsx, which will contain the following data:

- Fund name
- Whether the information is related to "Article 8" or "Article 9"
- Source page and content related to the fund
- ESG strategy bullet points generated by the OpenAI API

## Disclaimer

Please note that this script is a sample and may require further customization to work with your specific PDF document and data requirements. The OpenAI API usage may also depend on your subscription and API key availability.

Use this script responsibly and ensure compliance with any legal and ethical considerations related to data extraction and usage.

Happy data extraction and analysis!

```vbnet

Make sure to replace `'your_api_key_here'` with your actual OpenAI API key in the Usage section. Additionally, customize the README as needed to provide more specific instructions or additional context for users of your script.

```
