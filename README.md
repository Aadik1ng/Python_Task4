# Personalized Document Generator

A Streamlit-based application to generate personalized PDF and Word documents from employee data. Users can either upload one or more Excel files (each containing employee details) or enter the details manually through a form. The generated documents are saved in an `output` folder.

## Features

- **Multiple Input Options:**
  - **Excel Upload:** Upload one or more Excel files (`.xlsx` or `.xls`) containing employee data.
  - **Manual Entry:** Fill in employee data manually via an interactive form.
  
- **Document Generation:**
  - Generate personalized **Word** documents using `python-docx`.
  - Generate personalized **PDF** documents using `ReportLab`.
  
- **Customizable Template:**
  - Includes a company logo (if available) and a welcome message with placeholders for employee details.
  
- **Output:**
  - Generated documents are saved in the `output` directory.

## Prerequisites

Ensure you have Python 3.7 or later installed. You will also need to install the following Python packages:

- Streamlit
- pandas
- python-docx
- reportlab
- openpyxl (for Excel file handling)

## Installation

> **Clone this repository or download the script:**
> 
> ```bash
> git clone https://github.com/yourusername/personalized-document-generator.git
> cd personalized-document-generator
> ```

> **Create a virtual environment (optional but recommended):**
> 
> ```bash
> python -m venv venv
> source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
> ```

> **Install the required packages:**
> 
> ```bash
> pip install streamlit pandas python-docx reportlab openpyxl
> ```

## Usage

> **Place Your Logo (Optional):**
> 
> Save your company logo as `logo.png` in the same folder as the script if you wish to include it in the generated documents.

> **Run the Streamlit App:**
> 
> ```bash
> streamlit run your_script.py
> ```
> Replace `your_script.py` with the actual name of your script file.

> **Using the Application:**
> 
> **Excel Upload:**
> 
> 1. Choose the "Upload Excel File" option.
> 2. Upload one or more Excel files containing the required columns:
>    - `Name`
>    - `Email`
>    - `Company Name`
>    - `Position`
>    - `Joining Date`
> 3. Click the **Generate Documents** button to process the files and generate the documents.
> 
> **Manual Entry:**
> 
> 1. Select the "Manual Entry" option to add employee details one by one.
> 2. Use the form to fill in the data and add entries to the list.
> 3. When ready, click **Generate Documents** to create the documents.

> **Output:**
> 
> Generated files are stored in the `output/` folder. You will see a list of generated file names on the Streamlit interface.

## Excel File Format

Ensure your Excel files contain the following columns (headers):

- **Name:** Employee's full name.
- **Email:** Employee's email address.
- **Company Name:** Name of the company.
- **Position:** Employee's job position.
- **Joining Date:** Date of joining (ensure the date is in a recognizable date format).
