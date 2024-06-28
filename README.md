# DOCX Comparison Tool

## Description

The DOCX Comparison Tool is an advanced utility designed to meticulously compare two DOCX files, providing both a similarity percentage and a detailed PDF report that highlights any discrepancies. This tool identifies and reports a wide range of errors, making it invaluable for proofreading and quality control. Key features include:

- **Missing Punctuation:**
  - Commas
  - Periods (dots)
  - Exclamation marks
  - Single inverted commas
  - Double inverted commas
- **Incorrect Spacing:**
  - Extra or missing spaces
- **Typographical Errors:**
  - Spelling mistakes

Ideal for ensuring the accuracy of typed documents against their original versions, this application is perfect for anyone needing to maintain high standards of document fidelity.

## Outputs 
### GUI Output 
<img src="https://github.com/codeterrayt/DOCX-Comparison-Tool/assets/76623307/b80ffd41-ab59-421b-b0df-a363eb07a3bc" width=400 />

### PDF REPORT OUTPUT 
<img src="https://github.com/codeterrayt/DOCX-Comparison-Tool/assets/76623307/6782962e-9eaf-439e-9dea-d6585db29376" width=400 />

## Installation Guide

### Prerequisites

Before you begin, ensure that Python 3.12.1 is installed on your system. If not, download it from the official [Python website](https://www.python.org/downloads/).

### Installation Steps

1. **Clone the Repository**

    Begin by cloning the repository to your local machine:

    ```sh
    git clone https://github.com/codeterrayt/DOCX-Comparison-Tool.git
    cd DOCX-Comparison-Tool
    ```
2. **Create and Activate a Virtual Environment (venv)**

    It is recommended to use a virtual environment to manage dependencies. Follow these steps:

    - **Create a virtual environment:**

        ```sh
        python3 -m venv venv
        ```

    - **Activate the virtual environment:**

        - On Windows:

            ```sh
            venv\Scripts\activate
            ```

        - On macOS/Linux:

            ```sh
            source venv/bin/activate
            ```

3. **Install Requirements**

    Next, install all necessary packages from the `requirements.txt` file using pip:

    ```sh
    pip3 install -r requirements.txt
    ```

4. **Run the Main Script**

    Finally, execute the main script to launch the application:

    ```sh
    python main.py
    ```

## Usage Instructions

Using the DOCX Comparison Tool is straightforward. Follow these steps:

1. **First Input:** Select the actual DOCX file.
2. **Second Input:** Select the written or typed DOCX file.

Once the inputs are provided, the application will:

1. Compare the two DOCX files.
2. Calculate and display the similarity percentage.
3. Generate a comprehensive PDF report detailing errors such as missing punctuation, incorrect spacing, and typographical mistakes.

This powerful tool ensures your documents are accurate and professionally formatted, saving you time and improving your workflow.
