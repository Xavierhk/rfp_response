# RFP Response Automation Tool

This tool automates the process of extracting tables from Word documents (.docx) and converting them into Excel format. It's particularly useful for processing RFP (Request for Proposal) documents and organizing their content.

## Features

- Extracts tables from Word documents (.docx)
- Filters out placeholder text ("Insert your response here...")
- Combines multiple tables into a single Excel file
- Handles duplicate column names automatically
- Preserves table source information

## Prerequisites

Before running this tool, make sure you have Python installed and the following packages:

```bash
pip install python-docx pandas openpyxl
```

## Installation

1. Clone the repository:
```bash
git clone https://github.com/YOUR_USERNAME/rfp_response.git
cd rfp_response
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Place your Word document (containing tables) in the project directory
2. Run the script:
```bash
python export2.py
```

The script will:
- Read tables from `questions.docx`
- Process and combine the tables
- Generate an output Excel file named `combined_tables.xlsx`

## File Structure

rfp_response/

├── export2.py # Main script for table extraction

├── questions.docx # Input Word document (your RFP document)

├── requirements.txt # Project dependencies

└── README.md # This file

## Contributing

Feel free to submit issues and enhancement requests!

## License

[MIT](https://choosealicense.com/licenses/mit/)
