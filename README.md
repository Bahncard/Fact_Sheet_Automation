# Fact Sheet Automation

A tool for automating PowerPoint reports and vendor fact sheet generation.

## Description

This project demonstrates how to automate the generation of PowerPoint reports and fact sheets using Python.

## Prerequisites

- Python 3.x
- Required Python packages listed in requirements.txt

## Setup and Usage

### 1. Generate Mock Data

```console
cd mock_tables
python mock_tables
```

This will create mock data for 50 vendors in the `/mock_tables` directory.

### 2. Vendor Information

- Vendor general information is stored in `all_vendors_data.json`
- Ensure this file is properly configured before proceeding

### 3. Generate Fact Sheets

```console
python generator.py
```

**Note**: Ensure that both the tables and `all_vendors_data.json` are created before running the generator.

### 4. Optional: LLM Summary Generation

To generate AI-powered summaries of vendor information:

```console
python vendor_data_generator.py
```
