import pandas as pd
import random
import numpy as np
from pathlib import Path

# Set the output directory
current_working_dir = Path.cwd()
output_dir = current_working_dir / "mock_tables"

# List of 50 IT Infrastructure vendors
known_vendors = [
    "AWS", "Microsoft", "Google", "Oracle", "IBM", "Dell", "Cisco", "HP", "VMware", "Intel",
    "Lenovo", "Fujitsu", "Juniper Networks", "SAP", "Adobe", "Salesforce", "Alibaba Cloud", "Tencent Cloud", "HPE",
    "Red Hat", "NetApp", "Nutanix", "Arista Networks", "Fortinet", "Palo Alto Networks", "Cloudflare", "Rackspace",
    "Equinix", "Atlassian", "ServiceNow", "Snowflake", "Cloudera", "Tableau", "Datadog", "Splunk", "Zoom",
    "Slack", "Dropbox", "Box", "Okta", "Ping Identity", "VMware Tanzu", "Citrix", "Extreme Networks", "Sophos",
    "Barracuda Networks", "Zscaler", "Veritas", "Commvault","MongoDB", "Elastic", "GitHub" 
]

# Function to generate mock data for Supplier fact sheet_IT spend 2024
def generate_it_spend():
    categories = ["IT Hardware", "Telecoms and Network"]

    data = {
        "Vendor Name": known_vendors, 
        "Category": [random.choice(categories) for _ in range(len(known_vendors))],
        "Spend 2024 (€m)": [round(random.uniform(1, 100), 2) for _ in range(len(known_vendors))]
    }


    df = pd.DataFrame(data)
    output_path = output_dir / "mock_Supplier_fact_sheet_IT_spend_2024.xlsx"
    df.to_excel(output_path, index=False)
    return df

# Function to generate mock data for Supplier fact sheet_Contracting report
def generate_contracting_report():
    contract_ids = [f"CW{random.randint(10000, 99999)}" for _ in range(200)]
    contracts = [f"Contract {i}" for i in range(1, 201)]
    descriptions = ["Master Service Agreement", "Enterprise Agreement", "Cloud Services", "IT Support"]
    categories = ["IT Infrastructure", "Software", "Consulting"]

    weights = [1] * len(known_vendors)  # All vendors have equal weight

    data = {
        "[PCW] Contract Id": contract_ids,
        "[PCW]Contract (Contract)": contracts,
        "[PCW] Description": [random.choice(descriptions) for _ in range(200)],
        "[PCW]Contract (Effective Date)": pd.date_range(start="2020-01-01", periods=200, freq="D"),
        "[PCW]Contract (Expiration Date)": pd.date_range(start="2025-01-01", periods=200, freq="D"),
        "[PCW] OneProcurement Category": [random.choice(categories) for _ in range(200)],
        "sum(Contract Amount) (€m)": [round(random.uniform(10, 500), 2) for _ in range(200)],
        "[PCW]Affected Parties (Supplier Name (L1))": [random.choice(known_vendors) for _ in range(200)]
    }

    df = pd.DataFrame(data)
    output_path = output_dir / "mock_Supplier_fact_sheet_Contracting_report.xlsx"
    df.to_excel(output_path, index=False)
    return df

# Function to generate mock data for Supplier fact sheet_sourcing event participation
def generate_sourcing_event():
    project_ids = [f"PRJ{random.randint(1000, 9999)}" for _ in range(200)]
    project_names = [f"Project {i}" for i in range(1, 201)]
    categories = ["IT Infrastructure", "Software", "Consulting"]

    weights = [1] * len(known_vendors)  # All vendors have equal weight

    data = {
        "[SPRJ]Project (Project Id)": project_ids,
        "[SPRJ]Project (Project Name)": project_names,
        "[SPRJ] OneProcurement Category": [random.choice(categories) for _ in range(200)],
        "sum(Baseline Spend) (€m)": [round(random.uniform(5, 200), 2) for _ in range(200)],
        "[SPT]Supplier (Supplier Name (L1))": [random.choice(known_vendors) for _ in range(200)],
        "Short Description": [f"Description for Project {project_id}" for project_id in project_ids]
    }

    df = pd.DataFrame(data)
    output_path = output_dir / "mock_Supplier_fact_sheet_sourcing_event_participation.xlsx"
    df.to_excel(output_path, index=False)
    return df

# Generate the mock data for all three tables
it_spend = generate_it_spend()
contracting_report = generate_contracting_report()
sourcing_event = generate_sourcing_event()

print("Mock data for Supplier fact sheets generated successfully!")
