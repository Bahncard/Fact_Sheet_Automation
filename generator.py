import pandas as pd
from pathlib import Path

def clean_data(input_dir=None, output_dir=None):
    """
    Clean supplier data by filtering specific categories from input Excel files.
    
    Args:
        input_dir (Path, optional): Directory containing input files. Defaults to 'mock_tables'.
        output_dir (Path, optional): Directory for output files. Defaults to 'clean_tables'.
    """
    # If no arguments provided, input_dir and output_dir will be None

    # Set up default input / output dirs  when no parameters received
    if input_dir is None:
        # Specify the obejct representing the relative path "mock_tables"
        input_dir = Path("mock_tables")
    if output_dir is None:
        output_dir = Path("clean_tables")
    
    # Create output dir if it doesn't exist
    output_dir.mkdir(exist_ok= True)

    # Load the data
    it_spend = pd.read_excel(input_dir / "mock_Supplier_fact_sheet_IT_spend_2024.xlsx")
    contracting_report = pd.read_excel(input_dir / "mock_Supplier_fact_sheet_Contracting_report.xlsx")
    sourcing_event = pd.read_excel(input_dir / "mock_Supplier_fact_sheet_sourcing_event_participation.xlsx")

    # Filter data based on the required categories
    filtered_it_spend = it_spend[it_spend["Category"].isin(["IT Hardware", "Telecoms and Network"])]
    filtered_contracting_report = contracting_report[
        contracting_report["[PCW] OneProcurement Category"] == "IT Infrastructure"
    ]
    filtered_sourcing_event = sourcing_event[
        sourcing_event["[SPRJ] OneProcurement Category"] == "IT Infrastructure"
    ]

    # Save the cleaned data
    filtered_it_spend.to_excel(output_dir / "cleaned_Supplier_fact_sheet_IT_spend_2024.xlsx", index=False)
    filtered_contracting_report.to_excel(output_dir / "cleaned_Supplier_fact_sheet_Contracting_report.xlsx", index=False)
    filtered_sourcing_event.to_excel(output_dir / "cleaned_Supplier_fact_sheet_sourcing_event_participation.xlsx", index=False)
    
    return filtered_it_spend, filtered_contracting_report, filtered_sourcing_event

def extract_unique_vendors(it_spend, contracting_report, sourcing_event):
    """
    Extract a list of unique vendors across the three datasets.
    
    Args:
        it_spend (DataFrame): Cleaned IT spend data.
        contracting_report (DataFrame): Cleaned contracting report data.
        sourcing_event (DataFrame): Cleaned sourcing event data.

    Returns:
        set: A set of unique vendor names.
    """
    # Extract unique vendors from IT spend data
    vendors_it_spend = set(it_spend["Vendor Name"].unique())
    
    # Extract unique vendors from Contracting Report
    vendors_contracting = set(contracting_report["[PCW]Affected Parties (Supplier Name (L1))"].unique())
    
    # Extract unique vendors from Sourcing Event data
    vendors_sourcing = set(sourcing_event["[SPT]Supplier (Supplier Name (L1))"].unique())
    
    # Combine all unique vendors into a single set to remove duplicates
    unique_vendors = vendors_it_spend.union(vendors_contracting, vendors_sourcing)
    
    return unique_vendors

def simulate_historical_spend(it_spend):
    """
    Simulate spend data for 2022 and 2023 based on 2024 spend.
    
    Args:
        it_spend (DataFrame): Cleaned IT spend data with 2024 spend.

    Returns:
        DataFrame: Updated DataFrame with simulated 2022, 2023 spend.
    """
    # Add 2022 and 2023 spend columns
    simulated_data = it_spend.copy()
    simulated_data["Spend 2023 (€m)"] = simulated_data["Spend 2024 (€m)"] * (1 + simulated_data.apply(lambda x: random.uniform(-0.2, 0.2), axis=1))
    simulated_data["Spend 2022 (€m)"] = simulated_data["Spend 2023 (€m)"] * (1 + simulated_data.apply(lambda x: random.uniform(-0.2, 0.2), axis=1))
    return simulated_data


def main():
    # Clean the data
    it_spend, contracting_report, sourcing_event = clean_data()
    print("Data cleaning completed!")

    # Extract the unique vendors
    unique_vendors = extract_unique_vendors(it_spend, contracting_report, sourcing_event)
    print(f"Unique vendors extracted: {len(unique_vendors)}")
    print(unique_vendors)

if __name__ == "__main__":
    main()