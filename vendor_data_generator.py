import json
import os
from dotenv import load_dotenv
from openai import OpenAI
import time
from pathlib import Path

# Load environment variables
load_dotenv()

# Initialize OpenAI client with DeepSeek base URL
DEEPSEEK_API_KEY = os.getenv('DEEPSEEK_API_KEY')
client = OpenAI(
    api_key=DEEPSEEK_API_KEY,
    base_url="https://api.deepseek.com"
)

def get_vendor_financials(vendor_name):
    """
    Fetch financial information for a vendor using DeepSeek API
    """
    prompt = f"""
    Provide financial information for {vendor_name}.
    Start your response with {{ and respond ONLY with a raw JSON object in this exact format:
    {{
        "Revenue": "<number> USD",
        "MarketCap": "<number> USD",
        "GrowthRate": "<number>%"
    }}
    Do not include any markdown formatting, code blocks, or additional text.
    """
    
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "You are a financial data assistant. Provide concise, accurate information in JSON format."},
                {"role": "user", "content": prompt}
            ],
            stream=False
        )
        # Print the raw API response
        print(f"Raw API response for {vendor_name}:")
        print(response.choices[0].message.content)


        # Try to parse the response as JSON
        try:
            return json.loads(response.choices[0].message.content)
        
        #Inner layer of try-except:handle JSON parsing error
        except json.JSONDecodeError as je:
            print(f"JSON parsing error: {str(je)}")
            print(f"Response content: {response.choices[0].message.content}")
            return {
                "Revenue": "Data not available",
                "MarketCap": "Data not available",
                "GrowthRate": "Data not available"
            }
        
    #Outer layer of try-except:handle API call error
    except Exception as e:
        print(f"API call error for {vendor_name}: {str(e)}")
        return {
            "Revenue": "Data not available",
            "MarketCap": "Data not available",
            "GrowthRate": "Data not available"
        }
def get_market_trends(vendor_name):
    """
    Fetch market trends for a vendor using DeepSeek API
    """
    prompt = f"""
    Provide 3 key current market trends for {vendor_name} in bullet points.
    Focus on:
    - Market position
    - Technology focus
    - Growth direction
    
    Keep the entire response under 130 characters and start each line with a dash (-).
    """
    
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "You are a market analysis assistant. Provide concise, current market trends."},
                {"role": "user", "content": prompt}
            ],
            stream=False
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"Error fetching market trends for {vendor_name}: {str(e)}")
        return "- Market trend data not available"

def main():
    # Test vendors
    test_vendors = ["Microsoft", "AWS", "Google"]
    
    print("Starting test with 3 vendors...")
    
    # Get data for each vendor and print immediately
    for vendor in test_vendors:
        print(f"\nProcessing {vendor}...")
        
        # Get financial data
        print("Fetching financial data...")
        financials = get_vendor_financials(vendor)
        print(f"Financials: {json.dumps(financials, indent=2)}")
        
        # Get market trends
        print("Fetching market trends...")
        trends = get_market_trends(vendor)
        print(f"Market Trends:\n{trends}")
        
        # Add delay to comply with API limits
        time.sleep(2)
    
    # Generate complete test data
    print("\nGenerating complete test data...")
    vendor_data = {}
    
    for vendor in test_vendors:
        vendor_data[vendor] = {
            "KeyAccountManagers": [
                f"{vendor} Manager 1, +49 123 456 7890",
                f"{vendor} Manager 2, +49 987 654 3210"
            ],
            "KeyStakeholders": [
                f"Stakeholder 1, {vendor} Team, Allianz Technology SE",
                f"Stakeholder 2, {vendor} Branch, Allianz Technology SE"
            ],
            "Financials": get_vendor_financials(vendor),
            "MarketTrends": get_market_trends(vendor),
            "Strategy": "- Negotiate competitive pricing.\n- Enhance collaboration on strategic projects.",
            "Msg": "- Strengthen partnership for mutual growth."
        }
        time.sleep(2)
    
    # Check if file exists and remove it
    output_file = Path("test_vendors_data.json")
    if output_file.exists():
        print(f"\nRemoving existing file: {output_file}")
        output_file.unlink()  # Delete the existing file
    
    # Save the test data to a JSON file
    print(f"Generating new file: {output_file}")
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(vendor_data, f, indent=4)
    
    print(f"Test data has been saved to {output_file}")
    print(f"\nTest data has been saved to {output_file}")

if __name__ == "__main__":
    main()