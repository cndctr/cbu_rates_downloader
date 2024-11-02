import requests
import pandas as pd
from datetime import datetime, timedelta
import configparser

# Read configuration from the .ini file
config = configparser.ConfigParser()
config.read('config.ini')

# Get the days_to_fetch and currencies from the config
days_to_fetch = config.getint('settings', 'days_to_fetch')
currencies = config.get('settings', 'currencies').split(', ')

# Get the end_date from the config, or use today's date if not provided
end_date_str = config.get('settings', 'end_date', fallback=None)
if end_date_str:
    end_date_str = end_date_str.split(';')[0].strip()  # Remove comments and whitespace
    end_date = datetime.strptime(end_date_str, "%Y-%m-%d")
else:
    end_date = datetime.now()

start_date = end_date - timedelta(days=days_to_fetch)

# Create a list of dates to iterate through
date_range = [start_date + timedelta(days=x) for x in range((end_date - start_date).days + 1)]

# Prepare an empty list for results
final_data = []

# Loop through each currency first
for currency in currencies:
    print(f"Checking currency: {currency}")  # Verbose output for currency iteration
    
    for single_date in date_range:
        url_date_str = single_date.strftime("%Y-%m-%d")  # Format date as YYYY-MM-DD for URL
        display_date_str = single_date.strftime("%d.%m.%Y")  # Format date as DD.MM.YYYY for final data
        
        print(f"  Fetching data for date: {display_date_str}")  # Verbose output for date iteration
        
        url = f"https://cbu.uz/ru/arkhiv-kursov-valyut/json/{currency}/{url_date_str}/"
        
        # Make the GET request
        response = requests.get(url)

        # Check if the request was successful
        if response.status_code == 200:
            # Parse JSON response
            data = response.json()

            # Always assign the date and add to final data
            for row in data:
                final_data.append({
                    'Date': display_date_str,  # Use DD.MM.YYYY format
                    'Rate': row['Rate'],
                    'Currency': currency,
                    'BaseNominal': 1,
                    'BaseCurrency': 'UZS'
                })

            # If there's no data available, still append with Rate as None
            if not data:
                final_data.append({
                    'Date': display_date_str,
                    'Rate': None,
                    'Currency': currency,
                    'BaseNominal': 1,
                    'BaseCurrency': 'UZS'
                })
        else:
            print(f"  Failed to retrieve data for {currency} on {url_date_str}: {response.status_code}")

# Convert the final data into a DataFrame
final_df = pd.DataFrame(final_data)

# Reorder columns to the desired order
final_df = final_df[['Date', 'Rate', 'Currency', 'BaseNominal', 'BaseCurrency']]

# Create a filename with start_date and end_date
filename = f"currency_data_{start_date.strftime('%Y-%m-%d')}_to_{end_date.strftime('%Y-%m-%d')}.xlsx"

# Save to XLSX
final_df.to_excel(filename, index=False)
print(f"Data saved to {filename}")
