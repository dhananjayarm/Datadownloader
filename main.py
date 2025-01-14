import requests
import pandas as pd
import os
import json


# Function to load configuration from JSON file
def load_config(config_file):
    with open(config_file, 'r') as file:
        config = json.load(file)
    return config


# Replace this with your actual Polygon.io API key
#API_KEY = 'oqfu6jpyeFlp35iSJ1v1sW3BApO2RtM_'

# Base URL for Polygon API (historical data endpoint)
BASE_URL = 'https://api.polygon.io/v2/aggs/ticker/'


# Function to download historical data for a given symbol
def download_symbol_data(symbol, start_date, end_date,timeframe,multiplier,apikey):
    url = f"{BASE_URL}{symbol}/range/{multiplier}/{timeframe}/{start_date}/{end_date}"
    params = {
        'apiKey': apikey
    }

    try:
        response = requests.get(url, params=params)
        response.raise_for_status()  # Raise an exception for HTTP errors

        data = response.json()
        # print(f"Response for {symbol}: {json.dumps(data, indent=2)}")

        if data['status'] == 'OK' and 'results' in data:
            # Convert the data to a pandas DataFrame
            df = pd.DataFrame(data['results'])

            df['symbol'] = symbol
            # Move the 'symbol' column to the first position
            symbol_col = df.pop('symbol')
            df.insert(0, 'symbol', symbol_col)  # Insert 'symbol' at the first position

            # Move the 'timestamp' column to the second position
            timestamp_col = df.pop('t')
            df.insert(1, 'timestamp', timestamp_col)  # Insert 'timestamp' at the second position

            # Convert the timestamp column to datetime (already moved to the second position)
            df['timestamp'] = pd.to_datetime(df['timestamp'], unit='ms')

            # Rename columns if necessary
            df.rename(columns={'o': 'open', 'c': 'close', 'h': 'high', 'l': 'low', 'v': 'volume', 'vw':'vwap', 'n': 'no. of trades'}, inplace=True)

            return df
        else:
            print(f"No data found for {symbol} or error: {data.get('error', 'Unknown error')}")
            return None

    except requests.exceptions.RequestException as e:
        print(f"An error occurred while fetching data for {symbol}: {e}")
        return None


# Function to append or create a new sheet in Excel
# Function to append or create a new sheet in Excel
def append_to_excel(symbol, new_data, output_file):

    if os.path.exists(output_file):
        # File exists, so load the existing Excel file
        with pd.ExcelFile(output_file) as xls:
            if symbol in xls.sheet_names:
                # Sheet exists, so read existing data
                existing_data = pd.read_excel(xls, sheet_name=symbol)

                # Convert both existing and new data's timestamp to datetime for comparison
                existing_data['timestamp'] = pd.to_datetime(existing_data['timestamp'])
                new_data['timestamp'] = pd.to_datetime(new_data['timestamp'])

                # Get the last date in the existing data
                last_date_existing = existing_data['timestamp'].max()

                # Filter new data to include only rows with a timestamp greater than the last existing date
                new_data_filtered = new_data[new_data['timestamp'] > last_date_existing]

                if not new_data_filtered.empty:
                    # Concatenate the existing data and the new filtered data
                    combined_data = pd.concat([existing_data, new_data_filtered], ignore_index=True)

                    # Write the combined data back to the sheet
                    with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        combined_data.to_excel(writer, sheet_name=symbol, index=False)
                    print(f"Data for {symbol} appended successfully.")
                else:
                    print(f"No new data for {symbol} to append.")
            else:
                # Sheet doesn't exist, create a new one
                with pd.ExcelWriter(output_file, engine='openpyxl', mode='a') as writer:
                    new_data.to_excel(writer, sheet_name=symbol, index=False)
                print(f"New sheet created for {symbol} and data added.")
    else:
        # Excel file doesn't exist, create a new file
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            new_data.to_excel(writer, sheet_name=symbol, index=False)
        print(f"Excel file created and data for {symbol} added.")


# Main function to download data for all symbols and save to Excel
def download_and_save_to_excel(config):
    symbols = config['symbols']
    start_date = config['start_date']
    end_date = config['end_date']
    output_file = config['output_file']
    output_dir = config['output_dir']
    apikey = config['API_KEY']
    timeframe = config['timeframe']
    multiplier = config['multiplier']
    if timeframe == 'minute':
        tf ='min'
    else: tf=timeframe
    if multiplier ==1:
        outputfilestr = output_dir+ output_file + '_' + tf+'.xlsx'
    else :
        outputfilestr = output_dir+ output_file + '_'+ multiplier+ '_' + tf+'.xlsx'
    for symbol in symbols:
        print(f"Downloading data for {symbol}...")
        df = download_symbol_data(symbol, start_date, end_date,timeframe,multiplier,apikey)
        if df is not None:
            append_to_excel(symbol, df, outputfilestr)


if __name__ == "__main__":
    # Load configuration from config.json
    config_file = 'config.json'
    config = load_config(config_file)

    # Run the data download process
    download_and_save_to_excel(config)
