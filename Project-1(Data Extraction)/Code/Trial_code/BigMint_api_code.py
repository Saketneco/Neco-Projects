import requests
from datetime import datetime

def get_price_data(api_key, from_date, to_date):
    # Validate date range
    from_date_obj = datetime.strptime(from_date, '%Y-%m-%d')
    to_date_obj = datetime.strptime(to_date, '%Y-%m-%d')
    if (to_date_obj - from_date_obj).days > 7:
        raise ValueError("The difference between fromdate and todate cannot exceed 7 days.")

    # Construct the URL
    url = 'https://www.steelmint.com/api/public/v2019/pricesSMAPI'
    params = {
        'api_key': api_key,
        'fromdate': from_date,
        'todate': to_date
    }

    try:
        # Make the GET request
        response = requests.get(url, params=params)
        
        # Raise an error for bad responses
        response.raise_for_status()
        
        # Return the JSON response
        return response.json()
    
    except requests.exceptions.HTTPError as err:
        print(f"HTTP error occurred: {err}")
        print(f"Error Response: {response.text}")  # Log the error response
        return None
    except requests.exceptions.RequestException as req_err:
        print(f"Request error occurred: {req_err}")
        return None
    except Exception as err:
        print(f"An unexpected error occurred: {err}")
        return None

# Example usage
api_key = 'qkoeI6XB7YZhxupVE2Oe93'
from_date = '2024-05-07'
to_date = '2024-05-07'

price_data = get_price_data(api_key, from_date, to_date)
if price_data:
    print(price_data)
