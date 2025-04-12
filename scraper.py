import logging
import time
import re
from typing import List, Dict, Any, Optional

import requests
from requests.exceptions import Timeout, ConnectionError, HTTPError
from bs4 import BeautifulSoup, Tag

# Standard User-Agent
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

def _parse_kabutan_table(soup: BeautifulSoup, url: str) -> List[Dict[str, Any]]:
    """Parses the ranking table (usually class 'stock_kabuka0') on Kabutan pages."""
    scraped_data = []
    table = soup.find('table', class_='stock_kabuka0')
    if not table:
        logging.warning(f"Could not find table.stock_kabuka0 on {url}")
        # Fallback: Check for other potential table classes if needed
        # table = soup.find('table', class_='some_other_class')
        # if not table: ...
        return []

    rows = table.find_all('tr')
    if not rows or len(rows) <= 1:
        logging.warning(f"Ranking table found on {url}, but no data rows detected.")
        return []

    rank = 1
    for row in rows[1:]: # Skip header row
        if rank > 10:
            break

        cols = row.find_all('td')
        # Expecting at least Code, Name, Price, Change
        # PTS pages might have fewer/different columns than regular market pages
        if len(cols) < 4:
            logging.debug(f"Skipping row on {url}: Insufficient columns ({len(cols)} < 4)")
            continue

        try:
            # --- Extract Data --- 
            # Code (Column 0)
            code_tag = cols[0].find('a')
            stock_code = code_tag.text.strip() if code_tag else cols[0].text.strip()
            if not stock_code.isdigit() or len(stock_code) != 4:
                logging.warning(f"Invalid stock code format found: '{stock_code}' on {url}. Skipping row.")
                continue

            # Name (Column 1)
            name_tag = cols[1].find('a')
            stock_name = name_tag.text.strip() if name_tag else cols[1].text.strip()
            if not stock_name:
                 logging.warning(f"Empty stock name found for code {stock_code} on {url}. Skipping row.")
                 continue

            # Price (Column 3 - Current Price or Closing Price)
            current_price: Optional[float] = None
            try:
                price_text = cols[3].text.strip().replace(',', '')
                # Handle potential non-numeric values like '--'
                if price_text and price_text != '--':
                    current_price = float(price_text)
            except (ValueError, IndexError):
                logging.debug(f"Could not parse current price for {stock_code} from '{cols[3].text if len(cols)>3 else 'N/A'}'.")

            # Change % (Find column, typically near the end with '％')
            change_col_index = -1
            for i in range(len(cols) - 1, 2, -1): # Search backwards from end
                # Prioritize columns explicitly marked with class containing 'change' or 'rate'
                col_class = cols[i].get('class', [])
                if any('change' in c or 'rate' in c for c in col_class):
                    change_col_index = i
                    break
                # Fallback: Check for '%' symbol
                if '％' in cols[i].text:
                     change_col_index = i
                     break
            
            if change_col_index == -1:
                 logging.warning(f"Change % column not identified for {stock_code} on {url}. Skipping stock.")
                 continue
                 
            price_change: Optional[float] = None
            try:
                change_text_raw = cols[change_col_index].text.strip().replace('％', '').replace('+', '')
                change_match = re.search(r'([-−－]?\d*\.?\d+)', change_text_raw)
                if change_match:
                    price_change = float(change_match.group(1))
                else:
                    logging.warning(f"Could not parse numeric change % from '{change_text_raw}' for {stock_code}")
                    # Continue without price_change, maybe it can be inferred? For now, skip.
                    continue 
            except (ValueError, IndexError):
                 logging.warning(f"Error parsing change % for {stock_code} from '{cols[change_col_index].text if change_col_index < len(cols) else 'N/A'}'.")
                 continue
            
            # Stop High/Low Status (Check all columns for S高/S安 text)
            is_stop_limit = False
            for i, col in enumerate(cols):
                col_text = col.get_text(strip=True)
                # Check within the likely change column or status columns
                if i >= 3 and ("S高" in col_text or "S安" in col_text):
                    is_stop_limit = True
                    logging.debug(f"Stop limit text found for {stock_code} in column {i}.")
                    break

            # --- Construct Data Dict --- 
            stock_info = {
                "rank": rank,
                "code": stock_code,
                "name": stock_name,
                "change_percent": price_change, # Should not be None if we reached here
                "is_stop_limit": is_stop_limit,
                "current_price": current_price, # Can be None
                "source_url": url
            }
            scraped_data.append(stock_info)
            rank += 1

        except Exception as e:
            # Catch unexpected errors during parsing of a specific row
            logging.warning(f"Unexpected error parsing row on {url}: {e} - Row: {row.text[:150]}...", exc_info=False)
            continue # Skip to the next row

    if not scraped_data:
        logging.warning(f"Parser processed {len(rows)-1} rows but extracted 0 valid stock data points from {url}.")
    return scraped_data

def scrape_ranking_data(url: str) -> List[Dict[str, Any]]:
    """Fetches and parses stock ranking data from Kabutan URLs.

    Args:
        url: The URL of the Kabutan ranking page.

    Returns:
        List of dicts, each containing stock data (up to 10), or empty list on failure.
    """
    logging.debug(f"Waiting 1 second before scraping {url}")
    time.sleep(1)
    logging.info(f"Attempting to scrape Kabutan URL: {url}")
    
    try:
        response = requests.get(url, headers=HEADERS, timeout=20)
        response.raise_for_status() # Check for HTTP errors (4xx, 5xx)
        # Ensure correct encoding is used for parsing
        response.encoding = response.apparent_encoding 
        soup = BeautifulSoup(response.content, 'lxml')

        scraped_data = _parse_kabutan_table(soup, url)

        if scraped_data:
            logging.info(f"Successfully extracted {len(scraped_data)} stocks from {url}. (Processed up to Top 10)." )
        else:
            # Parser function logs warnings internally if table/rows not found or empty
            logging.warning(f"No stock data returned by parser for {url}.")
        
        # Return only the top 10 results found
        return scraped_data[:10]

    # --- Network & Request Error Handling --- 
    except Timeout:
        logging.error(f"Request timed out fetching {url}")
    except ConnectionError:
         logging.error(f"Network connection error fetching {url}")
    except HTTPError as http_err:
        logging.error(f"HTTP error fetching {url}: {http_err} (Status: {http_err.response.status_code})" )
    except requests.exceptions.RequestException as req_err:
        logging.error(f"General request error fetching {url}: {req_err}")
    # --- Parsing & Other Error Handling --- 
    except Exception as e:
        logging.error(f"Unexpected error during scraping process for {url}: {e}", exc_info=True)

    return [] # Return empty list if any exception occurred

# --- Example Usage (for testing) --- 
if __name__ == '__main__':
    # Setup logging for the test run
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s [%(filename)s:%(lineno)d] %(message)s')
    
    test_urls = {
        "PTS Up": "https://kabutan.jp/warning/pts_night_price_increase",
        "Regular Up": "https://kabutan.jp/warning/value_increase",
        "PTS Down": "https://kabutan.jp/warning/pts_night_price_decrease",
        "Regular Down": "https://kabutan.jp/warning/value_decrease",
        "Invalid Page": "https://kabutan.jp/news/ --invalid--" # Test invalid URL
    }
    
    all_results = {}
    for name, test_url in test_urls.items():
        print(f"\n--- Testing: {name} ({test_url}) ---")
        data = scrape_ranking_data(test_url)
        all_results[name] = data
        if data:
            print(f"Result: Scraped {len(data)} stocks.")
            # Print details of first stock if available
            first_stock = data[0]
            print(f"  First Stock Example: Rank={first_stock.get('rank')}, "
                  f"Code={first_stock.get('code')}, Name={first_stock.get('name')}, "
                  f"Change={first_stock.get('change_percent')}%, Stop={first_stock.get('is_stop_limit')}, " 
                  f"Price={first_stock.get('current_price')}")
        else:
            print("Result: Scraping failed or returned no data.")
        print(f"--- Finished: {name} --- ")
        
    # Optional: You can add further checks here, e.g., assert len(all_results["PTS Up"]) > 0
    print("\nScraper test run complete.") 