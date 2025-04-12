import logging
import os
import time
import argparse # Added for command-line arguments
from datetime import datetime, date # Added date
from typing import Dict, List, Any, Optional

from dotenv import load_dotenv

# Import local modules
from scraper import scrape_ranking_data # Renamed from scrape_yahoo_finance
from analyzer import analyze_with_gemini
from reporter import create_google_doc

# Load environment variables from .env file first
load_dotenv()

# --- Constants --- 
# Define URLs and their corresponding keys for categorization - Using KABUTAN URLs
RANKING_URLS = {
    "regular_up": "https://kabutan.jp/warning/value_increase",
    "regular_down": "https://kabutan.jp/warning/value_decrease",
    "pts_up": "https://kabutan.jp/warning/pts_night_price_increase",
    "pts_down": "https://kabutan.jp/warning/pts_night_price_decrease",
}

# Threshold for triggering the stop limit warning
STOP_LIMIT_THRESHOLD = 10

# Default name for the log file (will include date)
LOG_FILE_BASENAME = "stock_report"

def setup_logging(log_date_str: str):
    """Configures logging to both console and a date-stamped file."""
    log_filename = f"{LOG_FILE_BASENAME}_{log_date_str}.log"
    log_format = '%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s'
    
    # Get root logger
    logger = logging.getLogger()
    logger.setLevel(logging.INFO) # Set base level
    
    # Remove existing handlers to avoid duplication if script is run multiple times
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
        handler.close()

    # Console Handler
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(logging.Formatter(log_format))
    logger.addHandler(console_handler)
    
    # File Handler
    try:
        file_handler = logging.FileHandler(log_filename, encoding='utf-8')
        file_handler.setFormatter(logging.Formatter(log_format))
        logger.addHandler(file_handler)
        logging.info(f"Logging initialized. Log file: {log_filename}")
    except Exception as e:
        logging.error(f"Failed to set up file logging to {log_filename}: {e}")

def main(report_date: date, api_delay: float, drive_folder_id: Optional[str] = None, creds_path: str = "credentials.json"):
    """Main function to orchestrate the scraping, analysis, and reporting."""
    start_time = time.time()
    report_date_str = report_date.strftime("%Y-%m-%d")
    
    # Setup logging for this run (using the potentially overridden date)
    setup_logging(report_date_str)
    
    logging.info(f"=== Starting Stock Market Analysis Report Generation for {report_date_str} ===")
    logging.info(f"Using API delay: {api_delay}s")
    if drive_folder_id:
        logging.info(f"Target Google Drive Folder ID: {drive_folder_id}")
    else:
        logging.info("Targeting root Google Drive folder.")
    logging.info(f"Using Google Credentials Path: {creds_path}")

    # Manually set environment variables if overridden by args, so modules can use them
    if drive_folder_id:
        os.environ['GOOGLE_DRIVE_FOLDER_ID'] = drive_folder_id
    # Ensure the correct creds path is seen by the reporter module
    os.environ['GOOGLE_CREDS_PATH'] = creds_path 

    # Initialize data structures
    report_data: Dict[str, List[Dict[str, Any]]] = {key: [] for key in RANKING_URLS}
    stop_limit_stocks: List[Dict[str, Any]] = []
    stop_limit_counter = 0

    try: # Add top-level try-except for orchestration errors
        # --- Scraping and Analysis Phase --- 
        logging.info("--- Phase 1: Scraping and Analyzing Stocks ---")
        for category, url in RANKING_URLS.items():
            logging.info(f"Processing category: {category} - URL: {url}")
            
            scraped_stocks = scrape_ranking_data(url)
            if not scraped_stocks:
                logging.warning(f"No stocks scraped from {url}. Skipping analysis.")
                continue
            
            logging.info(f"Successfully scraped {len(scraped_stocks)} stocks from {category}.")

            for stock_info in scraped_stocks:
                stock_code = stock_info.get('code', 'N/A')
                stock_name = stock_info.get('name', 'N/A')
                logging.info(f"Analyzing stock: {stock_name} ({stock_code})")
                
                analysis_result = analyze_with_gemini(stock_info)
                
                stock_info['analysis_text'] = analysis_result.get('analysis_text', 'Analysis Failed')
                stock_info['source_urls'] = analysis_result.get('source_urls', [])
                report_data[category].append(stock_info)

                if stock_info.get('is_stop_limit'):
                    stop_limit_counter += 1
                    stop_limit_stocks.append({
                        'name': stock_name,
                        'code': stock_code,
                        'change': stock_info.get('change_percent', 0.0),
                        'market': "PTS" if "pts" in category else "Regular"
                    })

                if api_delay > 0:
                    logging.debug(f"Pausing for {api_delay}s before next API call.")
                    time.sleep(api_delay)
            
            logging.info(f"Finished analysis for category: {category}.")

        logging.info("--- Completed Scraping and Analysis Phase ---")
        logging.info(f"Total stocks analyzed: {sum(len(v) for v in report_data.values())}")
        logging.info(f"Total stop-high/low stocks found: {stop_limit_counter}")

        # --- Reporting Phase --- 
        logging.info("--- Phase 2: Generating Google Docs Report ---")
        stop_limit_warning_message: Optional[str] = None

        if stop_limit_counter >= STOP_LIMIT_THRESHOLD:
            logging.warning(f"Stop limit threshold ({STOP_LIMIT_THRESHOLD}) met or exceeded! Generating warning.")
            warning_lines = [
                f"【Warning】{stop_limit_counter} Stop-High/Low Stocks Today!",
                "----------------------------------------"
            ]
            for stock in stop_limit_stocks:
                market_type = stock['market']
                change_str = f"{stock['change']:+.2f}%"
                stop_type = "(Ｓ高)" if stock['change'] > 0 else "(Ｓ安)"
                warning_lines.append(f"- {stock['name']} ({stock['code']}) - Market: {market_type}, Change: {change_str} {stop_type}")
            stop_limit_warning_message = "\n".join(warning_lines)
        
        # Pass the correct date string and optional warning
        document_url = create_google_doc(
            report_data=report_data, 
            date_str=report_date_str, # Use formatted date string
            stop_limit_warning=stop_limit_warning_message
        )

        if document_url:
            logging.info(f"Successfully created Google Docs report: {document_url}")
        else:
            logging.error("Failed to create Google Docs report. Check previous logs.")
            
    except Exception as e:
        logging.critical(f"An unexpected error occurred during the main process: {e}", exc_info=True)

    # --- Finish --- 
    end_time = time.time()
    logging.info(f"=== Stock Market Analysis Report Generation Finished ==")
    logging.info(f"Total execution time: {end_time - start_time:.2f} seconds")

# --- Entry Point --- 
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate a stock market analysis report using Kabutan data and Gemini analysis.")
    
    # Date Argument
    parser.add_argument(
        "-d", "--date", 
        type=str, 
        default=datetime.now().strftime("%Y-%m-%d"),
        help="Report date in YYYY-MM-DD format. Defaults to today."
    )
    # Drive Folder ID Argument
    parser.add_argument(
        "-f", "--folder-id", 
        type=str, 
        default=os.getenv("GOOGLE_DRIVE_FOLDER_ID"), # Default to env var if set
        help="Google Drive Folder ID where the report document should be created. Overrides GOOGLE_DRIVE_FOLDER_ID env var."
    )
    # Credentials Path Argument
    parser.add_argument(
        "-c", "--creds-path",
        type=str,
        default=os.getenv("GOOGLE_CREDS_PATH", "credentials.json"),
        help="Path to the Google service account credentials JSON file. Overrides GOOGLE_CREDS_PATH env var."
    )
    # API Delay Argument
    parser.add_argument(
        "--api-delay",
        type=float,
        default=float(os.getenv("GEMINI_API_DELAY", 1.0)), # Default to env var or 1.0
        help="Delay in seconds between Gemini API calls. Defaults to 1.0 or GEMINI_API_DELAY env var."
    )
    
    args = parser.parse_args()

    # Validate date format
    try:
        report_date_obj = datetime.strptime(args.date, "%Y-%m-%d").date()
    except ValueError:
        print(f"Error: Invalid date format ({args.date}). Please use YYYY-MM-DD.")
        logging.error(f"Invalid date format provided via command line: {args.date}")
        exit(1) # Exit if date format is wrong

    # Validate API delay
    if args.api_delay < 0:
        print("Error: API delay cannot be negative.")
        logging.error(f"Invalid API delay provided: {args.api_delay}")
        exit(1)

    # Check for essential API key and credentials file before running main logic
    gemini_key = os.getenv("GEMINI_API_KEY")
    if not gemini_key:
        logging.error("CRITICAL: GEMINI_API_KEY environment variable not set. Set it in your .env file or environment.")
        print("Error: GEMINI_API_KEY not found.")
        exit(1)
        
    if not os.path.exists(args.creds_path):
        logging.error(f"CRITICAL: Google credentials file not found at '{args.creds_path}'. Check the path provided via --creds-path or the default.")
        print(f"Error: Credentials file not found at '{args.creds_path}'.")
        exit(1)

    # Run the main logic with parsed arguments
    main(
        report_date=report_date_obj,
        api_delay=args.api_delay,
        drive_folder_id=args.folder_id,
        creds_path=args.creds_path
    ) 