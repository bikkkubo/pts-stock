import logging
import os
from typing import List, Dict, Any, Optional

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Load environment variables - Ensure load_dotenv() is called in main.py

# Define necessary scopes for Google Docs and Drive (if using folder ID)
SCOPES = [
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/drive.file' # Needed to specify/create in folder
]

# --- Helper for building batchUpdate requests --- 
def _build_requests(report_data: Dict[str, List[Dict[str, Any]]], date_str: str, stop_limit_warning: Optional[str]) -> List[Dict[str, Any]]:
    """Constructs the list of requests for docsService.batchUpdate.
       Builds content in reverse and inserts at index 1 for correct final order.
    """
    requests = []

    def add_req(text: str, bold: bool = False, heading: Optional[str] = None):
        """Appends requests to insert text and apply styles at the beginning.
        Inserts at the beginning (index 1) of the current state.
        """
        insert_len = len(text)
        # Request to insert text at the beginning (index 1)
        insert_text_req = {
            'insertText': {
                'location': {'index': 1},
                'text': text + "\n" # Add newline
            }
        }
        requests.append(insert_text_req)

        # Apply styles to the range just inserted (index 1 to 1 + length)
        if bold:
            style_req = {
                'updateTextStyle': {
                    'range': {'startIndex': 1, 'endIndex': 1 + insert_len},
                    'textStyle': {'bold': True},
                    'fields': 'bold'
                }
            }
            requests.append(style_req)
        if heading:
            para_style_req = {
                'updateParagraphStyle': {
                    'range': {'startIndex': 1, 'endIndex': 1 + insert_len},
                    'paragraphStyle': {'namedStyleType': heading},
                    'fields': 'namedStyleType'
                }
            }
            requests.append(para_style_req)

    # --- Build content in reverse order --- 

    # 3. Add Market Sections (Iterate sections/stocks in reverse)
    section_map = {
        "regular_up": "Regular Market - Top Gainers",
        "regular_down": "Regular Market - Top Losers",
        "pts_up": "PTS Market - Top Gainers",
        "pts_down": "PTS Market - Top Losers",
    }
    for key in reversed(["regular_up", "regular_down", "pts_up", "pts_down"]):
        section_title = section_map.get(key)
        stocks = report_data.get(key, [])
        if not section_title or not stocks:
            continue 

        add_req("") # Blank line after section
        
        for stock in reversed(stocks):
            add_req("") # Blank line between stocks
            
            rank = stock.get('rank', 'N/A')
            name = stock.get('name', 'N/A')
            code = stock.get('code', 'N/A')
            change = stock.get('change_percent', 0.0)
            stop_status = " (Ｓ高)" if stock.get('is_stop_limit') and change > 0 else " (Ｓ安)" if stock.get('is_stop_limit') else ""
            analysis = stock.get('analysis_text', 'N/A').strip()
            sources = stock.get('source_urls', [])

            sources_text = "  Source(s): " + (", ".join(sources) if sources else "N/A")
            add_req(sources_text)

            analysis_text = f"  Gemini Analysis: {analysis}"
            add_req(analysis_text)

            stock_line = f"{rank}. {name} ({code}) - Change: {change:+.2f}%{stop_status}"
            add_req(stock_line)

        # Add Section Header 
        add_req(section_title, bold=True, heading='HEADING_2')

    # 2. Add Report Date 
    add_req("") 
    add_req(f"Report Date: {date_str}")

    # 1. Add Optional Warning 
    if stop_limit_warning:
        add_req("") 
        add_req(stop_limit_warning, bold=True, heading='HEADING_1')

    # The list is built in reverse, so it's ready for batchUpdate
    return requests
# --- End Helper --- 

def create_google_doc(report_data: Dict[str, List[Dict[str, Any]]], date_str: str, stop_limit_warning: Optional[str] = None) -> str:
    """Creates a Google Document report from the analysis data.

    Args:
        report_data: Dict keyed by market/ranking type containing lists of analyzed stocks.
        date_str: The date string (YYYY-MM-DD).
        stop_limit_warning: Optional pre-formatted warning string.

    Returns:
        The URL of the created Google Document, or an empty string if failed.
    """
    logging.info("Attempting to create Google Docs report...")
    doc_url = ""
    creds = None
    creds_path = os.getenv("GOOGLE_CREDS_PATH", "credentials.json")
    folder_id = os.getenv("GOOGLE_DRIVE_FOLDER_ID")

    if not os.path.exists(creds_path):
        logging.error(f"Google credentials file not found: '{creds_path}'. Set GOOGLE_CREDS_PATH or place credentials.json.")
        return ""

    try:
        creds = service_account.Credentials.from_service_account_file(creds_path, scopes=SCOPES)
        docs_service = build('docs', 'v1', credentials=creds)
        drive_service = build('drive', 'v3', credentials=creds)

        title = f"{date_str} Stock Market Analysis Report"
        doc_id = None

        # Create the document (optionally in a folder)
        if folder_id:
            try:
                # Check folder exists
                drive_service.files().get(fileId=folder_id, fields='id').execute()
                logging.info(f"Target folder ID '{folder_id}' found.")
                file_metadata = {'name': title, 'mimeType': 'application/vnd.google-apps.document', 'parents': [folder_id]}
                document = drive_service.files().create(body=file_metadata, fields='id').execute()
                doc_id = document.get('id')
                logging.info(f"Created Google Doc ID: {doc_id} in folder {folder_id}")
            except HttpError as folder_error:
                if folder_error.resp.status == 404:
                    logging.error(f"Google Drive Folder ID '{folder_id}' not found.")
                else:
                    logging.error(f"Error accessing folder '{folder_id}': {folder_error}")
                return "" # Stop if folder invalid
        else:
            doc_body = {'title': title}
            document = docs_service.documents().create(body=doc_body).execute()
            doc_id = document.get('documentId')
            logging.info(f"Created Google Doc ID: {doc_id} in root Drive")

        if not doc_id:
            logging.error("Failed to get Google Doc ID after creation.")
            return ""

        # Build and execute content update requests
        update_requests = _build_requests(report_data, date_str, stop_limit_warning)

        if update_requests:
            docs_service.documents().batchUpdate(
                documentId=doc_id, body={'requests': update_requests}
            ).execute()
            logging.info(f"Document content updated for Doc ID: {doc_id}")
        else:
            logging.warning(f"No content requests generated for Doc ID: {doc_id}")

        doc_url = f"https://docs.google.com/document/d/{doc_id}/edit"

    except HttpError as error:
        status_code = error.resp.status
        reason = getattr(error, 'reason', 'Unknown')
        content_str = getattr(error, 'content', b'').decode('utf-8', 'ignore')
        logging.error(f"Google API HTTP error: {status_code} ({reason}) - {content_str}")
        if status_code == 403:
            logging.error("Permission denied. Check service account permissions for Docs API and Drive folder.")
        elif status_code == 401:
            logging.error("Authentication error. Check credentials.")
    except Exception as e:
        logging.error(f"Failed Google Docs operation: {e}", exc_info=True)

    return doc_url

# Example Usage (for testing)
if __name__ == '__main__':
    from dotenv import load_dotenv
    load_dotenv()

    mock_report_data = {
        "regular_up": [
            {'rank': 1, 'code': '1111', 'name': 'アップル', 'change_percent': 10.5, 'is_stop_limit': True, 'analysis_text': '決算良好', 'source_urls': ['url1']},
            {'rank': 2, 'code': '2222', 'name': 'バナナ', 'change_percent': 8.2, 'is_stop_limit': False, 'analysis_text': '提携発表', 'source_urls': []}
        ],
        "regular_down": [
            {'rank': 1, 'code': '3333', 'name': 'チェリー', 'change_percent': -12.1, 'is_stop_limit': True, 'analysis_text': '下方修正', 'source_urls': ['url2']}
        ],
        "pts_up": [],
        "pts_down": [
             {'rank': 1, 'code': '4444', 'name': 'デート', 'change_percent': -9.9, 'is_stop_limit': False, 'analysis_text': '市場全体安', 'source_urls': []}
        ]
    }
    mock_date = "2024-07-31"
    mock_warning = (
        "【Warning】2 Stop-High/Low Stocks Today!\n"
        "----------------------------------------\n"
        "- アップル (1111) - Market: Regular, Change: +10.50% (Ｓ高)\n"
        "- チェリー (3333) - Market: Regular, Change: -12.10% (Ｓ安)"
    )

    print("--- Testing Google Docs Creation ---")
    creds_path_to_check = os.getenv("GOOGLE_CREDS_PATH", "credentials.json")
    if os.path.exists(creds_path_to_check):
        print(f"Using credentials: {creds_path_to_check}")
        target_folder = os.getenv("GOOGLE_DRIVE_FOLDER_ID")
        print(f"Target Folder ID: {target_folder if target_folder else 'Root Drive'}")
        generated_url = create_google_doc(mock_report_data, mock_date, stop_limit_warning=mock_warning)
        if generated_url:
            print(f"\nReport URL: {generated_url}")
        else:
            print("\nFailed to generate report. Check logs.")
    else:
        print(f"Credentials file not found at '{creds_path_to_check}'. Skipping test.")
        print("Set GOOGLE_CREDS_PATH or place credentials.json.")

    print("\nReporter module loaded.") 