import logging
import os
import re
from typing import Dict, Any, List

import google.generativeai as genai
from google.api_core import exceptions as google_exceptions

# Load environment variables - Ensure load_dotenv() is called in main.py
# from dotenv import load_dotenv
# load_dotenv()

GEMINI_MODEL = "gemini-pro"

def analyze_with_gemini(stock_info: Dict[str, Any]) -> Dict[str, Any]:
    """Analyzes the reason for a stock's price movement using the Gemini API.

    Args:
        stock_info: A dictionary containing stock details:
            {
                'rank': int,
                'code': str,
                'name': str,
                'change_percent': float,
                'is_stop_limit': bool,
                'source_url': str # To determine market type if needed
            }

    Returns:
        A dictionary containing:
        {
            'analysis_text': str, # Analysis from Gemini or error message
            'source_urls': List[str] # List of URLs found
        }
    """
    stock_name = stock_info.get('name', 'N/A')
    stock_code = stock_info.get('code', 'N/A')
    change_percent = stock_info.get('change_percent', 0.0)
    is_stop_limit = stock_info.get('is_stop_limit', False)
    # Infer market type (optional, could be passed explicitly)
    market_type = "PTS" if "pts" in stock_info.get('source_url', '') else "Regular"

    logging.info(f"Analyzing {stock_name} ({stock_code}) - Market: {market_type}, Change: {change_percent}%")

    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        logging.error("GEMINI_API_KEY not found in environment variables.")
        return {"analysis_text": "Analysis failed: API key missing", "source_urls": []}

    try:
        genai.configure(api_key=api_key)
        # Define safety settings to block harmful content
        safety_settings = [
            {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
            {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
        ]
        model = genai.GenerativeModel(GEMINI_MODEL, safety_settings=safety_settings)

        # Construct the prompt in Japanese
        change_direction = "上昇" if change_percent > 0 else "下落" # Japanese for increase/decrease
        stop_limit_text = " (ストップ高/安)" if is_stop_limit else ""

        prompt = (
            f"{market_type}市場の銘柄「{stock_name}」(コード: {stock_code})について、"
            f"前営業日に株価が{abs(change_percent):.2f}% {change_direction}した{stop_limit_text}主な要因を分析してください。"
            f"分析には、直近の適時開示情報や信頼できる主要ニュースソースを参考にしてください。"
            f"可能であれば、情報源のURLを含めてください。回答は日本語でお願いします。"
        )
        logging.debug(f"Gemini Prompt for {stock_code}: {prompt}")

        response = model.generate_content(prompt)

        # Check for safety blocks before accessing text
        if not response.parts:
             if response.prompt_feedback.block_reason:
                 block_reason = response.prompt_feedback.block_reason
                 logging.warning(f"Gemini call for {stock_name} blocked. Reason: {block_reason}")
                 analysis_text = f"分析がブロックされました。理由: {block_reason}"
             else:
                 logging.warning(f"Gemini call for {stock_name} returned no content.")
                 analysis_text = "分析結果がありませんでした。"
             source_urls = []
        else:
            analysis_text = response.text
            # Use regex to find URLs more reliably
            source_urls = re.findall(r'https?://[\S]+', analysis_text)
            # Optional: Clean up URLs if needed (e.g., remove trailing punctuation)
            source_urls = [url.rstrip('.,;:!?') for url in source_urls]

            logging.info(f"Gemini analysis successful for {stock_name}.")

        return {"analysis_text": analysis_text, "source_urls": source_urls}

    except google_exceptions.GoogleAPIError as e:
        logging.error(f"Gemini API call failed for {stock_name} (GoogleAPIError): {e}")
        return {"analysis_text": f"分析失敗 (APIエラー): {e}", "source_urls": []}
    except Exception as e:
        logging.error(f"Gemini API call failed for {stock_name} (General Error): {e}")
        return {"analysis_text": f"分析失敗 (一般エラー): {e}", "source_urls": []}

# Example Usage (for testing)
if __name__ == '__main__':
    # Requires GEMINI_API_KEY in .env
    from dotenv import load_dotenv
    load_dotenv() # Load .env for local testing

    # Example stock_info (replace with actual data if needed)
    test_stock = {
        'rank': 1,
        'code': '1234',
        'name': 'テスト株式会社',
        'change_percent': 15.5,
        'is_stop_limit': True,
        'source_url': 'https://finance.yahoo.co.jp/stocks/ranking/up?market=all'
    }

    # Make sure API key is set in .env before running this
    if os.getenv("GEMINI_API_KEY"):
        analysis_result = analyze_with_gemini(test_stock)
        print("--- Gemini Analysis Result ---")
        print(f"Stock: {test_stock['name']} ({test_stock['code']})")
        print(f"Analysis: {analysis_result['analysis_text']}")
        print(f"Sources: {analysis_result['source_urls']}")
    else:
        print("GEMINI_API_KEY not found in environment variables. Skipping test analysis.")

    print("\nAnalyzer module loaded. Run main.py to execute full workflow.") 