from io import BytesIO
import os
import json
import asyncio
import pandas as pd
import requests
from dateutil import parser
from datetime import datetime, timedelta
from aiohttp import web
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError
from slack_sdk.signature import SignatureVerifier

from config import *


class LlmParser:
    def __init__(self, api_key):
        self.api_key = api_key
        self.api_url = OPENAI_API_URL

    def parse_message(self, message_text):
        """Parse message using OPENAI API to extract items."""
        prompt = f"""Extract order items from the following message. Format the output as a JSON object with the structure: 
        {{"items": [{{"name": "item"}}]}}. 

        If no items are found in the message, return: 
        {{"items": [{{"name": ""}}]}}

        Do not include any additional text, explanations, or markdown formatting (e.g., ```json). 

        Message: {message_text}"""
        
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        
        data = {
            "model": "gpt-4o",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.1
        }

        try:
            response = requests.post(self.api_url, headers=headers, json=data)
            response.raise_for_status()
            result = response.json()
            return json.loads(result['choices'][0]['message']['content'])
        except Exception as e:
            print(f"Error parsing message with OPENAI: {e}")
            return {"items": []}

class OutputManager:  
    def generate_excel_bytes(self, data):
        """Safer DataFrame creation"""
        try:
            # Ensure we always have a list of dicts
            clean_data = data or []
            if not isinstance(clean_data, list):
                clean_data = [clean_data]
                
            df = pd.DataFrame(clean_data)
            
            # Ensure required columns exist
            for col in ['name', 'date', 'products']:
                if col not in df.columns:
                    df[col] = ''
            
            output = BytesIO()
            df.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)
            return output.getvalue()
            
        except Exception as e:
            print(f"Excel generation error: {e}")
            # Return empty Excel file as fallback
            return pd.DataFrame([{'error': str(e)}]).to_excel(index=False)

class KitchenRequestBot:
    def __init__(self):
        self.slack_client = WebClient(token=SLACK_BOT_TOKEN)
        self.parser = LlmParser(OPENAI_API_KEY)
        self.output_manager = OutputManager()

    def validate_date(self, date_str):
        """Validate date is within past 30 days"""
        try:
            input_date = parser.parse(date_str).replace(tzinfo=None)
            current_date = datetime.now().replace(tzinfo=None)
            
            # Calculate date boundaries
            max_past_date = current_date - timedelta(days=30)
            max_future_date = current_date + timedelta(days=1)  # Allow today
            
            # Check date range
            if input_date < max_past_date:
                raise ValueError(f"Date cannot be older than 30 days ({max_past_date.strftime('%m/%d/%Y')})")
                
            if input_date > max_future_date:
                raise ValueError("Future dates are not allowed")
            
            return input_date
            
        except ValueError as e:
            error_msg = f"""
            Invalid date: {date_str}
            • Must be in MM/DD/YYYY format
            • Cannot be older than 30 days
            • Cannot be in the future
            Example valid date: {(datetime.now() - timedelta(days=7)).strftime('%m/%d/%Y')}
            """
            raise ValueError(error_msg)
    
    def _open_dm_channel(self, user_id):
        """Open a DM channel with the user"""
        try:
            response = self.slack_client.conversations_open(users=user_id)
            return response['channel']['id']
        except SlackApiError as e:
            print(f"Error opening DM channel: {e}")
            return None
    
    def send_file_via_dm(self, user_id, file_bytes, filename):
        """Send file to user via DM"""
        dm_channel = self._open_dm_channel(user_id)
        if not dm_channel:
            return False
        
        try:
            self.slack_client.files_upload_v2(
                channels=dm_channel,
                file=file_bytes,
                filename=filename,
                title="Your Order Requests Report"
            )
            return True
        except SlackApiError as e:
            print(f"Error sending file: {e}")
            return False

    def fetch_messages(self, channel_id, start_date):
        """Fetch messages from Slack channel after start_date."""
        messages = []
        try:
            start_timestamp = start_date.timestamp()
            result = self.slack_client.conversations_history(
                channel=channel_id,
                oldest=str(start_timestamp)
            )
            messages.extend(result['messages'])
            
            while result.get('has_more', False):
                result = self.slack_client.conversations_history(
                    channel=channel_id,
                    oldest=str(start_timestamp),
                    cursor=result['response_metadata']['next_cursor']
                )
                messages.extend(result['messages'])
                
        except SlackApiError as e:
            print(f"Error fetching messages: {e}")
        
        return messages

    def get_user_info(self, user_id):
        """Get user information from Slack."""
        try:
            result = self.slack_client.users_info(user=user_id)
            return result['user']['real_name']
        except SlackApiError as e:
            print(f"Error fetching user info: {e}")
            return user_id

    def process_messages(self, start_date_str):
        """Process messages from the specified start date."""
        try:
            start_date = self.validate_date(start_date_str)
            messages = self.fetch_messages(SLACK_CHANNEL_ID, start_date)
            
            if not messages:
                raise ValueError("No messages found in the specified date range")
            
            parsed_data = []
            for msg in messages:
                if 'user' not in msg or 'text' not in msg:
                    continue
                                
                user_name = self.get_user_info(msg['user'])
                msg_date = datetime.fromtimestamp(float(msg['ts'])).strftime('%m/%d/%Y')
                parsed_items = self.parser.parse_message(msg['text'])
                
                for item in parsed_items['items']:
                    if item['name'] == "":
                        continue
                    
                    parsed_data.append({
                        'name': user_name,
                        'date': msg_date,
                        'products': item['name']
                    })
            print("Parsed data: ", parsed_data)
            return parsed_data
            
        except SlackApiError as e:
            error_msg = {
                "channel_error": ":lock: Bot doesn't have access to this channel",
                "not_in_channel": ":eyes: Bot needs to be added to the channel first"
            }.get(e.response['error'], "Error fetching messages from Slack")
            raise ValueError(error_msg)

        except json.JSONDecodeError:
            raise ValueError("Failed to parse order information from messages")

        except Exception as e:
            raise ValueError("Error processing messages") from e
        
def format_error(title, message, example=None):
    """Format error messages for Slack with proper formatting"""
    blocks = [
        {
            "type": "section",
            "text": {
                "type": "mrkdwn",
                "text": f":x: *{title}* \n{message}"
            }
        }
    ]
    
    if example:
        blocks.append({
            "type": "section",
            "text": {
                "type": "mrkdwn",
                "text": f":bulb: *Example:* `{example}`"
            }
        })
        
    return {"blocks": blocks}
        
async def verify_request(request):
    """Verify Slack request signature"""
    verifier = SignatureVerifier(SLACK_SIGNING_SECRET)
    body = await request.text()
    return verifier.is_valid_request(body, request.headers)

async def handle_slack_command(request):
    """Handle incoming Slack command"""
    if not await verify_request(request):
        return web.json_response({"text": "Unauthorized request"}, status=403)

    data = await request.post()
    user_id = data.get('user_id')
    command_text = data.get('text', '').strip()
    response_url = data.get('response_url')

    # Immediate acknowledgement
    asyncio.create_task(process_command_background(user_id, command_text, response_url))
    return web.json_response({
        "response_type": "ephemeral",
        "text": "Processing your request. You'll receive the report via DM shortly."
    })

async def process_command_background(user_id, date_str, response_url):
    try:
        bot = KitchenRequestBot()
        
        # Process messages
        parsed_data = bot.process_messages(date_str)
        
        # Handle no data found
        if not parsed_data:
            requests.post(response_url, json={
                "response_type": "ephemeral",
                **format_error(
                    "No Data Found",
                    f"No valid orders found since {date_str}",
                    f"/shopping-list {(datetime.now() - timedelta(days=3)).strftime('%m/%d/%Y')}"
                )
            })
            return

        # Generate Excel file
        try:
            excel_bytes = bot.output_manager.generate_excel_bytes(parsed_data)
        except Exception as e:
            requests.post(response_url, json={
                "response_type": "ephemeral",
                **format_error(
                    "Report Generation Failed",
                    "Could not create the Excel file. Please try again later."
                )
            })
            return

        # Send via DM
        try:
            success = bot.send_file_via_dm(
                user_id=user_id,
                file_bytes=excel_bytes,
                filename="kitchen_orders.xlsx"
            )
        except SlackApiError as e:
            requests.post(response_url, json={
                "response_type": "ephemeral",
                **format_error(
                    "Delivery Failed",
                    "Couldn't send you a DM. Please check if you have DMs enabled with this app."
                )
            })
            return

        # Final response
        if success:
            requests.post(response_url, json={
                "response_type": "ephemeral",
                "text": ":white_check_mark: Your report has been sent to your DMs!"
            })
        else:
            requests.post(response_url, json={
                "response_type": "ephemeral",
                "text": ":x: Failed to send report. Please try again later."
            })

    except ValueError as e:
        # Handle date validation errors
        error_lines = str(e).strip().split('\n')
        requests.post(response_url, json={
            "response_type": "ephemeral",
            **format_error(
                "Invalid Date Format",
                "\n".join(error_lines[1:]) if len(error_lines) > 1 else str(e),
                f"/shopping-list {(datetime.now() - timedelta(days=3)).strftime('%m/%d/%Y')}"
            )
        })
        
    except requests.exceptions.RequestException as e:
        # Handle OpenAI API errors
        requests.post(response_url, json={
            "response_type": "ephemeral",
            **format_error(
                "API Error",
                "Failed to process messages. Please try again in a few minutes."
            )
        })
        
    except Exception as e:
        # Generic error handler
        print(f"Unexpected error: {str(e)}")
        requests.post(response_url, json={
            "response_type": "ephemeral",
            **format_error(
                "Something Went Wrong",
                "We encountered an unexpected error. Our team has been notified."
            )
        })

async def start_server():
    app = web.Application()
    app.router.add_post('/slack/command', handle_slack_command)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, '0.0.0.0', int(os.getenv('PORT', 3000)))
    await site.start()
    print("Server running...")
    await asyncio.Event().wait()

if __name__ == "__main__":
    asyncio.run(start_server())
