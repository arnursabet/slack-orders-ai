[![Review Assignment Due Date](https://classroom.github.com/assets/deadline-readme-button-22041afd0340ce965d47ae6ef1cefeee28c7c493a6346c4f15d667ab976d596c.svg)](https://classroom.github.com/a/G2xRUaVe)

# Kitchen Request Slack Bot

I currently live in a student house with 20 people, and we use a Slack workspace for communication. One of our channels is dedicated to ordering kitchen supplies, where residents post their requests. Currently, a person is responsible for going through the messages and placing the orders. To streamline this, I built a Slack bot that automates the process. The bot scans the channel for kitchen supply requests, compiles the items into an Excel spreadsheet, and sends it via Direct Message (DM). Users just need to provide a starting date, and the bot takes care of generating the report.

---

## Features

- **Slash Command Integration**: Trigger report generation with `/shopping-list [date]`
- **Date Validation**: Ensures dates are within the past 30 days
- **AI-Powered Parsing**: Uses OpenAI to extract order details from messages
- **Excel Reports**: Generates and delivers Excel files via DM
- **Error Handling**: Provides clear, user-friendly error messages

---

## Setup

### 1. Prerequisites

- Python 3.8+
- Slack App with:
  - Slash command `/shopping-list`
  - Bot token with permissions:
    - `channels:history`
    - `files:write`
    - `users:read`
    - `commands`
- OpenAI API key

### 2. Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/arnursabet/slack-ai-bot.git
   cd slack-ai-bot
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Create `.env` file:
   ```bash
   SLACK_BOT_TOKEN=xoxb-your-slack-bot-token
   SLACK_SIGNING_SECRET=your-signing-secret
   OPENAI_API_KEY=your-openai-api-key
   SLACK_CHANNEL_ID=your-channel-id
   ```

### 3. Running Locally

1. Start the bot:
   ```bash
   python main.py
   ```

2. Use ngrok for local testing:
   ```bash
   ngrok http 3000
   ```
   Update Slack app's Request URL to your ngrok URL.

---

## Usage

### Basic Command
```
/shopping-list 06/25/2024
```

### Expected Flow
1. User enters command
2. Bot acknowledges with "Processing your request..."
3. Bot sends Excel file via DM
4. User receives file and success message

---

## Error Handling

The bot provides clear error messages for common issues:

### Invalid Date Format
```
:x: Invalid Date Format
• Must be in MM/DD/YYYY format
• Cannot be older than 30 days (06/01/2024)
• Cannot be in the future
:bulb: Example: /shopping-list 06/25/2024
```

### No Data Found
```
:mag: No Data Found
No valid orders found since 06/25/2024
:bulb: Try a different date: /shopping-list 06/20/2024
```

### Channel Access Issues
```
:lock: Access Denied
Bot doesn't have access to the channel
:bulb: Please add @KitchenBot to #orders
```

---

## Demo

![Functionality Demo](demo.gif)  

---

## Configuration

| Environment Variable | Description |
|----------------------|-------------|
| `SLACK_BOT_TOKEN` | Your Slack bot's OAuth token |
| `SLACK_SIGNING_SECRET` | Slack app's signing secret |
| `OPENAI_API_KEY` | OpenAI API key |
| `SLACK_CHANNEL_ID` | Channel ID to monitor |

---

## Deployment

### Heroku
1. Create new app
2. Set environment variables
3. Deploy:
   ```bash
   heroku git:remote -a your-app-name
   git push heroku main
   ```

### Docker
```bash
docker build -t kitchen-bot .
docker run -p 3000:3000 kitchen-bot
```

---

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

---

## Reflection

I encountered a few challenges while building this project:

1. At first, I was gonna use DeepSeek API to extract the items from the messages, but I realized that it was not working as expected. So, I switched to OpenAI API.
2. I had to handle different date formats and ensure the date is within the past 30 days.
3. I had to generate an Excel spreadsheet with the extracted data. I used the `openpyxl` library to create the Excel file. At first, I was gonna use Google Sheets, but I realized that it is more convenient to send the file via DM.
4. Even though the `OutputManager` class has only one method I decided to keep it this way in case I need to add more output formats in the future.

---
