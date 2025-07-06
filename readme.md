# Gmail Email Sender v1.0

A powerful Python automation tool for sending bulk emails through Gmail accounts using Selenium WebDriver. This tool supports multiple Gmail accounts, parallel processing, HTML email content, and comprehensive error handling.

## ⚠️ Important Notice

This tool is for educational and legitimate business purposes only. Ensure you comply with:
- Gmail's Terms of Service
- Anti-spam laws (CAN-SPAM Act, GDPR, etc.)
- Email marketing regulations
- Recipients' consent requirements

## Features

- **Multi-Account Support**: Manage multiple Gmail accounts simultaneously
- **Bulk Email Sending**: Send emails to large recipient lists efficiently
- **Parallel Processing**: Use multiple browser instances for faster execution
- **HTML Email Support**: Send rich HTML emails with custom templates
- **Error Handling**: Comprehensive error logging with screenshot capture
- **Account Verification**: Automatic handling of Gmail security challenges
- **Profile Management**: Persistent browser profiles for each account
- **Daily Limits Tracking**: Monitor and track email sending limits

## Installation

### Prerequisites

- Python 3.8 or higher
- Google Chrome browser (latest version)
- Valid Gmail accounts
- Excel files with account credentials and recipient lists

### Setup

1. **Clone or download the repository**
```bash
git clone <repository-url>
cd gmail-email-sender
```

2. **Install required packages**
```bash
pip install -r requirements.txt
```

3. **Create necessary directories**
```bash
mkdir GmailProfiles
mkdir ErrorScreenshots
```

4. **Configure settings**
Create a `settings.cfg` file with your configuration (see Configuration section)

## Configuration

Create a `settings.cfg` file in the project root:

```ini
[GMAIL_ACCOUNTS]
emails_excel_file = gmail_accounts.xlsx
email_col = Email
pass_col = Password
recovery_email_col = Recovery_Email
start_index = 0
end_index = -1

[RECIPIENT]
recipient_emails_excel = recipients.xlsx
email_col = Email
start_index = 0
end_index = -1

[EMAIL_INFO]
email_subject = Your Email Subject Here
email_html_file = email_template.html

[BROWSER]
parallel_browsers = 3
```

## File Structure

```
gmail-email-sender/
├── gmail_email_sender.py          # Main script
├── settings.cfg                   # Configuration file
├── requirements.txt               # Python dependencies
├── gmail_accounts.xlsx            # Gmail accounts data
├── recipients.xlsx                # Recipient email addresses
├── email_template.html            # HTML email template
├── GmailProfiles/                 # Browser profiles directory
├── ErrorScreenshots/              # Error screenshots directory
├── gmail_email_sender.log         # Application logs
└── daily_limit.csv               # Daily limits tracking
```

## Usage

### 1. Prepare Your Data Files

**Gmail Accounts Excel File (`gmail_accounts.xlsx`)**:
| Email | Password | Recovery_Email |
|-------|----------|----------------|
| account1@gmail.com | password123 | recovery1@gmail.com |
| account2@gmail.com | password456 | recovery2@gmail.com |

**Recipients Excel File (`recipients.xlsx`)**:
| Email |
|-------|
| recipient1@example.com |
| recipient2@example.com |

### 2. Create Email Template

Create an `email_template.html` file with your email content:

```html
<!DOCTYPE html>
<html>
<head>
    <title>Your Email Title</title>
</head>
<body>
    <h1>Hello!</h1>
    <p>This is your email content...</p>
    <p>Best regards,<br>Your Name</p>
</body>
</html>
```

### 3. Run the Script

```bash
python gmail_email_sender.py
```

The script will present a menu:
- **Add Gmail**: Test and verify Gmail account login
- **Send Emails**: Start bulk email sending process

## Script Components

### Classes

#### `BrowserHandler`
Base class for browser automation with common operations:
- Element finding and interaction
- Text input and clicking
- Browser lifecycle management

#### `GmailAccount`
Specialized Gmail account handler:
- Gmail login and authentication
- Account verification handling
- Email composition and sending
- Error handling and recovery

### Key Functions

#### `send_email(recipients, row)`
Sends emails to all recipients using a specific Gmail account.

#### `add_gmail(row, close=True)`
Creates and logs into a Gmail account instance.

#### `parallel_browsing(func, op)`
Manages parallel execution across multiple Gmail accounts.

#### `logging_error_screenshot(message, driver)`
Logs errors and captures screenshots for debugging.

## Error Handling

The script includes comprehensive error handling:

- **Screenshot Capture**: Automatically saves screenshots when errors occur
- **Retry Logic**: Automatic retry for transient failures
- **Logging**: Detailed logs in `gmail_email_sender.log`
- **Account Verification**: Handles Gmail security challenges
- **Exception Handling**: Graceful handling of Selenium exceptions

## Troubleshooting

### Common Issues

1. **Chrome Driver Issues**
   - Ensure Chrome browser is up to date
   - SeleniumBase automatically manages ChromeDriver

2. **Gmail Login Failures**
   - Check account credentials
   - Verify 2FA settings
   - Ensure "Less secure app access" is configured if needed

3. **Email Sending Failures**
   - Check Gmail daily sending limits
   - Verify HTML template format
   - Review error logs and screenshots

4. **Performance Issues**
   - Reduce `parallel_browsers` count in settings
   - Ensure sufficient system resources

### Log Files

- `gmail_email_sender.log`: Application logs
- `ErrorScreenshots/`: Screenshots of errors
- `daily_limit.csv`: Daily sending limits tracking

## Security Considerations

- **Credential Storage**: Store credentials securely
- **Profile Management**: Browser profiles contain sensitive data
- **Network Security**: Use secure networks for account access
- **Rate Limiting**: Respect Gmail's sending limits

## Legal and Ethical Use

- Obtain proper consent from recipients
- Include unsubscribe options in emails
- Comply with local anti-spam laws
- Respect Gmail's Terms of Service
- Use for legitimate business purposes only

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

This project is provided as-is for educational purposes. Users are responsible for ensuring compliance with all applicable laws and regulations.

## Support

For issues and questions:
1. Check the troubleshooting section
2. Review error logs and screenshots
3. Ensure proper configuration
4. Verify account credentials and permissions

## Changelog

### v1.0 (2025-07-06)
- Initial release
- Multi-account Gmail automation
- Parallel processing support
- HTML email templates
- Comprehensive error handling
- Account verification system

## Disclaimer

This tool is designed for legitimate email marketing and communication purposes. The authors are not responsible for any misuse of this software. Users must ensure compliance with:

- Gmail Terms of Service
- Local and international anti-spam laws
- GDPR and other privacy regulations
- Recipient consent requirements

Use responsibly and ethically.
