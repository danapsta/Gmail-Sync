# Calendar Sync Tool

A Python application that synchronizes events from Google Calendar to Office 365 Calendar. This tool provides a simple GUI interface for authentication and synchronization management.

## Features

- Secure authentication for both Google Calendar and Office 365
- Graphical user interface for easy interaction
- Automatic duplicate event detection
- Secure credential storage
- Detailed logging system
- DUO authentication support for Office 365

## Prerequisites

- Python 3.7 or higher
- Google Calendar API credentials
- Office 365 account with calendar access
- Chrome browser installed (for Office 365 authentication)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/calendar-sync.git
cd calendar-sync
```

2. Create a virtual environment and activate it:
```bash
python -m venv venv
source venv/bin/activate  # On Windows, use: venv\Scripts\activate
```

3. Install required packages:
```bash
pip install -r requirements.txt
```

4. Set up Google Calendar API:
   - Go to the [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project or select an existing one
   - Enable the Google Calendar API
   - Create credentials (OAuth 2.0 Client ID)
   - Download the client configuration file
   - Rename it to `client_secrets.json` and place it in the `credentials` directory

5. Create a configuration file:
   - Copy `config.json.template` to `config.json`
   - Update the email addresses in `config.json`

## Configuration

1. Create a `credentials` directory in the project root if it doesn't exist:
```bash
mkdir credentials
```

2. Place your `client_secrets.json` file in the `credentials` directory

3. Update `config.json` with your email addresses:
```json
{
    "gmail_email": "your.email@gmail.com",
    "o365_email": "your.email@office365.com"
}
```

## Usage

1. Start the application:
```bash
python calendar_sync.py
```

2. Enter your credentials in the GUI:
   - Gmail account (must match the account used for Google Calendar API setup)
   - Office 365 email
   - Office 365 password

3. Click "Save Credentials" to store your email addresses securely

4. Click "Start Sync" to begin the synchronization process

5. Complete the DUO authentication when prompted

## Logging

The application creates a log file `calendar_sync.log` in the project directory. This file contains detailed information about the synchronization process and any errors that occur.

## Security Notes

- Credentials are stored securely using the system keyring
- OAuth 2.0 is used for Google Calendar authentication
- Office 365 cookies are stored temporarily for session management
- No passwords are stored in plain text
- All sensitive files are excluded from Git tracking

## Troubleshooting

Common issues and solutions:

1. **Google Authentication Error**
   - Verify that `client_secrets.json` is present in the credentials directory
   - Ensure the Google Calendar API is enabled in your Google Cloud Console
   - Check that your Gmail account matches the one used to create the API credentials

2. **Office 365 Authentication Error**
   - Ensure you have a stable internet connection
   - Check that your Office 365 credentials are correct
   - Verify that Chrome browser is installed and updated
   - Make sure you respond to the DUO prompt in a timely manner

3. **Sync Issues**
   - Check the `calendar_sync.log` file for detailed error messages
   - Verify that both accounts have the necessary calendar permissions
   - Ensure your system time is correctly synchronized

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Disclaimer

This tool is not officially associated with Google or Microsoft. Use at your own risk and ensure compliance with your organization's security policies.