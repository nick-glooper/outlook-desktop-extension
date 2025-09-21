# Outlook Desktop Extension for Claude

A Claude Desktop Extension that provides full read/write access to Microsoft Outlook, including emails, calendar, and contacts.

## Features

- **Send Emails**: Compose and send emails through Outlook
- **Read Emails**: Search and read emails from any folder
- **Calendar Management**: Create events and view calendar
- **Contact Management**: Search and create contacts
- **Full Write Access**: Unlike web connectors, this extension provides complete Outlook access

## Prerequisites

1. **Microsoft Azure App Registration**
   - Go to [Azure Portal](https://portal.azure.com/)
   - Navigate to "Azure Active Directory" > "App registrations"
   - Click "New registration"
   - Name: "Claude Outlook Extension"
   - Account types: "Accounts in this organizational directory only" or "Personal Microsoft accounts"
   - Redirect URI: Leave blank for now
   - Click "Register"

2. **Configure API Permissions**
   - In your app registration, go to "API permissions"
   - Click "Add a permission"
   - Select "Microsoft Graph"
   - Choose "Delegated permissions"
   - Add these permissions:
     - `Mail.ReadWrite`
     - `Mail.Send`
     - `Calendars.ReadWrite`
     - `Contacts.ReadWrite`
     - `User.Read`
   - Click "Grant admin consent" (if you're an admin)

3. **Get Application Details**
   - Copy the "Application (client) ID"
   - Copy the "Directory (tenant) ID"

## Installation

1. **Download the Extension**
   - Download the `.mcpb` file from releases

2. **Install in Claude Desktop**
   - Open Claude Desktop
   - Go to Settings → Extensions
   - Click "Install Extension..."
   - Select the `.mcpb` file
   - Enter your Microsoft App Client ID and Tenant ID when prompted

3. **First-time Authentication**
   - When you first use any Outlook tool, you'll see authentication instructions
   - Follow the device code flow to authenticate with Microsoft

## Available Tools

### Email Tools
- `send_email`: Send emails with attachments
- `read_emails`: Read and search emails from any folder

### Calendar Tools
- `create_calendar_event`: Create new calendar events
- `get_calendar_events`: Retrieve calendar events for a date range

### Contact Tools
- `search_contacts`: Find contacts by name or email
- `create_contact`: Add new contacts to Outlook

## Example Usage

```
Send an email to john@example.com with the subject "Meeting Tomorrow" and ask them about the project status.

Create a calendar event for next Tuesday at 2 PM for a team meeting in the conference room.

Search for contacts with "Smith" in their name.
```

## Security & Privacy

- All authentication tokens are stored securely in your OS keychain
- No data is sent to external servers except Microsoft's APIs
- The extension runs locally on your machine
- All communications use Microsoft's secure OAuth 2.0 flow

## Troubleshooting

### Authentication Issues
- Ensure your Azure app has the correct permissions
- Check that your tenant ID and client ID are correct
- Try re-authenticating by restarting Claude Desktop

### Permission Errors
- Verify API permissions are granted in Azure Portal
- Ensure admin consent is provided if required
- Check that your Microsoft account has access to Outlook

## Development

To build from source:

```bash
npm install
mcpb pack
```

## License

MIT License - see LICENSE file for details
