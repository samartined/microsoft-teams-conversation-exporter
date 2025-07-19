# Microsoft Teams Conversation Exporter

Export your Microsoft Teams conversations to JSON and PDF with professional formatting and integrity verification.

## What This Tool Does

- **Exports complete conversations** from Microsoft Teams to JSON format
- **Converts to professional PDF** with integrity verification
- **Handles pagination automatically** - no manual intervention needed
- **Documentation with cryptographic hashes** for data integrity
- **Multi-language PDFs** - English, Spanish, French, German
- **Professional documentation** suitable for business use

## Quick Start Guide

### Prerequisites

1. **Python 3.7+** installed on your system
2. **Microsoft Teams account** with access to the conversation you want to export
3. **Required packages**: Install with `pip install -r requirements.txt`

### Step 1: Get Your Access Token

1. Go to [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
2. Sign in with your Microsoft account
3. Look for the **"Access token"** section on the right side
4. Click **"Copy"** next to the token
5. The token will be copied to your clipboard

### Step 2: Find Your Chat ID

1. In Graph Explorer, execute this query:
   ```
   GET https://graph.microsoft.com/v1.0/me/chats?$expand=members
   ```
2. Look for the chat with the person you want to export
3. Copy the **"id"** field from that chat

### Step 3: Run the Export

```bash
python teams_exporter.py
```

The script will:
1. Ask for your Chat ID
2. Ask for your access token
3. Export all messages to JSON with integrity verification
4. Let you choose PDF language
5. Generate a professional PDF with integrity certificates

## Output Files

- **JSON file**: `exported_messages/complete_conversation_YYYYMMDD_HHMMSS.json`
- **PDF file**: `teams_conversation.pdf` with integrity verification

## Features

### Professional PDF Export with Integrity Verification
- **Integrity certificate** with cryptographic hashes
- **API response verification** - hash of each original API response
- **Session metadata** - export context and timestamps
- **Master hash** - verification of all data integrity
- **Multi-language support** (EN, ES, FR, DE)
- **Clean formatting** with message timestamps
- **User identification** and message content
- **Professional documentation** for business use

### Export Process
- **No temporary files** - everything in memory
- **Automatic pagination** - handles any conversation size
- **Progress indicators** - see export status in real-time
- **Error handling** - clear messages for common issues
- **Professional documentation** - suitable for business records

## Integrity Verification Features

### Cryptographic Integrity
- **API Response Hashes**: Each page of data gets a unique SHA-256 hash
- **Session Metadata**: Context of the export process
- **Master Hash**: Combined hash of all API responses
- **Timestamp Verification**: UTC timestamps for each operation

### Data Verification
- **Verifiable Authenticity**: Any modification breaks the cryptographic chain
- **Independent Verification**: Third parties can verify data integrity
- **Complete Documentation**: Audit trail of the export process
- **Professional Standards**: Meets business documentation standards

## Troubleshooting

### "Token expired or invalid"
- Get a fresh token from Graph Explorer
- Tokens expire after some time

### "Chat ID not found"
- Make sure you have access to the conversation
- Verify the Chat ID starts with "19:"
- Check that the conversation is accessible via Graph API

### "No messages retrieved"
- Verify your Chat ID is correct
- Ensure you have permission to access the conversation
- Check your internet connection

## Language Support

The PDF can be generated in:
- **English (en)** - Default
- **Spanish (es)** - Español
- **French (fr)** - Français  
- **German (de)** - Deutsch

## File Structure

```
├── teams_exporter.py    # Main script with integrity verification
├── language_config.json         # PDF text translations
├── requirements.txt             # Python dependencies
├── exported_messages/           # Output directory
│   └── complete_conversation_*.json
└── teams_conversation.pdf       # Final PDF with integrity certificates
```

## Legal Notice

This tool exports conversations using the official Microsoft Graph API and implements integrity verification protocols. The generated PDF includes cryptographic integrity certificates for data verification. Use responsibly and in compliance with your organization's policies and applicable laws.

### Integrity Features
- **Cryptographic Verification**: SHA-256 hashes of original API responses
- **Data Integrity**: Complete audit trail from API to PDF
- **Professional Standards**: Meets business documentation standards
- **Independent Verification**: Third-party validation possible

## Support

For issues or questions:
1. Check the troubleshooting section above
2. Verify your Chat ID and access token
3. Ensure you have proper permissions for the conversation
4. For legal use, consult with qualified legal professionals

---

**Note**: This tool is designed for personal and authorized business use with integrity verification. Always respect privacy and data protection regulations. The integrity features help ensure data authenticity for business documentation. 