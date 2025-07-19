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
- **PDF file**: `exported_messages/teams_conversation_YYYYMMDD_HHMMSS.pdf` with integrity verification

## Features

### Professional PDF Export with Dual Hash Integrity Verification
- **Dual hash system** - Content hash (deterministic) + Forensic hash (complete)
- **Content verification** - Identical hashes for same data across exports
- **Forensic integrity** - Complete chain of custody with all metadata
- **Session metadata** - Export context and UTC timestamps
- **Master hashes** - Combined verification of all content and forensic data
- **Multi-language support** (EN, ES, FR, DE) with hash explanations
- **Clean formatting** with message timestamps
- **User identification** and message content
- **Professional documentation** suitable for legal and business use

### Export Process
- **No temporary files** - everything in memory
- **Automatic pagination** - handles any conversation size
- **Progress indicators** - see export status in real-time
- **Error handling** - clear messages for common issues
- **Professional documentation** - suitable for business records

## Integrity Verification Features

### Understanding the Dual Hash System

#### What is a Hash?
A **hash** is like a digital fingerprint - it's a unique code that represents the exact content of data. Think of it as a "digital signature" that changes if even a single character is modified.

**Example**: 
- Original text: "Hello World" → Hash: `a591a6d40bf420404a011733cfb7b190d62c65bf0bcda32b57b277d9ad9f146e`
- Modified text: "Hello World!" → Hash: `7f83b1657ff1fc53b92dc18148a1d65dfc2d4b1fa3d677284addd200126d9069`

Notice how adding just one exclamation mark completely changed the hash!

#### Why Do We Need Two Types of Hashes?

When exporting conversations from Microsoft Teams, we face a challenge: **the same conversation can produce different data each time we export it**, even though the actual messages haven't changed. This happens because:

1. **Server timestamps** change with each request
2. **Pagination URLs** contain dynamic tokens
3. **API metadata** varies between requests
4. **Session information** is different each time

This creates a problem: **How do we verify that the conversation content is authentic if the data keeps changing?**

#### The Solution: Dual Hash System

We solve this problem by creating **two different types of hashes** for each page of messages:

### Content Hash (The "What" Hash)
**Think of this as the "message content fingerprint"**

- **What it includes**: Only the actual message content (text, sender, timestamps, edits)
- **What it excludes**: Server metadata, pagination URLs, session tokens
- **Behavior**: **Stays the same** for identical message content, regardless of when you export
- **Purpose**: Verifies that the actual conversation hasn't been tampered with

**Real-world analogy**: Like taking a photo of a document - the photo (content hash) shows what the document says, regardless of when you took the photo.

### Forensic Hash (The "Everything" Hash)
**Think of this as the "complete digital evidence fingerprint"**

- **What it includes**: Everything - messages, server metadata, timestamps, URLs, headers
- **What it excludes**: Nothing - captures the complete digital evidence
- **Behavior**: **May change** between exports due to server metadata differences
- **Purpose**: Provides complete forensic integrity and chain of custody

**Real-world analogy**: Like collecting all evidence from a crime scene - you want to preserve everything exactly as it was found, including the time, location, and conditions.

#### How This Solves the Problem

**Scenario**: You export the same conversation twice, 1 hour apart

**Content Hash**: 
- Export 1: `abc123...`
- Export 2: `abc123...` ✅ **Same!** (Messages are identical)

**Forensic Hash**:
- Export 1: `def456...`
- Export 2: `ghi789...` ✅ **Different!** (Server metadata changed)

**Result**: You can prove that:
1. **The conversation content is authentic** (Content Hash matches)
2. **You have complete forensic evidence** (Forensic Hash captures everything)

#### When Do Hashes Change?

**Content Hash Changes When**:
- ✅ Someone edits a message
- ✅ New messages are added
- ✅ Message content is modified
- ❌ NOT when server metadata changes

**Forensic Hash Changes When**:
- ✅ Any of the above (Content Hash changes)
- ✅ Server timestamps update
- ✅ Pagination URLs change
- ✅ API metadata varies
- ✅ Session tokens refresh

#### Why This Matters for Legal and Business Use

**For Legal Cases**:
- **Content Hash**: Proves the conversation content is authentic and unchanged
- **Forensic Hash**: Provides complete chain of custody for court evidence

**For Business Records**:
- **Content Hash**: Ensures data integrity for compliance and audits
- **Forensic Hash**: Maintains complete audit trail

**For Data Verification**:
- **Content Hash**: Allows comparison of conversation content across time
- **Forensic Hash**: Preserves complete digital evidence

#### Quick Reference Guide

| Situation | Content Hash | Forensic Hash | What It Means |
|-----------|--------------|---------------|---------------|
| Same conversation, different export times | ✅ Same | ❌ Different | **Normal** - Content unchanged, metadata varied |
| Conversation with new messages | ❌ Different | ❌ Different | **Normal** - New content added |
| Conversation with edited messages | ❌ Different | ❌ Different | **Normal** - Content was modified |
| Identical exports | ✅ Same | ✅ Same | **Rare** - Perfect timing coincidence |

**✅ = Expected behavior**  
**❌ = Different values (also expected in most cases)**

### Cryptographic Integrity
- **Dual Hash System**: Both content and forensic hashes for each page
- **Session Metadata**: Context of the export process with UTC timestamps
- **Master Content Hash**: Combined deterministic hash of all message content
- **Master Forensic Hash**: Combined complete hash of all API responses
- **Timestamp Verification**: UTC timestamps for each operation

### Data Verification
- **Content Verification**: Deterministic hashes ensure message content integrity
- **Forensic Verification**: Complete hashes provide full chain of custody
- **Independent Verification**: Third parties can verify both content and forensic integrity
- **Complete Documentation**: Audit trail with dual verification system
- **Professional Standards**: Meets forensic and business documentation standards

### Why Dual Hashes?
- **Content Hash**: Solves the problem of varying hashes for identical data
- **Forensic Hash**: Maintains complete forensic integrity including all metadata
- **Legal Compliance**: Provides both content verification and complete forensic trail
- **Professional Use**: Suitable for legal, business, and forensic applications

## FAQ - Hash Behavior

### "Why do some hashes change between exports?"
This is **expected and correct behavior** due to the dual hash system:

- **Content Hash**: Should remain **identical** for the same conversation data
- **Forensic Hash**: May **vary** due to dynamic server metadata (timestamps, pagination URLs, etc.)

**Think of it this way**: If you take two photos of the same document at different times, the document content (Content Hash) stays the same, but the lighting and camera settings (Forensic Hash) might be different.

### "Which hash should I use for verification?"
- **Content Hash**: Use for data integrity verification and comparison
- **Forensic Hash**: Use for complete forensic chain of custody

**Simple rule**: 
- Want to check if messages changed? → Use **Content Hash**
- Want complete digital evidence? → Use **Forensic Hash**

### "Is it normal for forensic hashes to change?"
**Yes**, forensic hashes may change because they include:
- Server response timestamps
- Dynamic pagination URLs
- API metadata that varies between requests
- Headers and response metadata

**This is actually good!** It means the system is capturing all the digital evidence, including when and how the data was retrieved.

### "What if both hashes are different?"
If **both** Content Hash and Forensic Hash are different, it likely means:
- New messages were added to the conversation
- Existing messages were edited
- The conversation content actually changed between exports

This is normal if the conversation is active and being used.

### "What if Content Hash is the same but Forensic Hash is different?"
This is **perfect and expected**! It means:
- ✅ The conversation content is identical
- ✅ The system is working correctly
- ✅ Server metadata changed (which is normal)

### "How do I know if my export is authentic?"
Check the **Content Hash**:
- If it matches between exports → Content is authentic
- If it's different → Content may have changed

The **Forensic Hash** provides additional proof that you have complete digital evidence.

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
│   ├── complete_conversation_YYYYMMDD_HHMMSS.json
│   └── teams_conversation_YYYYMMDD_HHMMSS.pdf
```

## Legal Notice

This tool exports conversations using the official Microsoft Graph API and implements integrity verification protocols. The generated PDF includes cryptographic integrity certificates for data verification. Use responsibly and in compliance with your organization's policies and applicable laws.

### Integrity Features
- **Dual Hash System**: Content hash (deterministic) + Forensic hash (complete)
- **Cryptographic Verification**: SHA-256 hashes of both content and full responses
- **Data Integrity**: Complete audit trail from API to PDF with dual verification
- **Professional Standards**: Meets forensic and business documentation standards
- **Independent Verification**: Third-party validation of both content and forensic integrity

## Support

For issues or questions:
1. Check the troubleshooting section above
2. Verify your Chat ID and access token
3. Ensure you have proper permissions for the conversation
4. For legal use, consult with qualified legal professionals

---

**Note**: This tool is designed for personal and authorized business use with dual hash integrity verification. Always respect privacy and data protection regulations. The dual hash system ensures both content verification and complete forensic integrity for legal and business documentation.

**💡 Simple Summary**: Think of it like this - Content Hash tells you "what was said" and Forensic Hash tells you "everything about how it was captured." Both are important for proving authenticity and maintaining complete records. 