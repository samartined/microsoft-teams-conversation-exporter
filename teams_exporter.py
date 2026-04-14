import json
import os
import sys
import argparse
import requests
import hashlib
import re
from datetime import datetime, timezone
from tqdm import tqdm
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors

def get_chat_id_from_user():
    """
    Guides the user to obtain the Chat ID using Graph Explorer with participants info
    """
    print("=== CHAT ID OBTAINMENT USING GRAPH EXPLORER ===")
    print("You need to provide the Chat ID of the conversation you want to export.")
    print("Follow these steps to find it using Graph Explorer:")
    print()
    
    print("📋 METHOD 1: Get chats with participants (RECOMMENDED)")
    print("-" * 50)
    print("1. Go to https://developer.microsoft.com/en-us/graph/graph-explorer")
    print("2. Sign in with your Microsoft account")
    print("3. Execute this query: GET https://graph.microsoft.com/v1.0/me/chats?$expand=members")
    print("4. This will show participants for each chat")
    print("5. Look for the chat with the person you want to export")
    print("6. Copy the 'id' field from that chat")
    print()
    
    print("📋 METHOD 2: Check participants for each chat")
    print("-" * 50)
    print("1. First use: GET https://graph.microsoft.com/v1.0/me/chats")
    print("2. For each chat ID, check participants with:")
    print("   GET https://graph.microsoft.com/v1.0/chats/{chat-id}/members")
    print("3. This will show you who is in each conversation")
    print()
    
    print("📋 METHOD 3: Use Teams URL (if you know the conversation)")
    print("-" * 50)
    print("1. Open the conversation in Teams")
    print("2. Copy the URL from your browser")
    print("3. Extract the Chat ID from the URL")
    print("   Example: https://teams.microsoft.com/l/chat/19:xxxxx...")
    print("   Chat ID: 19:xxxxx... (the part after /chat/)")
    print()
    
    print("⚠️  IMPORTANT NOTES:")
    print("- Chat IDs are long strings with format: 19:xxxxx...")
    print("- They usually end with @unq.gbl.spaces or similar")
    print("- For one-on-one chats, topic is null, so check participants")
    print("- Group chats show topic names")
    print("- Meeting chats show meeting names")
    print("- Make sure you have access to the conversation")
    print("- The conversation must be accessible via Microsoft Graph API")
    print()
    
    print("📝 ENTER CHAT ID")
    print("-" * 20)
    chat_id = input("Paste the Chat ID here: ").strip()
    
    if not chat_id:
        print("❌ No Chat ID provided. Using default for testing.")
        print("No Chat ID provided. Exiting...")
        exit()
    
    # Basic validation
    if not chat_id.startswith("19:"):
        print("⚠️  Warning: Chat ID doesn't start with '19:'. This might be incorrect.")
        confirm = input("Continue anyway? (y/n): ").strip().lower()
        if confirm != 'y':
            return get_chat_id_from_user()
    
    print(f"✅ Chat ID accepted: {chat_id}")
    return chat_id

def get_token_from_browser():
    """
    Instructions to get Graph Explorer token
    """
    print("=== GRAPH EXPLORER TOKEN OBTAINMENT ===")
    print("Follow these steps to get the token:")
    print()
    print("1. Go to https://developer.microsoft.com/en-us/graph/graph-explorer")
    print("2. Sign in with your Microsoft account")
    print("3. Look for the 'Access token' section on the right side")
    print("4. Click on 'Copy' next to the token")
    print("5. The token will be copied to your clipboard")
    print()
    
    token = input("Paste the token here: ").strip()
    
    # Clean token and add Bearer prefix
    if token.startswith("Bearer "):
        token = token[7:]  # Remove Bearer if user included it
    
    if not token:
        print("❌ No token provided")
        return None
    
    # Add Bearer prefix
    token = f"Bearer {token}"
    
    return token

def extract_participants_from_messages(chat_id, token):
    """
    Fallback method: extract participant names from messages
    """
    headers = {
        'Authorization': token,
        'Content-Type': 'application/json'
    }
    
    url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages?$top=50"
    
    try:
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            data = response.json()
            participants = set()
            
            if 'value' in data:
                for message in data['value']:
                    from_info = message.get('from')
                    if from_info and 'user' in from_info:
                        user_info = from_info['user']
                        display_name = user_info.get('displayName', '')
                        if display_name and display_name.strip():
                            participants.add(display_name)
            
            if participants:
                print(f"✅ Extracted {len(participants)} participants from messages")
                return [{'displayName': name, 'email': 'From messages'} for name in participants]
            else:
                print("❌ No participants found in messages")
                return []
        else:
            print(f"❌ Cannot access messages: {response.status_code}")
            return []
            
    except Exception as e:
        print(f"❌ Error extracting from messages: {str(e)[:50]}...")
        return []

def get_chat_participants(chat_id, token):
    """
    Get participants information for a specific chat
    """
    headers = {
        'Authorization': token,
        'Content-Type': 'application/json'
    }
    
    # Try API endpoints for participants (only valid ones)
    approaches = [
        f"https://graph.microsoft.com/v1.0/chats/{chat_id}/members",
        f"https://graph.microsoft.com/v1.0/chats/{chat_id}?$expand=members"
    ]
    
    for i, url in enumerate(approaches, 1):
        try:
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                data = response.json()
                participants = []
                
                # Extract participants from response
                if 'value' in data:
                    for member in data['value']:
                        display_name = member.get('displayName', '')
                        email = member.get('email', '')
                        if display_name and display_name.strip():
                            participants.append({
                                'displayName': display_name,
                                'email': email
                            })
                elif 'members' in data:
                    for member in data['members']:
                        display_name = member.get('displayName', '')
                        email = member.get('email', '')
                        if display_name and display_name.strip():
                            participants.append({
                                'displayName': display_name,
                                'email': email
                            })
                
                if participants:
                    print(f"✅ Retrieved {len(participants)} participants from API")
                    return participants
                    
            elif response.status_code == 403:
                print(f"⚠️  Insufficient permissions for participants API")
            elif response.status_code == 401:
                print(f"❌ Token authentication failed")
                break
                
        except Exception as e:
            print(f"⚠️  API error: {str(e)[:50]}...")
    
    # Fallback: extract from messages
    print("🔄 Extracting participants from messages...")
    return extract_participants_from_messages(chat_id, token)

def sanitize_folder_name(name: str) -> str:
    """
    Strips filesystem-illegal characters and normalises whitespace.
    Returns 'unknown_chat' if the result would be empty.
    """
    sanitized = re.sub(r'[/\\:*?"<>|]', '', name)
    sanitized = re.sub(r'\s+', ' ', sanitized).strip()
    return sanitized if sanitized else 'unknown_chat'


def derive_chat_name(chat: dict) -> str:
    """
    Returns a human-readable, filesystem-safe name for a chat dict
    as returned by GET /me/chats?$expand=members.

    Priority:
      1. topic field (group / meeting chats)
      2. member displayNames joined with ' - ' (1:1 chats where topic is null)
      3. truncated chat id (fallback)
    """
    topic = (chat.get('topic') or '').strip()
    if topic:
        return sanitize_folder_name(topic)

    members = chat.get('members') or []
    names = [
        m.get('displayName', '').strip()
        for m in members
        if m.get('displayName', '').strip()
    ]
    if names:
        return sanitize_folder_name(' - '.join(names))

    # Last resort: truncated chat id
    return sanitize_folder_name((chat.get('id') or 'unknown')[:40])


def _request_with_retry(url, headers, max_retries=3):
    """
    Performs a GET request and retries up to max_retries times on HTTP 429,
    honoring the Retry-After response header (defaults to 60s if absent).
    Returns the final response object regardless of status.
    """
    import time as _time
    for attempt in range(max_retries):
        response = requests.get(url, headers=headers)
        if response.status_code != 429:
            return response
        retry_after = int(response.headers.get('Retry-After', 60))
        print(f"  ⏳ Rate limited (429). Waiting {retry_after}s before retry "
              f"{attempt + 1}/{max_retries}...")
        _time.sleep(retry_after)
    return response  # return last response even if still 429


def get_all_chats(token: str) -> list:
    """
    Fetches all chats for the current user from Microsoft Graph API.
    Paginates automatically via @odata.nextLink (1s delay between pages).
    Returns a list of chat dicts with at minimum: id, topic, chatType, members.
    """
    import time as _time
    headers = {
        'Authorization': token,
        'Content-Type': 'application/json'
    }

    url = "https://graph.microsoft.com/v1.0/me/chats?$expand=members"
    all_chats = []

    print("📋 Fetching list of all chats...")

    while url:
        response = _request_with_retry(url, headers)

        if response.status_code == 200:
            data = response.json()
            page_chats = data.get('value', [])
            all_chats.extend(page_chats)
            print(f"  Fetched {len(page_chats)} chats (total: {len(all_chats)})")
            url = data.get('@odata.nextLink')
            if url:
                _time.sleep(1)  # avoid hammering the chat-list endpoint
        elif response.status_code == 401:
            print("❌ Token expired or invalid while fetching chats")
            break
        elif response.status_code == 403:
            print("❌ Insufficient permissions to list chats. "
                  "Ensure Chat.Read or Chat.ReadBasic scope is granted.")
            break
        else:
            print(f"❌ HTTP {response.status_code} while fetching chats: "
                  f"{response.text[:200]}")
            break

    print(f"✅ Total chats found: {len(all_chats)}")
    return all_chats


def export_all_chats(token: str, base_dir: str = "exported_messages",
                     language: str = "en", delay: float = 3.0) -> None:
    """
    Fetches all chats and exports each one to a dedicated subfolder under base_dir.

    Output structure:
        base_dir/
            {chat_name}/
                complete_conversation_{timestamp}.json
                teams_conversation_{timestamp}.pdf

    Per-chat errors are caught and logged; the full run continues.
    Prints a summary at the end.
    """
    import time as _time

    chats = get_all_chats(token)

    if not chats:
        print("❌ No chats found or unable to fetch chats. Exiting.")
        return

    total = len(chats)
    exported_count = 0
    failed_count = 0
    failed_chats = []

    print(f"\n🚀 Starting bulk export of {total} chats...")
    print("=" * 60)

    for index, chat in enumerate(chats, start=1):
        chat_id = chat.get('id', '')
        chat_name = derive_chat_name(chat)
        chat_output_dir = os.path.join(base_dir, chat_name)

        print(f"\n[{index}/{total}] {chat_name}")
        print(f"  Chat ID : {chat_id}")
        print(f"  Folder  : {chat_output_dir}")

        try:
            result_file, chain_of_custody, participant_names = export_messages(
                chat_id, token, chat_output_dir
            )

            if not result_file:
                raise RuntimeError("No messages retrieved or API error")

            pdf_timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            pdf_output = os.path.join(
                chat_output_dir, f"teams_conversation_{pdf_timestamp}.pdf"
            )

            convert_json_to_pdf(
                result_file, chain_of_custody, participant_names,
                language, pdf_output
            )

            exported_count += 1
            print(f"  ✅ Done: {chat_name}")

        except Exception as e:
            failed_count += 1
            failed_chats.append((chat_name, str(e)))
            print(f"  ❌ Failed: {chat_name} — {str(e)[:120]}")

        # Inter-chat delay (skip after the last chat)
        if index < total:
            print(f"  ⏳ Waiting {delay}s before next chat...")
            _time.sleep(delay)

    # Final summary
    print("\n" + "=" * 60)
    print("BULK EXPORT COMPLETE")
    print("=" * 60)
    print(f"  Total chats  : {total}")
    print(f"  Exported OK  : {exported_count}")
    print(f"  Failed       : {failed_count}")

    if failed_chats:
        print("\nFailed chats:")
        for name, err in failed_chats:
            print(f"  - {name}: {err[:100]}")

    print(f"\n📁 All exported files are under: {os.path.abspath(base_dir)}")


def create_dual_hashes(response, page_number):
    """
    Crea hashes duales para diferentes niveles de integridad forense
    """
    response_data = response.json()
    messages = response_data.get('value', [])
    
    clean_messages = []
    for message in messages:
        clean_message = {
            'id': message.get('id'),
            'createdDateTime': message.get('createdDateTime'),
            'lastModifiedDateTime': message.get('lastModifiedDateTime'),  # ✅ INCLUIR
            'subject': message.get('subject'),
            'importance': message.get('importance'),
            'replyToId': message.get('replyToId')
        }
        
        # Información del remitente
        from_info = message.get('from')
        if from_info and 'user' in from_info:
            user_info = from_info['user']
            clean_message['from'] = {
                'displayName': user_info.get('displayName'),
                'id': user_info.get('id')
            }
        
        # Contenido del mensaje
        body_info = message.get('body')
        if body_info:
            clean_message['body'] = {
                'content': body_info.get('content'),
                'contentType': body_info.get('contentType')
            }
        
        clean_messages.append(clean_message)
    
    # Hash determinístico incluyendo lastModifiedDateTime
    messages_data = json.dumps(clean_messages, sort_keys=True, separators=(',', ':'))
    content_hash = hashlib.sha256(messages_data.encode()).hexdigest()
    
    # Hash forense completo
    forensic_hash = hashlib.sha256(response.content).hexdigest()
    
    return {
        "content_hash": content_hash,
        "forensic_hash": forensic_hash,
        "messages_count": len(messages)
    }

def export_messages(chat_id, token, output_dir="exported_messages"):
    """
    Export: saves all messages in memory and creates final file directly
    Returns: (file_path, chain_of_custody_data, participants)
    """
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    print("=== EXPORT WITH TOKEN ===")
    print(f"Chat ID: {chat_id}")
    print(f"Directory: {output_dir}")
    print("=" * 50)
    
    headers = {
        'Authorization': token,
        'Content-Type': 'application/json'
    }
    
    all_messages = []
    page_hashes = []  # Lista de hashes de cada página
    page = 1
    next_link = None
    
    # Get participants information
    print("👥 Getting participants information...")
    participants = get_chat_participants(chat_id, token)
    
    if participants:
        participant_names = [p['displayName'] for p in participants]
        print(f"✅ Found {len(participants)} participants: {', '.join(participant_names)}")
    else:
        print("⚠️  Could not retrieve participants information")
        participant_names = ["Unknown Participants"]
    
    # Chain of custody metadata - CORREGIDO: usar datetime.now(timezone.utc)
    export_timestamp = datetime.now(timezone.utc).isoformat()
    session_metadata = {
        "chat_id": chat_id,
        "export_timestamp": export_timestamp,
        "api_endpoint": f"/chats/{chat_id}/messages",
        "token_scope": "Chat.Read",
        "user_agent": "Microsoft Graph API",
        "export_method": "Microsoft Graph API v1.0"
    }
    
    print("🚀 Starting export...")
    
    # Loading indicator with dots
    import itertools
    import time
    dots = itertools.cycle(['⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏'])
    start_time = time.time()
    
    while True:
        if next_link:
            url = next_link
        else:
            url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages?$top=50"
        
        try:
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                # Crear hashes duales para esta página
                dual_hashes = create_dual_hashes(response, page)
                
                # Guardar hash de esta página con sistema dual
                page_hash = {
                    "page": page,
                    "url": url,
                    "timestamp": datetime.now(timezone.utc).isoformat(),
                    "status_code": response.status_code,
                    "content_hash": dual_hashes["content_hash"],      # Hash determinístico
                    "forensic_hash": dual_hashes["forensic_hash"],    # Hash forense completo
                    "messages_count": dual_hashes["messages_count"]
                }
                page_hashes.append(page_hash)
                
                data = response.json()
                
                if 'value' in data:
                    messages_in_page = len(data['value'])
                    all_messages.extend(data['value'])
                    
                    # Show loading with time-based spinner update
                    elapsed = time.time() - start_time
                    spinner_char = next(dots)
                    print(f"\r📄 Exporting {spinner_char} Page {page} - {len(all_messages)} messages ({elapsed:.1f}s)", end="", flush=True)
                
                # Check if there are more pages
                next_link = data.get('@odata.nextLink')
                
                if next_link:
                    page += 1
                    time.sleep(1)  # Pause to avoid API overload
                else:
                    print(f"\n🎉 Export completed! Total: {len(all_messages)} messages")
                    break
                    
            elif response.status_code == 401:
                print("\n❌ Error: Token expired or invalid")
                print("You need to get a new token from Graph Explorer")
                break
            else:
                print(f"\n❌ HTTP Error {response.status_code}: {response.text}")
                break
                
        except Exception as e:
            print(f"\n❌ Request error: {e}")
            break
    
    # Create chain of custody data - CORREGIDO: usar datetime.now(timezone.utc) y eliminar datos sensibles
    chain_of_custody = {
        "peritaje_info": {
            "page_hashes": page_hashes,
            "session_metadata": session_metadata,
            "total_messages": len(all_messages),
            "total_pages": page,
            "export_completion_time": datetime.now(timezone.utc).isoformat(),
            "perito_responsable": "Automated Export System",  # Genérico
            "metodo_captura": "Microsoft Graph API"  # Genérico
        }
    }
    
    # Calculate master hash from all content hashes (deterministic)
    all_content_hashes_combined = "".join([ph["content_hash"] for ph in page_hashes])
    chain_of_custody["peritaje_info"]["master_content_hash"] = hashlib.sha256(all_content_hashes_combined.encode()).hexdigest()
    
    # Calculate forensic master hash from all forensic hashes (complete integrity)
    all_forensic_hashes_combined = "".join([ph["forensic_hash"] for ph in page_hashes])
    chain_of_custody["peritaje_info"]["master_forensic_hash"] = hashlib.sha256(all_forensic_hashes_combined.encode()).hexdigest()
    
    # Save final file directly (no temporary files)
    if all_messages:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        final_file = os.path.join(output_dir, f"complete_conversation_{timestamp}.json")
        integrity_file = os.path.join(output_dir, f"integrity_{timestamp}.json")

        print(f"\n💾 Saving file...")
        with tqdm(total=1, desc="💾 Saving", unit="file") as pbar:
            with open(final_file, 'w', encoding='utf-8') as f:
                json.dump(all_messages, f, ensure_ascii=False, indent=2)
            pbar.update(1)

        # Save chain of custody as standalone integrity file for future verification
        with open(integrity_file, 'w', encoding='utf-8') as f:
            json.dump(chain_of_custody, f, ensure_ascii=False, indent=2)
        
        print(f"\n🎉 EXPORT COMPLETED")
        print(f"📁 Final file: {final_file}")
        print(f"📊 Total messages: {len(all_messages)}")
        print(f"📄 Pages processed: {page}")
        print(f"✅ No temporary files created")
        
        # Show statistics
        first_msg = all_messages[0]
        last_msg = all_messages[-1]
        print(f"📅 First message: {first_msg.get('createdDateTime', 'N/A')}")
        print(f"📅 Last message: {last_msg.get('createdDateTime', 'N/A')}")
        
        return final_file, chain_of_custody, participant_names
    else:
        print("❌ No messages retrieved")
        return None, None, None

def clean_html_content(html_content):
    """
    Cleans HTML content from Teams to make it compatible with ReportLab
    """
    if not html_content:
        return ""
    
    # Remove complex HTML tags and keep only text
    # Remove tags with styles
    content = re.sub(r'<[^>]*style="[^"]*"[^>]*>', '', html_content)
    
    # Remove HTML tags but keep text
    content = re.sub(r'<[^>]+>', '', content)
    
    # Decode common HTML entities
    content = content.replace('&nbsp;', ' ')
    content = content.replace('&amp;', '&')
    content = content.replace('&lt;', '<')
    content = content.replace('&gt;', '>')
    content = content.replace('&quot;', '"')
    
    # Clean extra spaces
    content = re.sub(r'\s+', ' ', content).strip()
    
    return content

def load_language_config():
    """
    Loads language configuration from JSON file
    """
    try:
        with open('language_config.json', 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print("❌ language_config.json not found. Using default English.")
        return {
            "en": {
                "document_title": "MICROSOFT TEAMS CONVERSATION",
                "document_subtitle": "Official certified export",
                "export_date": "Export date:",
                "original_file": "Original file:",
                "total_messages": "Total messages:",
                "conversation_with": "Conversation Participants:",
                "chain_of_custody": "CHAIN OF CUSTODY CERTIFICATE",
                "api_response_hash": "API Response Hash:",
                "session_metadata": "Session Metadata:",
                "authenticity_declaration": "AUTHENTICITY DECLARATION",
                "authenticity_text": "This conversation has been officially exported from Microsoft Teams using the official Microsoft Graph API. The chain of custody includes original API responses, session metadata, and cryptographic hashes to guarantee data integrity and authenticity.",
                "complete_conversation": "COMPLETE CONVERSATION",
                "message": "MESSAGE",
                "from": "From:",
                "content": "Content:",
                "unidentified_user": "Unidentified user",
                "unknown_user": "Unknown user",
                "final_certificate": "FINAL CERTIFICATE",
                "certification_date": "Certification date:",
                "final_certificate_text": "This document is a faithful conversion of the original data. The chain of custody provides cryptographic proof of data integrity and authenticity."
            }
        }

def select_language():
    """
    Language selection for PDF content
    """
    print("🌍 SELECT PDF LANGUAGE")
    print("-" * 25)
    print("Available languages:")
    print("1. English (en)")
    print("2. Spanish (es)")
    print("3. French (fr)")
    print("4. German (de)")
    print()
    
    while True:
        choice = input("Enter language number or code (default: 1): ").strip()
        
        if not choice:
            return "en"
        
        # Handle numeric input
        if choice == "1":
            return "en"
        elif choice == "2":
            return "es"
        elif choice == "3":
            return "fr"
        elif choice == "4":
            return "de"
        
        # Handle direct language codes
        if choice in ["en", "es", "fr", "de"]:
            return choice
        
        print("❌ Invalid choice. Please enter 1-4 or en/es/fr/de")

def convert_json_to_pdf(json_file, chain_of_custody, participant_names, language="en", output_file=None):
    """
    Converts Teams JSON to PDF with real chain of custody certificate
    """
    print(f"\n📄 Converting {json_file} to PDF in {language}...")
    
    # Generate output filename with timestamp if not provided
    if output_file is None:
        # Ensure output directory exists
        output_dir = "exported_messages"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = f"{output_dir}/teams_conversation_{timestamp}.pdf"
    
    # Load language configuration
    lang_config = load_language_config()
    texts = lang_config.get(language, lang_config["en"])  # Default to English
    
    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # Create PDF document
    doc = SimpleDocTemplate(output_file, pagesize=A4)
    styles = getSampleStyleSheet()
    story = []
    
    # Custom style for headers
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=16,
        spaceAfter=30,
        alignment=1  # Centered
    )
    
    # Style for messages
    message_style = ParagraphStyle(
        'MessageStyle',
        parent=styles['Normal'],
        fontSize=10,
        spaceAfter=12,
        leftIndent=20
    )
    
    # Style for hash and metadata (smaller font)
    metadata_style = ParagraphStyle(
        'MetadataStyle',
        parent=styles['Normal'],
        fontSize=8,
        fontName='Courier'
    )
    
    # Style for hash explanation (smaller than heading but larger than normal)
    hash_explanation_style = ParagraphStyle(
        'HashExplanationStyle',
        parent=styles['Normal'],
        fontSize=9,
        fontName='Helvetica-Bold'
    )
    
    # Document header
    story.append(Paragraph(texts["document_title"], title_style))
    story.append(Paragraph(texts["document_subtitle"], styles['Heading2']))
    story.append(Spacer(1, 20))
    
    # Document information - CORREGIDO: usar nombres reales de participantes
    participants_text = ', '.join(participant_names) if participant_names else 'Unknown Participants'
    
    info_data = [
        [texts["export_date"], datetime.now().strftime('%d/%m/%Y %H:%M:%S')],
        [texts["original_file"], os.path.basename(json_file)],
        [texts["total_messages"], str(len(data))],
        [texts["conversation_with"], participants_text]
    ]
    
    info_table = Table(info_data, colWidths=[2*inch, 4*inch])
    info_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.grey),
        ('TEXTCOLOR', (0, 0), (0, -1), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ('BACKGROUND', (1, 0), (1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    story.append(info_table)
    story.append(Spacer(1, 20))
    
    # Chain of custody certificate
    story.append(Paragraph(texts["chain_of_custody"], styles['Heading2']))
    
    if chain_of_custody and "peritaje_info" in chain_of_custody:
        peritaje_info = chain_of_custody["peritaje_info"]
        
        # Session metadata
        metadata = peritaje_info["session_metadata"]
        story.append(Paragraph(f"{texts['session_metadata']}:", styles['Normal']))
        story.append(Paragraph(f"Chat ID: {metadata['chat_id']}", metadata_style))
        story.append(Paragraph(f"Export Time: {metadata['export_timestamp']}", metadata_style))
        story.append(Paragraph(f"API Endpoint: {metadata['api_endpoint']}", metadata_style))
        story.append(Paragraph(f"Total Pages: {peritaje_info['total_pages']}", metadata_style))
        story.append(Spacer(1, 10))
        
        # Hash explanation section
        story.append(Paragraph(texts["hash_explanation"], hash_explanation_style))
        story.append(Paragraph(texts["content_hash_explanation"], styles['Normal']))
        story.append(Paragraph(texts["forensic_hash_explanation"], styles['Normal']))
        story.append(Spacer(1, 10))
        
        # Page hashes with dual system
        story.append(Paragraph("Page Hashes:", styles['Normal']))
        page_hashes = peritaje_info["page_hashes"]
        for page_hash in page_hashes:
            story.append(Paragraph(
                f"Page {page_hash['page']}:", 
                styles['Normal']
            ))
            story.append(Paragraph(
                f"  {texts['content_hash']} {page_hash['content_hash']}", 
                metadata_style
            ))
            story.append(Paragraph(
                f"  {texts['forensic_hash']} {page_hash['forensic_hash']}", 
                metadata_style
            ))
            story.append(Spacer(1, 5))
        
        story.append(Spacer(1, 10))
        
        # Master hashes
        master_content_hash = peritaje_info.get("master_content_hash", "N/A")
        master_forensic_hash = peritaje_info.get("master_forensic_hash", "N/A")
        
        story.append(Paragraph(f"Master Content Hash (deterministic): {master_content_hash}", metadata_style))
        story.append(Paragraph(f"Master Forensic Hash (complete): {master_forensic_hash}", metadata_style))
    
    story.append(Spacer(1, 20))
    
    # Authenticity declaration
    story.append(Paragraph(texts["authenticity_declaration"], styles['Heading2']))
    story.append(Paragraph(texts["authenticity_text"], styles['Normal']))
    story.append(Spacer(1, 20))
    
    # Conversation
    story.append(Paragraph(texts["complete_conversation"], styles['Heading2']))
    story.append(Spacer(1, 20))
    
    total_messages = len(data)
    print(f"📝 Processing {total_messages} messages...")
    
    # Create progress bar for message processing
    with tqdm(total=total_messages, desc="📝 Processing", unit="msg") as pbar:
        for i, message in enumerate(data, 1):
            # Extract message information
            created_date = message.get('createdDateTime', '')
            
            # Handle cases where 'from' is None
            from_info = message.get('from')
            if from_info is None:
                from_user = texts["unidentified_user"]
            else:
                user_info = from_info.get('user', {})
                from_user = user_info.get('displayName', texts["unknown_user"])
            
            body_info = message.get('body', {})
            body_content = body_info.get('content', '') if body_info else ''
            
            # Clean HTML content
            clean_content = clean_html_content(body_content)
            # Re-escape any residual XML-special chars so ReportLab's parser doesn't choke
            clean_content = clean_content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            from_user = from_user.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

            # Convert date
            try:
                dt = datetime.fromisoformat(created_date.replace('Z', '+00:00'))
                formatted_date = dt.strftime('%d/%m/%Y %H:%M:%S')
            except:
                formatted_date = created_date

            # Create message in PDF (without HTML)
            message_text = f"<b>{texts['message']} {i}</b> - {formatted_date}<br/>"
            message_text += f"<b>{texts['from']}</b> {from_user}<br/>"
            message_text += f"<b>{texts['content']}</b><br/>{clean_content}"
            
            story.append(Paragraph(message_text, message_style))
            story.append(Spacer(1, 10))
            
            # Add separator line every 10 messages
            if i % 10 == 0:
                story.append(Paragraph("_" * 80, styles['Normal']))
                story.append(Spacer(1, 10))
            
            # Update progress bar
            pbar.update(1)
    
    # Final certificate
    story.append(Paragraph(texts["final_certificate"], styles['Heading2']))
    story.append(Paragraph(
        f"{texts['certification_date']} {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}<br/>"
        f"{texts['final_certificate_text']}",
        styles['Normal']
    ))
    
    # Generate PDF
    print(f"\n📄 Generating PDF...")
    with tqdm(total=1, desc="📄 Generating", unit="file") as pbar:
        doc.build(story)
        pbar.update(1)
    
    print(f"\n✅ PDF generated: {output_file}")
    return output_file


def recompute_master_content_hash(messages: list, page_size: int = 50) -> str:
    """
    Re-computes the master_content_hash from a list of messages using the exact
    same algorithm as create_dual_hashes() + export_messages(), so it can be
    compared against the value stored in integrity_{timestamp}.json.
    """
    pages = [messages[i:i + page_size] for i in range(0, len(messages), page_size)]
    if not pages:
        pages = [[]]

    page_content_hashes = []
    for page_messages in pages:
        clean_messages = []
        for message in page_messages:
            clean_message = {
                'id': message.get('id'),
                'createdDateTime': message.get('createdDateTime'),
                'lastModifiedDateTime': message.get('lastModifiedDateTime'),
                'subject': message.get('subject'),
                'importance': message.get('importance'),
                'replyToId': message.get('replyToId')
            }
            from_info = message.get('from')
            if from_info and 'user' in from_info:
                user_info = from_info['user']
                clean_message['from'] = {
                    'displayName': user_info.get('displayName'),
                    'id': user_info.get('id')
                }
            body_info = message.get('body')
            if body_info:
                clean_message['body'] = {
                    'content': body_info.get('content'),
                    'contentType': body_info.get('contentType')
                }
            clean_messages.append(clean_message)

        messages_data = json.dumps(clean_messages, sort_keys=True, separators=(',', ':'))
        page_content_hashes.append(hashlib.sha256(messages_data.encode()).hexdigest())

    combined = "".join(page_content_hashes)
    return hashlib.sha256(combined.encode()).hexdigest()


def validate_chat_export(chat_dir: str) -> dict:
    """
    Validates a single exported chat directory.

    Returns a dict:
      {
        "dir": str,
        "has_json": bool,
        "has_pdf": bool,
        "has_integrity": bool,
        "message_count": int or None,
        "hash_verified": bool or None,  # None = no integrity file to check
        "errors": [str]
      }
    """
    result = {
        "dir": chat_dir,
        "has_json": False,
        "has_pdf": False,
        "has_integrity": False,
        "message_count": None,
        "hash_verified": None,
        "errors": []
    }

    # Find files
    json_files = [f for f in os.listdir(chat_dir)
                  if f.startswith("complete_conversation_") and f.endswith(".json")]
    pdf_files  = [f for f in os.listdir(chat_dir) if f.endswith(".pdf")]
    int_files  = [f for f in os.listdir(chat_dir)
                  if f.startswith("integrity_") and f.endswith(".json")]

    result["has_json"] = bool(json_files)
    result["has_pdf"]  = bool(pdf_files)
    result["has_integrity"] = bool(int_files)

    if not json_files:
        result["errors"].append("No conversation JSON found")
        return result
    if not pdf_files:
        result["errors"].append("No PDF found")

    # Load and validate JSON
    json_path = os.path.join(chat_dir, sorted(json_files)[-1])  # latest
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            messages = json.load(f)
        if not isinstance(messages, list) or len(messages) == 0:
            result["errors"].append("JSON is empty or not a list")
            return result
        result["message_count"] = len(messages)
    except Exception as e:
        result["errors"].append(f"JSON parse error: {e}")
        return result

    # Cryptographic re-verification (only if integrity file exists)
    if int_files:
        int_path = os.path.join(chat_dir, sorted(int_files)[-1])  # latest
        try:
            with open(int_path, 'r', encoding='utf-8') as f:
                integrity = json.load(f)
            stored_hash = integrity["peritaje_info"]["master_content_hash"]
            recomputed  = recompute_master_content_hash(messages)
            result["hash_verified"] = (recomputed == stored_hash)
            if not result["hash_verified"]:
                result["errors"].append(
                    f"Hash MISMATCH — stored: {stored_hash[:16]}... "
                    f"recomputed: {recomputed[:16]}..."
                )
        except Exception as e:
            result["errors"].append(f"Integrity file error: {e}")
            result["hash_verified"] = False

    return result


def validate_exports(base_dir: str = "exported_messages") -> bool:
    """
    Validates all chat subdirectories under base_dir.
    Prints a report and returns True if everything passed, False otherwise.
    """
    if not os.path.isdir(base_dir):
        print(f"❌ Directory not found: {base_dir}")
        return False

    chat_dirs = sorted([
        os.path.join(base_dir, d)
        for d in os.listdir(base_dir)
        if os.path.isdir(os.path.join(base_dir, d))
    ])

    if not chat_dirs:
        print(f"⚠️  No chat subdirectories found in {base_dir}")
        return False

    print(f"\n{'=' * 60}")
    print(f"INTEGRITY VALIDATION REPORT")
    print(f"{'=' * 60}")
    print(f"Scanning {len(chat_dirs)} chat folder(s) in {os.path.abspath(base_dir)}\n")

    passed = 0
    failed = 0
    failed_dirs = []

    for chat_dir in chat_dirs:
        r = validate_chat_export(chat_dir)
        name = os.path.basename(chat_dir)

        if r["errors"]:
            status = "❌ FAIL"
            failed += 1
            failed_dirs.append((name, r["errors"]))
        else:
            if r["hash_verified"] is True:
                status = "✅ PASS (hash verified)"
            elif r["hash_verified"] is None:
                status = "✅ PASS (structural only — no integrity file)"
            else:
                status = "❌ FAIL (hash mismatch)"
                failed += 1
                failed_dirs.append((name, r["errors"]))
                continue
            passed += 1

        msgs = f"{r['message_count']} msgs" if r["message_count"] else "?"
        print(f"  {status} | {msgs:>8} | {name}")

    print(f"\n{'─' * 60}")
    print(f"  Total : {len(chat_dirs)}")
    print(f"  Passed: {passed}")
    print(f"  Failed: {failed}")

    if failed_dirs:
        print(f"\n⚠️  Issues found:")
        for name, errors in failed_dirs:
            for err in errors:
                print(f"    [{name}] {err}")

    all_passed = failed == 0
    print(f"\n{'✅ All exports validated successfully!' if all_passed else '❌ Validation found issues — see above.'}")
    print(f"{'=' * 60}\n")
    return all_passed


def parse_args():
    parser = argparse.ArgumentParser(description="Export Microsoft Teams conversations to PDF")
    parser.add_argument("--chat-id", help="Teams Chat ID (e.g. 19:xxx@unq.gbl.spaces)")
    parser.add_argument("--token", help="Microsoft Graph API access token")
    parser.add_argument("--language", choices=["en", "es", "fr", "de"], help="PDF language (default: en)")
    parser.add_argument("--all-chats", action="store_true", default=False,
                        help="Export all accessible chats instead of a single chat")
    parser.add_argument("--delay", type=float, default=3.0,
                        help="Seconds to wait between chat exports to avoid rate limiting (default: 3)")
    parser.add_argument("--validate", action="store_true", default=False,
                        help="Validate integrity of all exports without re-exporting")
    return parser.parse_args()


def main():
    """
    Main function - Complete automated process with export
    """
    args = parse_args()
    OUTPUT_DIR = "exported_messages"

    # --- STANDALONE VALIDATION MODE (no token needed) ---
    if args.validate and not args.all_chats and not args.chat_id:
        validate_exports(OUTPUT_DIR)
        return

    print("🚀 TEAMS EXPORT AUTOMATION")
    print("=" * 50)
    print("This script will:")
    print("1. Export all Teams messages to JSON")
    print("2. Convert JSON to certified PDF with chain of custody")
    print("3. Save integrity file for cryptographic verification")
    print("4. Validate integrity of exported files")
    print("=" * 50)

    # Get Chat ID: from CLI arg, interactive prompt, or skipped for --all-chats
    if args.chat_id:
        chat_id = args.chat_id.strip()
        print(f"✅ Chat ID (from argument): {chat_id}")
    elif not args.all_chats:
        chat_id = get_chat_id_from_user()
    else:
        chat_id = None  # not used in bulk path

    # Get token: from CLI arg or interactive prompt
    if args.token:
        raw_token = args.token.strip()
        if raw_token.startswith("Bearer "):
            raw_token = raw_token[7:]
        token = f"Bearer {raw_token}"
        print("✅ Token accepted (from argument)")
    else:
        token = get_token_from_browser()

    if not token or token == "Bearer ":
        print("❌ No valid token provided")
        return

    # --- BULK EXPORT BRANCH ---
    if args.all_chats:
        language = args.language if args.language else select_language()
        export_all_chats(token, base_dir=OUTPUT_DIR, language=language, delay=args.delay)
        # Auto-validate all exported chat folders
        print("\n🔍 STEP 4: INTEGRITY VALIDATION")
        print("-" * 40)
        validate_exports(OUTPUT_DIR)
        return

    # --- SINGLE CHAT EXPORT ---
    # Step 1: Export messages
    print(f"\n📤 STEP 1: EXPORTING MESSAGES")
    print("-" * 40)
    result_file, chain_of_custody, participant_names = export_messages(chat_id, token, OUTPUT_DIR)

    if not result_file:
        print("❌ Message export failed")
        return

    # Step 2: Select language for PDF
    print(f"\n🌍 STEP 2: SELECTING PDF LANGUAGE")
    print("-" * 30)
    if args.language:
        language = args.language
        print(f"✅ Language (from argument): {language}")
    else:
        language = select_language()

    # Step 3: Convert to PDF
    print(f"\n📄 STEP 3: CONVERTING TO PDF")
    print("-" * 30)
    pdf_file = convert_json_to_pdf(result_file, chain_of_custody, participant_names, language, None)

    if not pdf_file:
        print("❌ PDF conversion failed")
        return

    # Step 4: Validate integrity of this export
    print(f"\n🔍 STEP 4: INTEGRITY VALIDATION")
    print("-" * 30)
    validate_chat_export_result = validate_chat_export(OUTPUT_DIR)
    if validate_chat_export_result["errors"]:
        print(f"❌ Validation issues: {'; '.join(validate_chat_export_result['errors'])}")
    elif validate_chat_export_result["hash_verified"] is True:
        print(f"✅ Integrity verified — {validate_chat_export_result['message_count']} messages, hash matches")
    else:
        print(f"✅ Export complete — {validate_chat_export_result['message_count']} messages (no integrity file to verify against)")

    # Final summary
    print(f"\n🎉 COMPLETE PROCESS FINISHED!")
    print("=" * 45)
    print(f"📁 JSON file: {result_file}")
    print(f"📄 PDF file: {pdf_file}")
    print(f"🌍 Language: {language}")
    print(f"📊 Total messages: {len(json.load(open(result_file, 'r', encoding='utf-8')))}")
    print(f"📄 Total pages: {chain_of_custody['peritaje_info']['total_pages'] if chain_of_custody else 'N/A'}")
    print(f"🔒 Master Content Hash (deterministic): {chain_of_custody['peritaje_info'].get('master_content_hash', 'N/A') if chain_of_custody else 'N/A'}")
    print(f"🔒 Master Forensic Hash (complete): {chain_of_custody['peritaje_info'].get('master_forensic_hash', 'N/A') if chain_of_custody else 'N/A'}")
    print(f"👥 Participants: {', '.join(participant_names) if participant_names else 'N/A'}")

if __name__ == "__main__":
    main() 