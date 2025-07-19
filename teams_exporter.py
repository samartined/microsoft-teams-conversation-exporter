import json
import os
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
                # Guardar hash de esta página original - CORREGIDO: usar datetime.now(timezone.utc)
                page_hash = {
                    "page": page,
                    "url": url,
                    "timestamp": datetime.now(timezone.utc).isoformat(),
                    "status_code": response.status_code,
                    "response_hash": hashlib.sha256(response.content).hexdigest(),
                    "messages_count": len(response.json().get('value', []))
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
    
    # Calculate master hash from all page hashes
    all_hashes_combined = "".join([ph["response_hash"] for ph in page_hashes])
    chain_of_custody["peritaje_info"]["master_hash"] = hashlib.sha256(all_hashes_combined.encode()).hexdigest()
    
    # Save final file directly (no temporary files)
    if all_messages:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        final_file = os.path.join(output_dir, f"complete_conversation_{timestamp}.json")
        
        print(f"\n💾 Saving file...")
        with tqdm(total=1, desc="💾 Saving", unit="file") as pbar:
            with open(final_file, 'w', encoding='utf-8') as f:
                json.dump(all_messages, f, ensure_ascii=False, indent=2)
            pbar.update(1)
        
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

def convert_json_to_pdf(json_file, chain_of_custody, participant_names, language="en", output_file="exported_messages/teams_conversation.pdf"):
    """
    Converts Teams JSON to PDF with real chain of custody certificate
    """
    print(f"\n📄 Converting {json_file} to PDF in {language}...")
    
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
        
        # Page hashes
        story.append(Paragraph("API Response Hashes (by page):", styles['Normal']))
        page_hashes = peritaje_info["page_hashes"]
        for page_hash in page_hashes:
            story.append(Paragraph(
                f"Page {page_hash['page']}: {page_hash['response_hash']}", 
                metadata_style
            ))
        
        story.append(Spacer(1, 10))
        
        # Master hash
        master_hash = peritaje_info["master_hash"]
        story.append(Paragraph(f"Master Hash: {master_hash}", metadata_style))
    
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

def main():
    """
    Main function - Complete automated process with export
    """
    print("🚀 TEAMS EXPORT AUTOMATION")
    print("=" * 50)
    print("This script will:")
    print("1. Export all Teams messages to JSON")
    print("2. Convert JSON to certified PDF with chain of custody")
    print("3. No temporary files created")
    print("=" * 50)
    
    # Get Chat ID from user with guidance
    chat_id = get_chat_id_from_user()
    OUTPUT_DIR = "exported_messages"
    
    # Get token
    token = get_token_from_browser()
    
    if not token or token == "Bearer ":
        print("❌ No valid token provided")
        return
    
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
    language = select_language()
    
    # Step 3: Convert to PDF
    print(f"\n📄 STEP 3: CONVERTING TO PDF")
    print("-" * 30)
    pdf_file = convert_json_to_pdf(result_file, chain_of_custody, participant_names, language)
    
    if not pdf_file:
        print("❌ PDF conversion failed")
        return
    
    # Final summary (no cleanup needed)
    print(f"\n🎉 COMPLETE PROCESS FINISHED!")
    print("=" * 45)
    print(f"📁 JSON file: {result_file}")
    print(f"📄 PDF file: {pdf_file}")
    print(f"🌍 Language: {language}")
    print(f"📊 Total messages: {len(json.load(open(result_file, 'r', encoding='utf-8')))}")
    print(f"📄 Total pages: {chain_of_custody['peritaje_info']['total_pages'] if chain_of_custody else 'N/A'}")
    print(f"🔒 Master hash: {chain_of_custody['peritaje_info']['master_hash'] if chain_of_custody else 'N/A'}")
    print(f"👥 Participants: {', '.join(participant_names) if participant_names else 'N/A'}")

if __name__ == "__main__":
    main() 