import os
import tempfile
import win32com.client
from email.message import EmailMessage
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr, formatdate
from email import policy
import uuid

output_folder = "output"
os.makedirs(output_folder, exist_ok=True)

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

pst_path = os.path.abspath("input/emails.pst")

if not os.path.exists(pst_path):
    raise FileNotFoundError(f"PST file not found: {pst_path}")

pst_name = os.path.basename(pst_path)
already_mounted = False
for folder in outlook.Folders:
    if folder.Name == pst_name or pst_path in str(folder.StoreID):
        already_mounted = True
        pst_folder = folder
        print(f"PST file already mounted: {pst_name}")
        break

if not already_mounted:
    try:
        outlook.AddStore(pst_path)
        pst_folder = outlook.Folders.GetLast()
        print(f"Successfully mounted PST file: {pst_name}")
    except Exception as e:
        print(f"Error loading PST file: {e}")
        print(f"Trying to find existing PST in Outlook folders...")
        pst_folder = None
        for folder in outlook.Folders:
            if pst_name.lower() in folder.Name.lower():
                pst_folder = folder
                break
        if pst_folder is None:
            raise Exception(f"Could not load PST file. Make sure it's not corrupted and Outlook has access to it.")

def export_emails(folder, parent_path=""):
    if parent_path:
        folder_path = os.path.join(output_folder, parent_path, folder.Name)
    else:
        folder_path = os.path.join(output_folder, folder.Name)
    os.makedirs(folder_path, exist_ok=True)

    olMail = 43
    
    items = folder.Items
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception:
        pass
    
    for i in range(items.Count):
        try:
            item = items[i]
            
            if item.Class != olMail:
                continue
            
            subject = item.Subject if item.Subject else "No Subject"
            subject = "".join(c for c in subject if c.isalnum() or c in " _-")[:50]
            if not subject:
                subject = "No Subject"
            
            filename = os.path.join(folder_path, f"{subject}.eml")
            counter = 1
            while os.path.exists(filename):
                filename = os.path.join(folder_path, f"{subject}_{counter}.eml")
                counter += 1
            
            abs_filename = os.path.abspath(filename)
            
            email_basename = os.path.splitext(os.path.basename(filename))[0]
            attachments_folder = os.path.join(folder_path, f"{email_basename}_attachments")
            
            try:
                eml = EmailMessage()
                
                if item.SenderEmailAddress:
                    eml['From'] = formataddr((item.SenderName, item.SenderEmailAddress)) if item.SenderName else item.SenderEmailAddress
                elif item.SentOnBehalfOfName:
                    eml['From'] = item.SentOnBehalfOfName
                
                if item.To:
                    eml['To'] = item.To
                
                if item.CC:
                    eml['Cc'] = item.CC
                
                if item.BCC:
                    eml['Bcc'] = item.BCC
                
                if item.Subject:
                    eml['Subject'] = item.Subject
                
                if item.ReceivedTime:
                    try:
                        received_dt = item.ReceivedTime
                        eml['Date'] = formatdate(received_dt.timestamp())
                    except:
                        if item.SentOn:
                            sent_dt = item.SentOn
                            eml['Date'] = formatdate(sent_dt.timestamp())
                
                try:
                    if hasattr(item, 'PropertyAccessor'):
                        msg_id = item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F")
                        if msg_id:
                            eml['Message-ID'] = msg_id
                except:
                    pass
                
                body_text = item.Body
                body_format = item.BodyFormat
                
                valid_attachments = []
                if item.Attachments.Count > 0:
                    olByValue = 1
                    for att_idx in range(1, item.Attachments.Count + 1):
                        try:
                            attachment = item.Attachments[att_idx]
                            if attachment.Type == olByValue:
                                att_name = getattr(attachment, 'FileName', None) or f'attachment_{att_idx}'
                                valid_attachments.append((attachment, att_name))
                        except:
                            continue
                
                if valid_attachments:
                    os.makedirs(attachments_folder, exist_ok=True)
                    
                    if body_format == 2:
                        body_part = MIMEText(body_text, 'html', 'utf-8')
                    else:
                        body_part = MIMEText(body_text, 'plain', 'utf-8')
                    
                    multipart_msg = MIMEMultipart('mixed')
                    multipart_msg.attach(body_part)
                    
                    for key, value in eml.items():
                        multipart_msg[key] = value
                    
                    eml = multipart_msg
                    
                    temp_dir = tempfile.gettempdir()
                    image_counter = 0
                    for att_idx, (attachment, att_name) in enumerate(valid_attachments):
                        try:
                            temp_filename = f"{uuid.uuid4().hex}_{att_name}"
                            temp_path = os.path.join(temp_dir, temp_filename)
                            
                            attachment.SaveAsFile(temp_path)
                            
                            with open(temp_path, 'rb') as att_file:
                                att_data = att_file.read()
                            
                            content_type = 'application/octet-stream'
                            file_ext = ''
                            
                            if att_data.startswith(b'\x89PNG\r\n\x1a\n'):
                                content_type = 'image/png'
                                file_ext = '.png'
                            elif att_data.startswith(b'\xff\xd8\xff'):
                                content_type = 'image/jpeg'
                                file_ext = '.jpg'
                            elif att_data.startswith(b'GIF87a') or att_data.startswith(b'GIF89a'):
                                content_type = 'image/gif'
                                file_ext = '.gif'
                            elif att_data.startswith(b'%PDF'):
                                content_type = 'application/pdf'
                                file_ext = '.pdf'
                            elif att_data.startswith(b'PK\x03\x04'):
                                if att_name.endswith('.docx') or 'word' in att_name.lower():
                                    content_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                                    file_ext = '.docx'
                                elif att_name.endswith('.xlsx') or 'excel' in att_name.lower():
                                    content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                                    file_ext = '.xlsx'
                                else:
                                    content_type = 'application/zip'
                                    file_ext = '.zip'
                            else:
                                ext = os.path.splitext(att_name)[1].lower()
                                content_type_map = {
                                    '.pdf': 'application/pdf',
                                    '.doc': 'application/msword',
                                    '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                                    '.xls': 'application/vnd.ms-excel',
                                    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                    '.jpg': 'image/jpeg',
                                    '.jpeg': 'image/jpeg',
                                    '.png': 'image/png',
                                    '.gif': 'image/gif',
                                    '.txt': 'text/plain',
                                    '.zip': 'application/zip',
                                }
                                content_type = content_type_map.get(ext, 'application/octet-stream')
                                file_ext = ext if ext else ''
                            has_extension = bool(os.path.splitext(att_name)[1])
                            is_guid_like = (len(att_name) == 36 and att_name.count('-') == 4) or (not has_extension and len(att_name) > 20)
                            
                            if is_guid_like or not has_extension:
                                if content_type.startswith('image/'):
                                    image_counter += 1
                                    att_name = f'image{image_counter}{file_ext}'
                                elif not has_extension:
                                    att_name = f'attachment{att_idx + 1}{file_ext}'
                            
                            maintype, subtype = content_type.split('/', 1) if '/' in content_type else ('application', 'octet-stream')
                            att_part = MIMEBase(maintype, subtype)
                            att_part.set_payload(att_data)
                            att_part.add_header('Content-Disposition', 'attachment', filename=att_name)
                            
                            from email import encoders
                            encoders.encode_base64(att_part)
                            
                            eml.attach(att_part)
                            
                            try:
                                att_save_path = os.path.join(attachments_folder, att_name)
                                att_save_counter = 1
                                att_base, att_ext = os.path.splitext(att_name)
                                while os.path.exists(att_save_path):
                                    att_save_path = os.path.join(attachments_folder, f"{att_base}_{att_save_counter}{att_ext}")
                                    att_save_counter += 1
                                
                                with open(temp_path, 'rb') as src, open(att_save_path, 'wb') as dst:
                                    dst.write(src.read())
                            except Exception as save_error:
                                print(f"  Warning: Could not save attachment copy {att_name}: {str(save_error)[:80]}")
                            
                            try:
                                os.remove(temp_path)
                            except:
                                pass
                        except Exception as att_error:
                            att_display = att_name if 'att_name' in locals() else 'unknown'
                            print(f"  Warning: Could not attach {att_display}: {str(att_error)[:80]}")
                            continue
                else:
                    if body_format == 2:
                        eml.set_content(body_text, subtype='html')
                    else:
                        eml.set_content(body_text, subtype='plain')
                
                with open(abs_filename, 'wb') as f:
                    if isinstance(eml, EmailMessage):
                        eml_bytes = eml.as_bytes(policy=policy.SMTP)
                    else:
                        eml_str = eml.as_string()
                        eml_bytes = eml_str.replace('\r\n', '\n').replace('\n', '\r\n').encode('utf-8')
                    f.write(eml_bytes)
                
                print(f"Exported: {subject}")
            except Exception as e:
                print(f"Failed to save {subject}: {str(e)[:100]}")
        except Exception as e:
            print(f"Skipped item: {str(e)[:100]}")

    for sub in folder.Folders:
        if parent_path:
            relative_path = os.path.join(parent_path, folder.Name)
        else:
            relative_path = folder.Name
        export_emails(sub, relative_path)

for subfolder in pst_folder.Folders:
    export_emails(subfolder, "")
