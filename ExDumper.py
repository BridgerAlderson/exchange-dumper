import argparse
import sys
from exchangelib import Credentials, Account, Configuration, DELEGATE, IMPERSONATION
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
import requests
import os
import logging
from datetime import datetime
import warnings
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
warnings.filterwarnings("ignore")

BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

class ExDumper:
    def __init__(self, domain, user, password, server, target):
        self.domain = domain
        self.user = user
        self.password = password
        self.server = server
        self.target = target
        self.credentials = Credentials(f"{domain}\\{user}", password)
        self.config = Configuration(server=server, credentials=self.credentials)
        self.access_type = None
        
    def detect_access_type(self):
        try:
            account = Account(
                primary_smtp_address=self.target,
                config=self.config,
                autodiscover=False,
                access_type=DELEGATE
            )
            account.root.refresh()
            self.access_type = DELEGATE
            return "DELEGATE"
        except Exception as e:
            try:
                account = Account(
                    primary_smtp_address=self.target,
                    config=self.config,
                    autodiscover=False,
                    access_type=IMPERSONATION
                )
                account.root.refresh()
                self.access_type = IMPERSONATION
                return "IMPERSONATION"
            except Exception as e2:
                raise Exception(f"No valid access: {str(e2)}")
    
    def create_dump_directory(self):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        dir_name = f"ExDump_{self.target.split('@')[0]}_{timestamp}"
        os.makedirs(dir_name, exist_ok=True)
        return dir_name
    
    def dump_emails(self, dump_dir):
        account = Account(
            primary_smtp_address=self.target,
            config=self.config,
            autodiscover=False,
            access_type=self.access_type
        )
        
        available_folders = []
        
        if hasattr(account, 'inbox'): available_folders.append((account.inbox, "Inbox"))
        if hasattr(account, 'sent'): available_folders.append((account.sent, "Sent"))
        if hasattr(account, 'drafts'): available_folders.append((account.drafts, "Drafts"))
        if hasattr(account, 'junk'): available_folders.append((account.junk, "Junk"))
        if hasattr(account, 'trash'): available_folders.append((account.trash, "Trash"))
        if hasattr(account, 'archive'): available_folders.append((account.archive, "Archive"))

        for folder, folder_name in available_folders:
            try:
                folder_path = os.path.join(dump_dir, folder_name)
                os.makedirs(folder_path, exist_ok=True)
                
               
                try:
                    items = folder.all().order_by('-datetime_received')[:500]
                except:
                    
                    items = folder.all()[:500]

                for item in items:
                    try:
                        subject = item.subject if item.subject else "No_Subject"
                        sender = item.sender if item.sender else "No_Sender"
                        to_recipients = item.to_recipients if item.to_recipients else "No_Recipients"
                        
                        body_content = "No Body Content"
                        try:
                            if item.text_body:
                                body_content = item.text_body
                            elif item.body:
                                body_content = item.body
                        except:
                            pass

                        date_obj = item.datetime_received
                        if not date_obj:
                            date_obj = item.datetime_created
                        
                        if not date_obj:
                            date_str = "Unknown_Date"
                        else:
                            date_str = date_obj.strftime('%Y%m%d_%H%M%S')

                        email_content = f"Subject: {subject}\n"
                        email_content += f"From: {sender}\n"
                        email_content += f"To: {to_recipients}\n"
                        email_content += f"Date: {date_str}\n"
                        email_content += "-" * 30 + "\n"
                        email_content += f"Body:\n{body_content}\n"
                        
                        safe_subject = "".join([c for c in subject if c.isalnum() or c in (' ', '_', '-')]).strip()
                        safe_subject = safe_subject[:30] # Limit length
                        filename = f"{date_str}_{safe_subject}_{item.id[:10]}.txt"
                        
                        filepath = os.path.join(folder_path, filename)
                        
                        with open(filepath, 'w', encoding='utf-8') as f:
                            f.write(email_content)
                            
                    except Exception as e:
                        continue
            except Exception as e:
                print(f"[-] Error processing folder {folder_name}: {e}")
                continue
        
        available_folders = []
        
        if hasattr(account, 'inbox'): available_folders.append((account.inbox, "Inbox"))
        if hasattr(account, 'sent'): available_folders.append((account.sent, "Sent"))
        if hasattr(account, 'drafts'): available_folders.append((account.drafts, "Drafts"))
        if hasattr(account, 'junk'): available_folders.append((account.junk, "Junk"))
        if hasattr(account, 'trash'): available_folders.append((account.trash, "Trash"))
        if hasattr(account, 'archive'): available_folders.append((account.archive, "Archive"))

        for folder, folder_name in available_folders:
            try:
                folder_path = os.path.join(dump_dir, folder_name)
                os.makedirs(folder_path, exist_ok=True)
                
                for item in folder.all().order_by('-datetime_received')[:500]:
                    try:
                        email_content = f"Subject: {item.subject}\n"
                        email_content += f"From: {item.sender}\n"
                        email_content += f"To: {item.to_recipients}\n"
                        email_content += f"Received: {item.datetime_received}\n"
                        email_content += f"Body: {item.text_body or item.body}\n"
                        
                        filename = f"{item.datetime_received.strftime('%Y%m%d_%H%M%S')}_{item.id[:20]}.txt"
                        filepath = os.path.join(folder_path, filename)
                        
                        with open(filepath, 'w', encoding='utf-8') as f:
                            f.write(email_content)
                    except Exception as e:
                        continue
            except Exception as e:
                continue
    
    def dump_calendar(self, dump_dir):
        account = Account(
            primary_smtp_address=self.target,
            config=self.config,
            autodiscover=False,
            access_type=self.access_type
        )
        
        calendar_path = os.path.join(dump_dir, "Calendar")
        os.makedirs(calendar_path, exist_ok=True)
        
        try:
            for item in account.calendar.all().order_by('-start')[:500]:
                try:
                    event_content = f"Subject: {item.subject}\n"
                    event_content += f"Start: {item.start}\n"
                    event_content += f"End: {item.end}\n"
                    event_content += f"Location: {item.location}\n"
                    event_content += f"Body: {item.text_body or item.body}\n"
                    
                    filename = f"{item.start.strftime('%Y%m%d_%H%M%S')}_{item.id[:20]}.txt"
                    filepath = os.path.join(calendar_path, filename)
                    
                    with open(filepath, 'w', encoding='utf-8') as f:
                        f.write(event_content)
                except Exception as e:
                    continue
        except Exception as e:
            pass
    
    def dump_contacts(self, dump_dir):
        account = Account(
            primary_smtp_address=self.target,
            config=self.config,
            autodiscover=False,
            access_type=self.access_type
        )
        
        contacts_path = os.path.join(dump_dir, "Contacts")
        os.makedirs(contacts_path, exist_ok=True)
        
        try:
            for item in account.contacts.all()[:1000]:
                try:
                    contact_content = f"Display Name: {item.display_name}\n"
                    contact_content += f"Email: {item.email_addresses}\n"
                    contact_content += f"Company: {item.company_name}\n"
                    contact_content += f"Phone: {item.phone_numbers}\n"
                    
                    filename = f"{item.display_name.replace('/', '_')[:50]}_{item.id[:10]}.txt"
                    filepath = os.path.join(contacts_path, filename)
                    
                    with open(filepath, 'w', encoding='utf-8') as f:
                        f.write(contact_content)
                except Exception as e:
                    continue
        except Exception as e:
            pass
    
    def dump_tasks(self, dump_dir):
        account = Account(
            primary_smtp_address=self.target,
            config=self.config,
            autodiscover=False,
            access_type=self.access_type
        )
        
        tasks_path = os.path.join(dump_dir, "Tasks")
        os.makedirs(tasks_path, exist_ok=True)
        
        try:
            for item in account.tasks.all().order_by('-datetime_created')[:500]:
                try:
                    task_content = f"Subject: {item.subject}\n"
                    task_content += f"Status: {item.status}\n"
                    task_content += f"Priority: {item.importance}\n"
                    task_content += f"Due: {item.due_date}\n"
                    task_content += f"Body: {item.text_body or item.body}\n"
                    
                    filename = f"{item.datetime_created.strftime('%Y%m%d_%H%M%S')}_{item.id[:20]}.txt"
                    filepath = os.path.join(tasks_path, filename)
                    
                    with open(filepath, 'w', encoding='utf-8') as f:
                        f.write(task_content)
                except Exception as e:
                    continue
        except Exception as e:
            pass

    def dump_attachments(self, dump_dir):
        account = Account(
            primary_smtp_address=self.target,
            config=self.config,
            autodiscover=False,
            access_type=self.access_type
        )
        
        attachments_path = os.path.join(dump_dir, "Attachments")
        os.makedirs(attachments_path, exist_ok=True)
        
        try:
            for item in account.inbox.all().order_by('-datetime_received')[:200]:
                try:
                    if item.attachments:
                        for attachment in item.attachments:
                            try:
                                if hasattr(attachment, 'content'):
                                    filename = f"{item.datetime_received.strftime('%Y%m%d_%H%M%S')}_{attachment.name}"
                                    filepath = os.path.join(attachments_path, filename)
                                    with open(filepath, 'wb') as f:
                                        f.write(attachment.content)
                            except Exception as e:
                                continue
                except Exception as e:
                    continue
        except Exception as e:
            pass
    
    def execute_dump(self):
        access_type = self.detect_access_type()
        print(f"Access type detected: {access_type}")
        
        dump_dir = self.create_dump_directory()
        print(f"Dump directory created: {dump_dir}")
        
        print("Dumping emails...")
        self.dump_emails(dump_dir)
        
        print("Dumping calendar items...")
        self.dump_calendar(dump_dir)
        
        print("Dumping contacts...")
        self.dump_contacts(dump_dir)
        
        print("Dumping tasks...")
        self.dump_tasks(dump_dir)
        
        print("Dumping attachments...")
        self.dump_attachments(dump_dir)
        
        print(f"Dump completed successfully in {dump_dir}")

def main():
    parser = argparse.ArgumentParser(description='Exchange Server Data Dumper')
    parser.add_argument('--domain', required=True, help='Domain name')
    parser.add_argument('--user', required=True, help='Username')
    parser.add_argument('--password', required=True, help='Password')
    parser.add_argument('--server', required=True, help='Exchange server IP/hostname')
    parser.add_argument('--target', required=True, help='Target mailbox email address')
    
    args = parser.parse_args()
    
    try:
        dumper = ExDumper(args.domain, args.user, args.password, args.server, args.target)
        dumper.execute_dump()
    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
