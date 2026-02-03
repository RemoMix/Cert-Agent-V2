
import win32com.client
import os
import time
import sqlite3
from datetime import datetime, timezone
import shutil
import yaml
import logging

logger = logging.getLogger('CertPrintAgent')

class OutlookAgent:
    def __init__(self, config_path="config.yaml"):
        self.config = self.load_config(config_path)
        self.setup_paths()
        self.setup_database()
        self.outlook = None
    
    def load_config(self, config_path):
        """Load configuration file"""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            logger.error(f"Error loading config: {e}")
            return {}
    
    def setup_paths(self):
        """Create required directories"""
        base_dir = self.config.get('paths', {}).get('base_dir', 'Cert-Print-Agent')
        
        self.emails_dir = os.path.join(base_dir, self.config.get('paths', {}).get('emails_dir', 'GetCertAgent/MyEmails'))
        self.cert_inbox = os.path.join(base_dir, self.config.get('paths', {}).get('cert_inbox', 'GetCertAgent/Cert_Inbox'))
        
        os.makedirs(self.emails_dir, exist_ok=True)
        os.makedirs(self.cert_inbox, exist_ok=True)
        
        logger.info(f"Email dir: {self.emails_dir}")
        logger.info(f"Cert inbox: {self.cert_inbox}")
    
    def setup_database(self):
        """Setup SQLite database for tracking"""
        try:
            base_dir = self.config.get('paths', {}).get('base_dir', 'Cert-Print-Agent')
            db_path = self.config.get('monitoring', {}).get('processed_emails_db', 'processed_emails.db')
            db_full_path = os.path.join(base_dir, db_path)
            
            os.makedirs(os.path.dirname(db_full_path) if os.path.dirname(db_full_path) else base_dir, exist_ok=True)
            
            self.conn = sqlite3.connect(db_full_path, check_same_thread=False)
            self.cursor = self.conn.cursor()
            
            # Create table
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS processed_emails (
                    entry_id TEXT PRIMARY KEY,
                    subject TEXT,
                    sender TEXT,
                    received_time TEXT,
                    processed_time TEXT
                )
            """)
            self.conn.commit()
            logger.info("Database initialized")
            
        except Exception as e:
            logger.error(f"Database setup error: {e}")
            self.conn = None
            self.cursor = None
    
    def start_outlook(self):
        """Start/connect to Outlook"""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            logger.info("Connected to Outlook")
            return True
        except Exception as e:
            logger.error(f"Failed to connect to Outlook: {e}")
            
            if self.config.get('monitoring', {}).get('outlook_autostart', True):
                try:
                    os.startfile("outlook")
                    logger.info("Auto-starting Outlook...")
                    time.sleep(10)
                    self.outlook = win32com.client.Dispatch("Outlook.Application")
                    return True
                except Exception as e2:
                    logger.error(f"Auto-start failed: {e2}")
            
            return False
    
    def is_email_processed(self, entry_id):
        """Check if email was already processed"""
        if not self.cursor:
            return False
        try:
            self.cursor.execute(
                "SELECT COUNT(*) FROM processed_emails WHERE entry_id = ?",
                (entry_id,)
            )
            return self.cursor.fetchone()[0] > 0
        except Exception as e:
            logger.error(f"Database query error: {e}")
            return False
    
    def mark_email_processed(self, email):
        """Mark email as processed"""
        if not self.cursor:
            return
        try:
            received_time = email.ReceivedTime
            if hasattr(received_time, 'strftime'):
                received_str = received_time.strftime('%Y-%m-%d %H:%M:%S')
            else:
                received_str = str(received_time)
            
            self.cursor.execute("""
                INSERT OR REPLACE INTO processed_emails 
                (entry_id, subject, sender, received_time, processed_time)
                VALUES (?, ?, ?, ?, ?)
            """, (
                str(email.EntryID),
                str(email.Subject),
                str(getattr(email, 'SenderEmailAddress', 'Unknown')),
                received_str,
                datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            ))
            self.conn.commit()
            logger.info(f"Marked as processed: {email.Subject}")
        except Exception as e:
            logger.error(f"Database insert error: {e}")
    
    def is_certificate_file(self, filename):
        """Check if file is a certificate"""
        cert_exts = ['.pdf', '.jpg', '.jpeg', '.png', '.tiff', '.bmp', '.gif']
        return any(filename.lower().endswith(ext) for ext in cert_exts)
    
    def save_email_and_attachments(self, email):
        """Save email and its attachments"""
        saved_certs = []
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            subject_safe = "".join(c for c in str(email.Subject)[:30] if c.isalnum() or c in '-_')
            email_filename = f"{timestamp}_{subject_safe}.msg"
            email_path = os.path.join(self.emails_dir, email_filename)
            
            try:
                email.SaveAs(email_path, 3)
                logger.info(f"Email saved: {email_filename}")
            except Exception as e:
                logger.warning(f"Could not save email file: {e}")
            
            attachments = getattr(email, 'Attachments', None)
            if attachments:
                logger.info(f"Processing {attachments.Count} attachments")
                
                for attachment in attachments:
                    try:
                        filename = str(attachment.FileName)
                        
                        if self.is_certificate_file(filename):
                            att_path = os.path.join(self.emails_dir, filename)
                            attachment.SaveAsFile(att_path)
                            
                            cert_path = os.path.join(self.cert_inbox, filename)
                            shutil.copy2(att_path, cert_path)
                            saved_certs.append(cert_path)
                            
                            logger.info(f"Certificate saved: {filename}")
                    except Exception as e:
                        logger.error(f"Error saving attachment: {e}")
            
            return saved_certs
            
        except Exception as e:
            logger.error(f"Error processing email: {e}")
            return []
    
    def monitor_inbox(self, folder_name="Inbox"):
        """Monitor Outlook inbox for new certificates"""
        new_certs = []
        
        try:
            if not self.outlook:
                if not self.start_outlook():
                    return []
            
            namespace = self.outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)
            
            logger.info(f"Monitoring: {inbox.Name}")
            
            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)
            
            processed_count = 0
            
            for message in messages:
                try:
                    received_time = message.ReceivedTime
                    
                    try:
                        if hasattr(received_time, 'replace'):
                            if received_time.tzinfo is None:
                                received_time = received_time.replace(tzinfo=timezone.utc)
                            else:
                                received_time = received_time.astimezone(timezone.utc)
                            
                            current_time = datetime.now(timezone.utc)
                            time_diff = current_time - received_time
                            
                            if time_diff.days > 1:
                                break
                    except:
                        pass
                    
                    if not self.is_email_processed(message.EntryID):
                        subject = str(getattr(message, 'Subject', 'No Subject'))
                        logger.info(f"New email: {subject}")
                        
                        certs = self.save_email_and_attachments(message)
                        new_certs.extend(certs)
                        
                        self.mark_email_processed(message)
                        processed_count += 1
                
                except Exception as e:
                    logger.error(f"Error processing message: {e}")
                    continue
            
            logger.info(f"Processed {processed_count} new emails, found {len(new_certs)} certificates")
            return new_certs
            
        except Exception as e:
            logger.error(f"Monitor error: {e}")
            return []
    
    def run(self):
        """Run the agent"""
        logger.info("Starting OutlookAgent...")
        return self.monitor_inbox()
    
    def __del__(self):
        """Cleanup"""
        if hasattr(self, 'conn') and self.conn:
            try:
                self.conn.close()
            except:
                pass

def check_outlook_inbox(config_path="config.yaml"):
    agent = OutlookAgent(config_path)
    return agent.run()
