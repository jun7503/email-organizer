"""
Email Topic Organizer - Standalone Application
Drag & drop emails, auto-classify by topic, export to Excel
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import re
import hashlib
from datetime import datetime
from pathlib import Path
import email
from email import policy
from email.parser import BytesParser
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from sklearn.feature_extraction.text import HashingVectorizer
from sklearn.cluster import MiniBatchKMeans
import warnings
warnings.filterwarnings('ignore')


class EmailOrganizer:
    def __init__(self):
        self.emails = []
        self.excel_path = None
        self.vectorizer = HashingVectorizer(n_features=100, stop_words='english')
        self.kmeans = MiniBatchKMeans(n_clusters=5, random_state=42, n_init=3)
        
    def parse_email_file(self, filepath):
        """Parse .eml or .msg file"""
        try:
            if filepath.lower().endswith('.eml'):
                return self._parse_eml(filepath)
            elif filepath.lower().endswith('.msg'):
                return self._parse_msg(filepath)
            else:
                print(f"Unsupported file type: {filepath}")
                return None
        except Exception as e:
            print(f"Error parsing {filepath}: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def _parse_eml(self, filepath):
        """Parse .eml file"""
        with open(filepath, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)
        
        # Extract email data
        subject = msg.get('Subject', 'No Subject')
        from_addr = msg.get('From', 'Unknown')
        date_str = msg.get('Date', '')
        
        # Parse date
        try:
            if date_str:
                date_obj = email.utils.parsedate_to_datetime(date_str)
            else:
                date_obj = datetime.now()
        except:
            date_obj = datetime.now()
        
        # Extract body
        body = self._extract_body_eml(msg)
        
        # Create unique ID
        content_for_hash = f"{subject}{from_addr}{body[:500]}"
        msg_id = msg.get('Message-ID', hashlib.md5(content_for_hash.encode()).hexdigest())
        
        # Extract issue/milestone markers
        issue = self._extract_marker(subject + ' ' + body, r'(?i)(issue|bug|problem)\s*[:#]?\s*(\w+)')
        milestone = self._extract_marker(subject + ' ' + body, r'(?i)(milestone|phase|sprint)\s*[:#]?\s*(\w+)')
        
        return {
            'message_id': msg_id,
            'subject': subject,
            'from': from_addr,
            'date': date_obj,
            'body': body,
            'issue': issue,
            'milestone': milestone,
            'filepath': filepath
        }
    
    def _parse_msg(self, filepath):
        """Parse .msg file using extract-msg library"""
        try:
            import extract_msg
            
            msg = extract_msg.Message(filepath)
            
            # Extract email data
            subject = msg.subject or 'No Subject'
            from_addr = msg.sender or 'Unknown'
            
            # Get date
            try:
                date_obj = msg.date
                if date_obj is None:
                    date_obj = datetime.now()
            except:
                date_obj = datetime.now()
            
            # Extract body
            body = msg.body or ''
            if not body:
                body = msg.htmlBody or ''
            
            # Limit body length
            body = body[:1000]
            
            # Create unique ID
            content_for_hash = f"{subject}{from_addr}{body[:500]}"
            msg_id = msg.messageId or hashlib.md5(content_for_hash.encode()).hexdigest()
            
            # Extract issue/milestone markers
            issue = self._extract_marker(subject + ' ' + body, r'(?i)(issue|bug|problem)\s*[:#]?\s*(\w+)')
            milestone = self._extract_marker(subject + ' ' + body, r'(?i)(milestone|phase|sprint)\s*[:#]?\s*(\w+)')
            
            # Close the message
            msg.close()
            
            return {
                'message_id': msg_id,
                'subject': subject,
                'from': from_addr,
                'date': date_obj,
                'body': body,
                'issue': issue,
                'milestone': milestone,
                'filepath': filepath
            }
            
        except ImportError:
            # extract-msg not available, show error
            raise Exception("extract-msg library is required for .msg files. Please reinstall the application.")
        except Exception as e:
            print(f"Error parsing MSG file {filepath}: {e}")
            raise
    
    def _extract_body_eml(self, msg):
        """Extract email body text from EML"""
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    try:
                        body += part.get_content()
                    except:
                        pass
        else:
            try:
                body = msg.get_content()
            except:
                body = str(msg)
        
        return body[:1000]  # Limit to first 1000 characters
    
    def _extract_marker(self, text, pattern):
        """Extract issue or milestone markers"""
        match = re.search(pattern, text)
        if match:
            return match.group(0)
        return ''
    
    def classify_topics(self):
        """Classify emails into topics using clustering"""
        if len(self.emails) < 2:
            # Not enough emails to cluster
            for email_data in self.emails:
                email_data['topic'] = 'Topic_1'
            return
        
        # Prepare text for clustering
        texts = [f"{e['subject']} {e['body']}" for e in self.emails]
        
        # Vectorize
        X = self.vectorizer.fit_transform(texts)
        
        # Adjust number of clusters based on email count
        n_clusters = min(5, max(2, len(self.emails) // 3))
        self.kmeans.set_params(n_clusters=n_clusters)
        
        # Cluster
        labels = self.kmeans.fit_predict(X)
        
        # Assign topics
        for i, email_data in enumerate(self.emails):
            email_data['topic'] = f'Topic_{labels[i] + 1}'
    
    def load_existing_excel(self, filepath):
        """Load existing Excel file to check for duplicates"""
        processed_ids = set()
        
        if not os.path.exists(filepath):
            return processed_ids
        
        try:
            wb = openpyxl.load_workbook(filepath)
            if 'Index' in wb.sheetnames:
                ws = wb['Index']
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row[0]:  # Message ID column
                        processed_ids.add(row[0])
            wb.close()
        except Exception as e:
            print(f"Error loading existing Excel: {e}")
        
        return processed_ids
    
    def save_to_excel(self, filepath):
        """Save emails to Excel with all required sheets"""
        # Load existing file or create new
        if os.path.exists(filepath):
            wb = openpyxl.load_workbook(filepath)
        else:
            wb = openpyxl.Workbook()
            wb.remove(wb.active)  # Remove default sheet
        
        # Ensure all required sheets exist
        if 'Emails' not in wb.sheetnames:
            wb.create_sheet('Emails', 0)
        if 'Summary' not in wb.sheetnames:
            wb.create_sheet('Summary', 1)
        if 'TopicMap' not in wb.sheetnames:
            wb.create_sheet('TopicMap', 2)
        if 'Index' not in wb.sheetnames:
            wb.create_sheet('Index', 3)
        
        # Write to Emails sheet (append-only)
        self._write_emails_sheet(wb['Emails'])
        
        # Write to Summary sheet (regenerate)
        self._write_summary_sheet(wb)
        
        # Write to TopicMap sheet (initialize if empty)
        self._write_topicmap_sheet(wb['TopicMap'])
        
        # Write to Index sheet (append-only)
        self._write_index_sheet(wb['Index'])
        
        # Save
        wb.save(filepath)
        wb.close()
    
    def _write_emails_sheet(self, ws):
        """Append new emails to Emails sheet"""
        # Check if headers exist
        if ws.max_row == 1 or ws.cell(1, 1).value != 'Date':
            ws.append(['Date', 'Topic', 'Subject', 'From', 'Issue', 'Milestone', 'Body Preview'])
            self._style_header(ws)
        
        # Append new emails
        for email_data in self.emails:
            ws.append([
                email_data['date'].strftime('%Y-%m-%d %H:%M:%S'),
                email_data['topic'],
                email_data['subject'],
                email_data['from'],
                email_data['issue'],
                email_data['milestone'],
                email_data['body'][:200]
            ])
    
    def _write_summary_sheet(self, wb):
        """Regenerate Summary sheet with topic statistics"""
        ws = wb['Summary']
        ws.delete_rows(1, ws.max_row)
        
        # Headers
        ws.append(['Topic', 'Count', 'Latest Date', 'Mapped Topic Name'])
        self._style_header(ws)
        
        # Get all emails from Emails sheet
        emails_ws = wb['Emails']
        topic_data = {}
        
        for row in emails_ws.iter_rows(min_row=2, values_only=True):
            if row[1]:  # Topic column
                topic = row[1]
                date_str = row[0]
                
                if topic not in topic_data:
                    topic_data[topic] = {'count': 0, 'latest': None}
                
                topic_data[topic]['count'] += 1
                
                # Update latest date
                try:
                    date_obj = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
                    if topic_data[topic]['latest'] is None or date_obj > topic_data[topic]['latest']:
                        topic_data[topic]['latest'] = date_obj
                except:
                    pass
        
        # Get topic mappings
        topicmap = self._get_topic_mappings(wb)
        
        # Write summary rows
        for topic in sorted(topic_data.keys()):
            latest = topic_data[topic]['latest']
            latest_str = latest.strftime('%Y-%m-%d %H:%M:%S') if latest else 'N/A'
            mapped_name = topicmap.get(topic, '')
            
            ws.append([
                topic,
                topic_data[topic]['count'],
                latest_str,
                mapped_name
            ])
    
    def _write_topicmap_sheet(self, ws):
        """Initialize TopicMap sheet if empty"""
        if ws.max_row <= 1 or ws.cell(1, 1).value != 'Original Topic':
            ws.delete_rows(1, ws.max_row)
            ws.append(['Original Topic', 'Custom Name'])
            self._style_header(ws)
            
            # Add all unique topics
            unique_topics = set(e['topic'] for e in self.emails)
            for topic in sorted(unique_topics):
                ws.append([topic, ''])
    
    def _write_index_sheet(self, ws):
        """Append processed email IDs to Index sheet"""
        if ws.max_row == 1 or ws.cell(1, 1).value != 'Message ID':
            ws.append(['Message ID', 'Processed Date', 'Filename'])
            self._style_header(ws)
        
        # Append new email IDs
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        for email_data in self.emails:
            ws.append([
                email_data['message_id'],
                now,
                Path(email_data['filepath']).name
            ])
    
    def _get_topic_mappings(self, wb):
        """Get topic name mappings from TopicMap sheet"""
        mappings = {}
        if 'TopicMap' in wb.sheetnames:
            ws = wb['TopicMap']
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] and row[1]:
                    mappings[row[0]] = row[1]
        return mappings
    
    def _style_header(self, ws):
        """Apply styling to header row"""
        header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF')
        
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')


class EmailOrganizerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Topic Organizer")
        self.root.geometry("700x550")
        
        self.organizer = EmailOrganizer()
        self.dropped_files = []
        self.drag_drop_available = False
        
        self.setup_ui()
        self.setup_drag_drop()
    
    def setup_ui(self):
        """Create the user interface"""
        # Title
        title = tk.Label(self.root, text="📧 Email Topic Organizer", 
                        font=('Arial', 16, 'bold'))
        title.pack(pady=10)
        
        # Instructions
        instructions = tk.Label(self.root, 
            text="Use Browse button to select .eml or .msg files\nEmails will be automatically sorted by topic",
            justify=tk.CENTER)
        instructions.pack(pady=5)
        
        # Drop zone
        drop_frame = tk.Frame(self.root, bg='#e8f4f8', relief=tk.SOLID, borderwidth=2)
        drop_frame.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)
        
        drop_text = "📁 Select files using Browse button\n\nSupported: .eml, .msg"
        if self.drag_drop_available:
            drop_text = "📁 Drop email files here or use Browse button\n\nSupported: .eml, .msg"
        
        self.drop_label = tk.Label(drop_frame, 
            text=drop_text,
            bg='#e8f4f8', font=('Arial', 12), justify=tk.CENTER)
        self.drop_label.pack(expand=True)
        
        # File list
        list_frame = tk.Frame(self.root)
        list_frame.pack(padx=20, pady=5, fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.file_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, height=8)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)
        
        # Buttons
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)
        
        self.browse_btn = tk.Button(btn_frame, text="📂 Browse Files", 
                                    command=self.browse_files, width=15)
        self.browse_btn.pack(side=tk.LEFT, padx=5)
        
        self.clear_btn = tk.Button(btn_frame, text="🗑️ Clear List", 
                                   command=self.clear_files, width=15)
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        
        self.process_btn = tk.Button(btn_frame, text="⚡ Process & Export", 
                                     command=self.process_emails, 
                                     width=15, bg='#4CAF50', fg='white',
                                     font=('Arial', 10, 'bold'))
        self.process_btn.pack(side=tk.LEFT, padx=5)
        
        # Status
        self.status_label = tk.Label(self.root, text="Ready - Use Browse button to select files", fg='green')
        self.status_label.pack(pady=5)
    
    def setup_drag_drop(self):
        """Enable drag and drop functionality if available"""
        try:
            from tkinterdnd2 import TkinterDnD
            # Try to enable drag and drop
            self.root.drop_target_register('DND_Files')
            self.root.dnd_bind('<<Drop>>', self.on_drop)
            self.drag_drop_available = True
            self.drop_label.config(text="📁 Drop email files here or use Browse button\n\nSupported: .eml, .msg")
            self.status_label.config(text="Ready - Drag & drop enabled")
        except Exception as e:
            # Drag and drop not available, that's ok
            print(f"Drag and drop not available: {e}")
            self.drag_drop_available = False
    
    def on_drop(self, event):
        """Handle dropped files"""
        try:
            files = self.root.tk.splitlist(event.data)
            self.add_files(files)
        except Exception as e:
            print(f"Error handling drop: {e}")
    
    def browse_files(self):
        """Browse for email files"""
        files = filedialog.askopenfilenames(
            title="Select Email Files",
            filetypes=[("Email files", "*.eml *.msg"), ("All files", "*.*")]
        )
        if files:
            self.add_files(files)
    
    def add_files(self, files):
        """Add files to the list"""
        for filepath in files:
            if filepath.lower().endswith(('.eml', '.msg')):
                if filepath not in self.dropped_files:
                    self.dropped_files.append(filepath)
                    self.file_listbox.insert(tk.END, Path(filepath).name)
        
        self.update_status(f"{len(self.dropped_files)} files ready")
    
    def clear_files(self):
        """Clear the file list"""
        self.dropped_files = []
        self.file_listbox.delete(0, tk.END)
        self.update_status("Ready")
    
    def process_emails(self):
        """Process all emails and export to Excel"""
        if not self.dropped_files:
            messagebox.showwarning("No Files", "Please add email files first")
            return
        
        # Ask for Excel output location
        excel_path = filedialog.asksaveasfilename(
            title="Save Excel File As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if not excel_path:
            return
        
        self.update_status("Processing emails...")
        self.root.update()
        
        try:
            # Load existing processed IDs
            processed_ids = self.organizer.load_existing_excel(excel_path)
            
            # Parse emails
            new_emails = 0
            failed_files = []
            for filepath in self.dropped_files:
                try:
                    email_data = self.organizer.parse_email_file(filepath)
                    if email_data and email_data['message_id'] not in processed_ids:
                        self.organizer.emails.append(email_data)
                        new_emails += 1
                except Exception as e:
                    failed_files.append((Path(filepath).name, str(e)))
                    print(f"Failed to parse {filepath}: {e}")
            
            # Show warning if some files failed
            if failed_files:
                failed_list = "\n".join([f"- {name}: {error[:50]}..." for name, error in failed_files[:5]])
                messagebox.showwarning("Some Files Failed", 
                    f"{len(failed_files)} file(s) could not be processed:\n\n{failed_list}\n\n" +
                    f"Successfully processed: {new_emails} files")
            
            if new_emails == 0:
                messagebox.showinfo("No New Emails", 
                    "All selected emails have already been processed or failed to parse")
                self.update_status("No new emails to process")
                return
            
            # Classify topics
            self.update_status("Classifying topics...")
            self.root.update()
            self.organizer.classify_topics()
            
            # Save to Excel
            self.update_status("Saving to Excel...")
            self.root.update()
            self.organizer.save_to_excel(excel_path)
            
            # Success
            messagebox.showinfo("Success", 
                f"Processed {new_emails} new emails!\n\nSaved to:\n{excel_path}")
            
            self.update_status(f"✅ Success! Processed {new_emails} emails")
            
            # Clear for next batch
            self.clear_files()
            self.organizer.emails = []
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Error processing emails:\n{error_details}")
            messagebox.showerror("Error", f"Failed to process emails:\n{str(e)}")
            self.update_status("❌ Error occurred")
    
    def update_status(self, message):
        """Update status label"""
        self.status_label.config(text=message)
        self.root.update()


def main():
    try:
        # Try to use TkinterDnD if available
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except:
        # Fall back to regular Tkinter
        root = tk.Tk()
    
    app = EmailOrganizerGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
