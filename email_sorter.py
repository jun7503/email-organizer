"""
Email Topic Organizer - Enhanced with Thread/Conversation Parsing
Handles email chains with multiple back-and-forth replies
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
from openpyxl.utils import get_column_letter
from sklearn.feature_extraction.text import HashingVectorizer
from sklearn.cluster import MiniBatchKMeans
import warnings
warnings.filterwarnings('ignore')


class EmailOrganizer:
    def __init__(self):
        self.emails = []
        self.excel_path = None
        self.vectorizer = HashingVectorizer(
            n_features=200,
            stop_words='english',
            ngram_range=(1, 2)
        )
        self.kmeans = MiniBatchKMeans(n_clusters=5, random_state=42, n_init=10, max_iter=300)
        
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
        
        # Parse conversation thread
        thread_info = self._parse_conversation_thread(body)
        
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
            'filepath': filepath,
            'thread_count': thread_info['count'],
            'thread_summary': thread_info['summary'],
            'conversation_dates': thread_info['dates']
        }
    
    def _parse_msg(self, filepath):
        """Parse .msg file using extract-msg library"""
        try:
            import extract_msg
            
            msg = extract_msg.Message(filepath)
            
            # Extract email data
            subject = msg.subject or 'No Subject'
            from_addr = msg.sender or 'Unknown'
            
            # Get date - handle both datetime and string formats
            try:
                date_obj = msg.date
                if date_obj is None:
                    date_obj = datetime.now()
                elif isinstance(date_obj, str):
                    try:
                        date_obj = datetime.strptime(date_obj, '%a, %d %b %Y %H:%M:%S %z')
                    except:
                        try:
                            date_obj = datetime.strptime(date_obj, '%Y-%m-%d %H:%M:%S')
                        except:
                            date_obj = datetime.now()
            except:
                date_obj = datetime.now()
            
            # Extract body
            body = msg.body or ''
            if not body:
                body = msg.htmlBody or ''
            
            # Parse conversation thread from body
            thread_info = self._parse_conversation_thread(body)
            
            # Create unique ID
            content_for_hash = f"{subject}{from_addr}{body[:500]}"
            msg_id = msg.messageId or hashlib.md5(content_for_hash.encode()).hexdigest()
            
            # Extract issue/milestone markers
            issue = self._extract_marker(subject + ' ' + body, r'(?i)(issue|bug|problem)\s*[:#]?\s*(\w+)')
            milestone = self._extract_marker(subject + ' ' + body, r'(?i)(milestone|phase|sprint)\s*[:#]?\s*(\w+)')
            
            msg.close()
            
            return {
                'message_id': msg_id,
                'subject': subject,
                'from': from_addr,
                'date': date_obj,
                'body': body,
                'issue': issue,
                'milestone': milestone,
                'filepath': filepath,
                'thread_count': thread_info['count'],
                'thread_summary': thread_info['summary'],
                'conversation_dates': thread_info['dates']
            }
            
        except ImportError:
            raise Exception("extract-msg library is required for .msg files. Please reinstall the application.")
        except Exception as e:
            print(f"Error parsing MSG file {filepath}: {e}")
            raise
    
    def _parse_conversation_thread(self, body_text):
        """
        Parse email body to detect conversation threads
        Returns: dict with thread count, summary, and dates
        """
        if not body_text:
            return {'count': 1, 'summary': '', 'dates': []}
        
        # Detect conversation indicators
        from_pattern = r'From:\s*([^\r\n]+)'
        sent_pattern = r'Sent:\s*([^\r\n]+)'
        date_pattern = r'Date:\s*([^\r\n]+)'
        
        # Find all "From:" occurrences (indicates forwarded/replied emails)
        from_matches = re.findall(from_pattern, body_text)
        sent_matches = re.findall(sent_pattern, body_text)
        date_matches = re.findall(date_pattern, body_text)
        
        # Count emails in thread (original + replies)
        thread_count = max(len(from_matches), len(sent_matches), 1)
        
        # Extract conversation dates
        conversation_dates = []
        for date_str in (sent_matches + date_matches):
            try:
                # Try to parse date string
                parsed_date = self._parse_flexible_date(date_str)
                if parsed_date:
                    conversation_dates.append(parsed_date)
            except:
                pass
        
        # Create summary of key points
        summary = self._extract_conversation_summary(body_text, thread_count)
        
        return {
            'count': thread_count,
            'summary': summary,
            'dates': sorted(set(conversation_dates))
        }
    
    def _parse_flexible_date(self, date_str):
        """Try to parse date from various formats"""
        date_formats = [
            '%m/%d/%Y %I:%M:%S %p',  # 12/3/2025 1:15 PM
            '%m/%d/%Y %I:%M %p',      # 12/3/2025 1:15 PM
            '%d/%m/%Y %H:%M:%S',      # 03/12/2025 13:15:00
            '%Y-%m-%d %H:%M:%S',      # 2025-12-03 13:15:00
            '%a, %d %b %Y %H:%M:%S',  # Tue, 3 Dec 2025 13:15:00
            '%B %d, %Y %I:%M %p',     # December 3, 2025 1:15 PM
        ]
        
        for fmt in date_formats:
            try:
                return datetime.strptime(date_str.strip(), fmt)
            except:
                continue
        return None
    
    def _extract_conversation_summary(self, body_text, thread_count):
        """
        Extract key points from conversation thread
        Returns a concise summary of the discussion
        """
        if thread_count <= 1:
            # Single email, take first meaningful paragraph
            lines = body_text.split('\n')
            meaningful_lines = [line.strip() for line in lines if len(line.strip()) > 20 and not line.strip().startswith('>')]
            return ' '.join(meaningful_lines[:3])[:300]
        
        # Multi-email thread - extract key points
        summary_parts = []
        
        # Split by common separators
        sections = re.split(r'_{5,}|From:|Sent:|Original Message', body_text, flags=re.IGNORECASE)
        
        for section in sections[:5]:  # Take first 5 sections
            lines = section.split('\n')
            # Find meaningful content (not headers, not quoted)
            content_lines = []
            for line in lines:
                clean_line = line.strip()
                # Skip headers, quotes, and short lines
                if (len(clean_line) > 30 and 
                    not clean_line.startswith('>') and
                    not clean_line.startswith('-----') and
                    not any(header in clean_line for header in ['Subject:', 'To:', 'Cc:', 'Date:', 'From:'])):
                    content_lines.append(clean_line)
                    if len(content_lines) >= 2:  # Take max 2 lines per section
                        break
            
            if content_lines:
                summary_parts.extend(content_lines)
        
        # Combine and limit length
        full_summary = ' | '.join(summary_parts)
        return full_summary[:500] if full_summary else 'Conversation thread'
    
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
        
        return body
    
    def _extract_marker(self, text, pattern):
        """Extract issue or milestone markers"""
        match = re.search(pattern, text)
        if match:
            return match.group(0)
        return ''
    
    def classify_topics(self):
        """Classify emails into topics using clustering"""
        if len(self.emails) < 2:
            for email_data in self.emails:
                email_data['topic'] = 'Topic_1'
            return
        
        # Use subject + body + conversation summary for better clustering
        texts = [f"{e['subject']} {e['subject']} {e['body'][:1000]} {e.get('thread_summary', '')}" for e in self.emails]
        
        X = self.vectorizer.fit_transform(texts)
        n_clusters = min(5, max(2, len(self.emails) // 5))
        self.kmeans.set_params(n_clusters=n_clusters)
        labels = self.kmeans.fit_predict(X)
        
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
                    if row[0]:
                        msg_id = str(row[0]).strip()
                        processed_ids.add(msg_id)
            wb.close()
        except Exception as e:
            print(f"Error loading existing Excel: {e}")
        
        return processed_ids
    
    def save_to_excel(self, filepath):
        """Save emails to Excel with all required sheets"""
        if os.path.exists(filepath):
            wb = openpyxl.load_workbook(filepath)
        else:
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
        
        if 'Emails' not in wb.sheetnames:
            wb.create_sheet('Emails', 0)
        if 'Summary' not in wb.sheetnames:
            wb.create_sheet('Summary', 1)
        if 'TopicMap' not in wb.sheetnames:
            wb.create_sheet('TopicMap', 2)
        if 'Index' not in wb.sheetnames:
            wb.create_sheet('Index', 3)
        
        self._write_emails_sheet(wb['Emails'])
        self._write_summary_sheet(wb)
        self._write_topicmap_sheet(wb)
        self._write_index_sheet(wb['Index'])
        
        wb.save(filepath)
        wb.close()
    
    def _write_emails_sheet(self, ws):
        """Append new emails to Emails sheet with conversation info"""
        # Check if headers exist
        if ws.max_row == 1 or ws.cell(1, 1).value != 'Date':
            headers = ['Date', 'Topic', 'Subject', 'From', 'Thread Count', 'Conversation Summary', 
                      'Issue', 'Milestone', 'Body Preview']
            ws.append(headers)
            self._style_header(ws)
        
        # Append new emails
        for email_data in self.emails:
            # Format conversation dates
            conv_dates = email_data.get('conversation_dates', [])
            date_range = ''
            if len(conv_dates) >= 2:
                date_range = f"{conv_dates[0].strftime('%m/%d/%Y')} - {conv_dates[-1].strftime('%m/%d/%Y')}"
            elif len(conv_dates) == 1:
                date_range = conv_dates[0].strftime('%m/%d/%Y')
            
            ws.append([
                email_data['date'].strftime('%Y-%m-%d %H:%M:%S'),
                email_data['topic'],
                email_data['subject'],
                email_data['from'],
                email_data.get('thread_count', 1),
                email_data.get('thread_summary', '')[:500],  # Limit to 500 chars
                email_data.get('issue', '') or '',
                email_data.get('milestone', '') or '',
                email_data['body'][:200] or ''
            ])
        
        self._auto_adjust_columns(ws)
    
    def _auto_adjust_columns(self, ws):
        """Auto-adjust column widths for readability"""
        for column_cells in ws.columns:
            length = 0
            column = column_cells[0].column_letter
            
            for cell in column_cells:
                try:
                    if cell.value:
                        cell_length = len(str(cell.value))
                        if cell_length > length:
                            length = cell_length
                except:
                    pass
            
            adjusted_width = min(max(length + 2, 10), 80)  # Increased max width for summary column
            ws.column_dimensions[column].width = adjusted_width
    
    def _write_summary_sheet(self, wb):
        """Regenerate Summary sheet with topic statistics"""
        ws = wb['Summary']
        ws.delete_rows(1, ws.max_row)
        
        ws.append(['Topic', 'Count', 'Latest Date', 'Total Threads', 'Mapped Topic Name'])
        self._style_header(ws)
        
        emails_ws = wb['Emails']
        topic_data = {}
        
        for row in emails_ws.iter_rows(min_row=2, values_only=True):
            if row[1]:  # Topic column
                topic = row[1]
                date_str = row[0]
                thread_count = row[4] if len(row) > 4 else 1
                
                if topic not in topic_data:
                    topic_data[topic] = {'count': 0, 'latest': None, 'threads': 0}
                
                topic_data[topic]['count'] += 1
                topic_data[topic]['threads'] += thread_count
                
                try:
                    date_obj = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
                    if topic_data[topic]['latest'] is None or date_obj > topic_data[topic]['latest']:
                        topic_data[topic]['latest'] = date_obj
                except:
                    pass
        
        topicmap = self._get_topic_mappings(wb)
        
        for topic in sorted(topic_data.keys()):
            latest = topic_data[topic]['latest']
            latest_str = latest.strftime('%Y-%m-%d %H:%M:%S') if latest else 'N/A'
            mapped_name = topicmap.get(topic, '')
            
            ws.append([
                topic,
                topic_data[topic]['count'],
                latest_str,
                topic_data[topic]['threads'],
                mapped_name
            ])
        
        self._auto_adjust_columns(ws)
    
    def _write_topicmap_sheet(self, wb):
        """Initialize or update TopicMap sheet with all topics"""
        emails_ws = wb['Emails']
        all_topics = set()
        
        for row in emails_ws.iter_rows(min_row=2, values_only=True):
            if row[1]:
                all_topics.add(row[1])
        
        ws = wb['TopicMap']
        existing_mappings = {}
        
        if ws.max_row > 1 and ws.cell(1, 1).value == 'Original Topic':
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0]:
                    existing_mappings[row[0]] = row[1] or ''
        
        ws.delete_rows(1, ws.max_row)
        ws.append(['Original Topic', 'Custom Name'])
        self._style_header(ws)
        
        for topic in sorted(all_topics):
            custom_name = existing_mappings.get(topic, '')
            ws.append([topic, custom_name])
        
        self._auto_adjust_columns(ws)
    
    def _write_index_sheet(self, ws):
        """Append processed email IDs to Index sheet"""
        if ws.max_row == 1 or ws.cell(1, 1).value != 'Message ID':
            ws.append(['Message ID', 'Processed Date', 'Filename'])
            self._style_header(ws)
        
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        for email_data in self.emails:
            msg_id = str(email_data['message_id']).strip()
            ws.append([
                msg_id,
                now,
                Path(email_data['filepath']).name
            ])
        
        self._auto_adjust_columns(ws)
    
    def _get_topic_mappings(self, wb):
        """Get topic name mappings from TopicMap sheet"""
        mappings = {}
        if 'TopicMap' in wb.sheetnames:
            ws = wb['TopicMap']
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0]:
                    mappings[row[0]] = row[1] or ''
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
        self.root.title("Email Topic Organizer - With Thread Detection")
        self.root.geometry("700x550")
        
        self.organizer = EmailOrganizer()
        self.dropped_files = []
        self.drag_drop_available = False
        
        self.setup_ui()
        self.setup_drag_drop()
    
    def setup_ui(self):
        """Create the user interface"""
        title = tk.Label(self.root, text="📧 Email Topic Organizer", 
                        font=('Arial', 16, 'bold'))
        title.pack(pady=10)
        
        instructions = tk.Label(self.root, 
            text="Use Browse button to select .eml or .msg files\nAutomatically detects email threads and conversation history",
            justify=tk.CENTER)
        instructions.pack(pady=5)
        
        drop_frame = tk.Frame(self.root, bg='#e8f4f8', relief=tk.SOLID, borderwidth=2)
        drop_frame.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)
        
        drop_text = "📁 Select files using Browse button\n\nSupported: .eml, .msg\nThreads: Auto-detected"
        if self.drag_drop_available:
            drop_text = "📁 Drop email files here or use Browse button\n\nSupported: .eml, .msg\nThreads: Auto-detected"
        
        self.drop_label = tk.Label(drop_frame, 
            text=drop_text,
            bg='#e8f4f8', font=('Arial', 12), justify=tk.CENTER)
        self.drop_label.pack(expand=True)
        
        list_frame = tk.Frame(self.root)
        list_frame.pack(padx=20, pady=5, fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.file_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, height=8)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)
        
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
        
        self.status_label = tk.Label(self.root, text="Ready - Use Browse button to select files", fg='green')
        self.status_label.pack(pady=5)
    
    def setup_drag_drop(self):
        """Enable drag and drop functionality if available"""
        try:
            from tkinterdnd2 import TkinterDnD
            self.root.drop_target_register('DND_Files')
            self.root.dnd_bind('<<Drop>>', self.on_drop)
            self.drag_drop_available = True
            self.drop_label.config(text="📁 Drop email files here or use Browse button\n\nSupported: .eml, .msg\nThreads: Auto-detected")
            self.status_label.config(text="Ready - Drag & drop enabled")
        except Exception as e:
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
        
        excel_path = filedialog.asksaveasfilename(
            title="Save Excel File As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if not excel_path:
            return
        
        self.update_status("Processing emails and detecting threads...")
        self.root.update()
        
        try:
            processed_ids = self.organizer.load_existing_excel(excel_path)
            
            new_emails = 0
            skipped = 0
            failed_files = []
            total_threads = 0
            
            for filepath in self.dropped_files:
                try:
                    email_data = self.organizer.parse_email_file(filepath)
                    if email_data:
                        msg_id = str(email_data['message_id']).strip()
                        
                        if msg_id not in processed_ids:
                            self.organizer.emails.append(email_data)
                            new_emails += 1
                            total_threads += email_data.get('thread_count', 1)
                        else:
                            skipped += 1
                except Exception as e:
                    failed_files.append((Path(filepath).name, str(e)))
                    print(f"Failed to parse {filepath}: {e}")
            
            info_msg = f"Processed: {new_emails} emails ({total_threads} messages in threads)"
            if skipped > 0:
                info_msg += f"\nSkipped: {skipped} duplicates"
            if failed_files:
                failed_list = "\n".join([f"- {name}" for name, error in failed_files[:5]])
                info_msg += f"\nFailed: {len(failed_files)} files\n\n{failed_list}"
                if len(failed_files) > 5:
                    info_msg += f"\n... and {len(failed_files) - 5} more"
            
            if new_emails == 0:
                messagebox.showinfo("No New Emails", 
                    f"All selected emails have already been processed or failed to parse.\n\n{info_msg}")
                self.update_status("No new emails to process")
                return
            
            self.update_status("Classifying topics...")
            self.root.update()
            self.organizer.classify_topics()
            
            self.update_status("Saving to Excel with conversation summaries...")
            self.root.update()
            self.organizer.save_to_excel(excel_path)
            
            messagebox.showinfo("Success", f"{info_msg}\n\nSaved to:\n{excel_path}")
            
            self.update_status(f"✅ Success! Processed {new_emails} emails with {total_threads} messages")
            
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
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except:
        root = tk.Tk()
    
    app = EmailOrganizerGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
