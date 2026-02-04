# 📧 Email Topic Organizer

A standalone Windows application that automatically sorts and organizes your emails by topic using machine learning.

## ✨ Features

✅ **Drag & drop emails** into the window (supports .eml and .msg files)  
✅ **Automatic topic classification** using HashingVectorizer + MiniBatchKMeans clustering  
✅ **Extracts Issue and Milestone markers** via regex pattern matching  
✅ **Append-only writes** - never rewrites old rows; preserves manual Excel edits  
✅ **Summary sheet** with counts + latest date per topic  
✅ **TopicMap sheet** where you can rename topics without breaking anything  
✅ **Index sheet** tracks processed emails by Message-ID or content hash for idempotency  
✅ **100% Local & Private** - no cloud calls, all processing happens on your computer

## 🚀 Getting Your .exe File

### Option 1: Download Pre-built Executable (Easiest)

1. Go to the [Releases page](../../releases)
2. Download `EmailOrganizer.exe` from the latest release
3. Run the .exe file (no installation needed!)

### Option 2: Build Using GitHub Actions (Recommended for beginners)

Since you don't have Python installed, you can use GitHub to build the .exe for you:

#### Step-by-Step Instructions:

1. **Create a GitHub account** (if you don't have one)
   - Go to https://github.com/signup
   - Follow the signup process

2. **Create a new repository**
   - Click the `+` button in the top right
   - Select "New repository"
   - Name it `email-organizer`
   - Make it Public
   - Click "Create repository"

3. **Upload all the files**
   - Click "uploading an existing file"
   - Drag and drop ALL these files into the upload area:
     - `email_sorter.py`
     - `requirements.txt`
     - `.github/workflows/build.yml` (create the `.github/workflows/` folders first)
   - Click "Commit changes"

4. **Wait for the build**
   - Go to the "Actions" tab in your repository
   - You'll see a workflow running called "Build Email Organizer"
   - Wait 3-5 minutes for it to complete (green checkmark)

5. **Download your .exe**
   - Click on the completed workflow
   - Scroll down to "Artifacts"
   - Download `EmailOrganizer-Windows.zip`
   - Extract the .zip file
   - You now have `EmailOrganizer.exe`!

## 📖 How to Use

1. **Run the application**
   - Double-click `EmailOrganizer.exe`
   - Windows might show a security warning (click "More info" → "Run anyway")

2. **Add email files**
   - Drag and drop .eml or .msg files into the blue drop zone
   - OR click "📂 Browse Files" to select files manually

3. **Process emails**
   - Click "⚡ Process & Export"
   - Choose where to save your Excel file
   - The app will:
     - Parse all emails
     - Classify them into topics automatically
     - Export to Excel with 4 sheets

4. **View your organized emails**
   - Open the Excel file
   - **Emails sheet**: All your emails with topics
   - **Summary sheet**: Topic statistics and counts
   - **TopicMap sheet**: Rename topics here (e.g., "Topic_1" → "Project Updates")
   - **Index sheet**: Tracks which emails have been processed

## 💡 Tips

- **Idempotent processing**: You can run the same emails multiple times - duplicates are automatically skipped
- **Append-only**: New emails are always added to the bottom - your manual edits are safe
- **Topic renaming**: Edit the "Custom Name" column in the TopicMap sheet to give topics meaningful names
- **Issue/Milestone markers**: The app automatically detects patterns like "Issue #123" or "Milestone: Phase 2"

## 🔧 Excel Sheet Details

### Emails Sheet
Contains all processed emails with columns:
- Date
- Topic (e.g., Topic_1, Topic_2)
- Subject
- From
- Issue (extracted markers)
- Milestone (extracted markers)
- Body Preview (first 200 characters)

### Summary Sheet
Auto-generated statistics:
- Topic name
- Email count per topic
- Latest email date per topic
- Mapped custom topic name (from TopicMap)

### TopicMap Sheet
Map generic topic names to meaningful ones:
- Original Topic: Topic_1, Topic_2, etc.
- Custom Name: Your custom labels (e.g., "Customer Support", "Project Alpha")

### Index Sheet
Tracks processed emails to prevent duplicates:
- Message ID (unique identifier)
- Processed Date
- Original filename

## 🔒 Privacy

This application is 100% local and private:
- All processing happens on your computer
- No internet connection required
- No data sent to any server
- Your emails stay on your machine

## 🛠️ Technical Details

- **Language**: Python 3.10
- **GUI**: Tkinter with drag & drop support
- **ML**: scikit-learn (HashingVectorizer + MiniBatchKMeans)
- **Excel**: openpyxl for reading/writing .xlsx files
- **Email parsing**: Built-in email library for .eml, basic .msg support

## ❓ Troubleshooting

**"Windows protected your PC" warning**
- Click "More info" → "Run anyway"
- This is normal for unsigned executables

**Drag & drop not working**
- Use the "📂 Browse Files" button instead
- Make sure you're dropping .eml or .msg files

**"All emails already processed"**
- These emails are already in your Excel file
- Check the Index sheet to see processed emails

**Excel file is locked**
- Close the Excel file before running the app
- Excel locks files when they're open

## 📝 License

Free to use for personal and commercial purposes.

## 🤝 Support

If you encounter issues:
1. Check the Troubleshooting section above
2. Open an issue on GitHub
3. Include the error message (if any)

---

**Made for organizing emails efficiently! 📧✨**
