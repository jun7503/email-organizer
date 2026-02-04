# 🎯 BEGINNER'S GUIDE - Getting Your .exe File

This guide is for people with **ZERO programming knowledge**. Follow these steps exactly.

---

## 🌟 EASIEST METHOD: Use GitHub to Build Your .exe

Since you don't have Python installed, GitHub will build the .exe file for you automatically!

### Step 1: Create a GitHub Account

1. Go to **https://github.com/signup**
2. Enter your email address
3. Create a password
4. Choose a username
5. Verify your account (check your email)
6. Click "Create account"

✅ You now have a GitHub account!

---

### Step 2: Create a New Repository

1. Click the **`+`** button in the top-right corner of GitHub
2. Click **"New repository"**
3. Fill in:
   - **Repository name**: `email-organizer`
   - **Description** (optional): "Email sorting tool"
   - **Public** (must be selected)
   - ❌ Do NOT check "Add a README file"
4. Click **"Create repository"**

✅ You now have an empty repository!

---

### Step 3: Upload Your Files

You need to upload these 3 files to GitHub:

#### 3a. Create the folder structure first

1. On your computer, create a folder called **`email-organizer`**
2. Inside that folder, create another folder called **`.github`**
3. Inside the **`.github`** folder, create a folder called **`workflows`**

Your folder structure should look like:
```
email-organizer/
├── .github/
│   └── workflows/
```

#### 3b. Download and place the files

1. **File 1**: `email_sorter.py` 
   - Place in the main `email-organizer` folder

2. **File 2**: `requirements.txt`
   - Place in the main `email-organizer` folder

3. **File 3**: `build.yml`
   - Place in the `.github/workflows/` folder

Your complete structure:
```
email-organizer/
├── .github/
│   └── workflows/
│       └── build.yml
├── email_sorter.py
└── requirements.txt
```

#### 3c. Upload to GitHub

1. Go back to your repository on GitHub
2. Click **"uploading an existing file"**
3. Drag ALL 3 files into the upload area (you can drag the whole folder)
4. Scroll down and click **"Commit changes"**

✅ Files uploaded!

---

### Step 4: Let GitHub Build Your .exe

1. Click the **"Actions"** tab at the top of your repository
2. You should see **"Build Email Organizer"** running
3. Wait 3-5 minutes (grab a coffee ☕)
4. When you see a **green checkmark ✅**, it's done!

---

### Step 5: Download Your .exe File

**Method A: From Artifacts (Immediate Download)**
1. Click on the green workflow that just completed
2. Scroll down to **"Artifacts"**
3. Click **"EmailOrganizer-Windows"**
4. A .zip file will download
5. Extract the .zip file (right-click → Extract All)
6. Inside you'll find **`EmailOrganizer.exe`**

**Method B: From Releases (Better for Sharing)**
1. Click the **"Releases"** section on the right side
2. Click on the latest release
3. Download **`EmailOrganizer.exe`**

✅ You now have your .exe file!

---

## 🎮 How to Use Your App

### First Time Running

1. Double-click **`EmailOrganizer.exe`**
2. Windows might show a warning:
   - Click **"More info"**
   - Click **"Run anyway"**
   - This is normal for unsigned programs

### Using the App

1. **Window opens** with a blue drop zone
2. **Drag & drop** your .eml or .msg email files into the blue area
   - Or click **"📂 Browse Files"** to select files
3. Click **"⚡ Process & Export"**
4. Choose where to save your Excel file
5. **Done!** Your emails are organized by topic

### Understanding the Excel File

Your Excel file will have 4 sheets:

1. **Emails**: All your emails with topic labels
2. **Summary**: How many emails per topic
3. **TopicMap**: Rename topics (e.g., "Topic_1" → "Customer Support")
4. **Index**: Tracks which emails were processed

---

## 🆘 Common Problems

### "Windows protected your PC"
**Solution**: Click "More info" → "Run anyway"

### Drag & drop doesn't work
**Solution**: Use the "📂 Browse Files" button instead

### "All emails already processed"
**Solution**: These emails are already in your Excel file (check the Index sheet)

### Can't open Excel file
**Solution**: Close Excel if it's already open

---

## 🎉 You're Done!

You now have a working email organizer that:
- ✅ Sorts emails by topic automatically
- ✅ Works completely offline
- ✅ Keeps your emails private
- ✅ Never overwrites your manual edits

Enjoy organizing your emails! 📧✨

---

## 💾 Updating the App Later

If you want to make changes:

1. Edit the files on your computer
2. Go to your GitHub repository
3. Click "Add file" → "Upload files"
4. Upload the changed files
5. GitHub will automatically build a new .exe
6. Download from Actions → Artifacts

---

**Need more help? Create an Issue on GitHub and ask for assistance!**
