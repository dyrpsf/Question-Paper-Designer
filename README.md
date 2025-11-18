Question Paper Generator
A powerful Python application that automatically scrapes questions from multiple educational websites, detects relevant diagrams, and exports formatted question papers to DOCX format.

🚀 Features
Multi-URL Support: Import questions from multiple websites simultaneously

Smart Deduplication: Automatically removes duplicate questions across sources

Diagram Detection: Captures relevant images and diagrams near questions

MCQ Support: Properly formats multiple-choice questions with options

DOCX Export: Creates professionally formatted Word documents

User-Friendly GUI: Built with CustomTkinter for modern interface

Cross-Platform: Works on Windows (executable provided)

📋 Prerequisites
Windows OS (for the provided executable)

Python 3.8+ (if running from source)

Chrome Browser (for web scraping functionality)

🛠️ Installation
Method 1: Using Executable (Recommended for End Users)
Download the Question Paper Designer.zip file

Extract the zip file to your desired location:

Right-click the zip file

Select "Extract All..."

Choose destination folder

Click "Extract"

Run the application:

Navigate to the extracted folder

Double-click Question Paper Designer.exe

Allow access if Windows Defender prompts (first time only)

Method 2: From Source Code
bash
# Clone the repository
git clone https://github.com/yourusername/question-paper-generator.git

# Install dependencies
pip install -r requirements.txt

# Run the application
python main.py
📖 How to Use
Step 1: Specify Number of Websites
Enter how many website URLs you want to scrape

Click "Next" to proceed

Step 2: Paste URLs
Copy URLs from educational websites containing questions

Important: Questions must be in text format (not PDF)

Paste each URL in separate entry boxes

Click "Load Questions"

Step 3: Select Questions
Review extracted questions in the list

Use checkboxes to select questions for export

Use "Select All"/"Deselect All" for bulk operations

Click "Preview" to view diagrams (if available)

Step 4: Export to DOCX
Click "Export Selected Questions to DOCX"

Choose save location and filename

Wait for completion confirmation

⚙️ Features in Detail
Smart Question Detection
Identifies questions based on patterns and keywords

Captures MCQ options when properly formatted

Handles various question numbering styles

Image Processing
Automatically detects diagrams near questions

Supports multiple image formats (PNG, JPG, WebP, etc.)

Converts WebP to PNG for DOCX compatibility

Limits images per question to avoid clutter

Export Formatting
Clean numbering and formatting

Proper option labeling (a), b), c) style)

Image scaling to fit page width

Professional document structure

🎯 Tips for Best Results
Choose Compatible Websites: Sites with clean HTML structure work best

Check Preview: Always preview diagrams before final export

Multiple Sources: Combine questions from different websites for variety

Manual Review: Some websites may have layout issues affecting image detection

❗ Known Limitations
PDF-based question banks are not supported

Some complex website layouts may affect image detection

WebP images require Pillow library for conversion

Very large websites may take longer to process

🐛 Troubleshooting
Windows Defender Warning:

This is normal for first-time execution

Click "More info" → "Run anyway"

Slow Loading:

Some websites may take time to render completely

Progress bar shows current status

Missing Images:

Check if the website uses complex JavaScript

Verify images are in supported formats

📞 Support
For issues and suggestions:

Check the user manual for detailed instructions

Ensure all prerequisites are met

Verify website compatibility

📄 License
This project is provided for educational and personal use.

Thank you for using Question Paper Generator! 🎉