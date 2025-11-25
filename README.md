# 📊 Automated XML → Google Sheets Data Pipeline  
### **Python Extractor + Google Apps Script Sheet Organizer**

This project automates the extraction, cleaning, classification, and organization of dataset records from daily XML files.  
It consists of:

1. **Python Script (`main.py`):**  
   - Reads an XML file generated daily  
   - Extracts attributes  
   - Cleans & filters invalid rows  
   - Classifies unit descriptions (e.g., STUDIO, 1BR, 2BR, 3BR+)  
   - Uploads processed rows to a Google Sheet  
   - Sends email notifications for both success & failures  

2. **Google Apps Script (`Code.gs`):**  
   - Automatically organizes Sheet1 into multiple sheets  
   - Creates one sheet per **Group Code**  
   - Mirrors formatting from the main sheet  
   - Removes obsolete sheets  
   - Formats number fields  
   - Supports both UI and backend triggers  

---

## 🚀 Features

### **Python Script**
- ✔ Reads XML from a shared folder  
- ✔ Auto-detects file encoding  
- ✔ XML parsing with attribute flattening  
- ✔ Row exclusion rules  
- ✔ Google Sheets API upload  
- ✔ Gmail API notifications  
- ✔ Description classification logic  
- ✔ Auto-refresh of Google OAuth token  

---

### **Google Apps Script**
- ✔ Creates UI menu (“Sheet Organizer”)  
- ✔ Syncs all sheets with one click  
- ✔ Removes obsolete sheets  
- ✔ Creates new sheets per Group Code  
- ✔ Copies original formatting (header + column widths)  
- ✔ Formats numeric columns (#,##0.00)  
- ✔ Supports onChange, webhook, or manual trigger 
