
# 💼 Finance Project – Automated Financial Dashboard with Power BI

## 📘 Overview

This project solves a **real-world operational challenge** of manually processing 20+ financial Excel files received via email every day and turning them into an interactive dashboard — all within just 5 hours.

The manual process was **time-consuming**, **error-prone**, and **costly**.
I **automated** the entire workflow — from email to dashboard — using free tools like **Outlook**, **Power Automate**, **Google Cloud Platform**, **Python**, and **Power BI**.

⚡️ This automation saved hours of **manual effort**, **reduced errors**, and **eliminated the need for additional labor**.

---

## 🚀 Problem Statement

### 🔍 Current Process:
- **Emails received:** 25+ Excel/CSV files daily from a survey team across various locations.
- **Deadline:** Processed and analyzed to produce a dashboard by **8 PM** (emails come at **3 PM**).
- **Manual Steps:**
  1. Download files from email.
  2. Combine them into a single file.
  3. Clean the data for accuracy.
  4. Build the dashboard.
  5. Share insights with the client.

### ⚠️ Challenges Faced:
- Manual processing was **time-consuming and error-prone**.
- Required hiring **2 extra employees (+$12,000/month cost)**.
- Errors delayed reports and reduced accuracy.
- Workload was impacting other projects.

---

## 🔧 Solution Workflow

- **Step 1: Email Automation**  
  → Used **Outlook Rules** to filter emails based on subject keywords like "Daily Financial Report".  
  → Saved these emails automatically into a specific Outlook folder.  

- **Step 2: File Extraction with Power Automate**  
  → Connected Outlook to **Google Drive** using Power Automate.  
  → Automated the extraction of attachments from filtered emails and stored them in a designated Google Drive folder.  

- **Step 3: Google Drive API Connection**  
  → Created API credentials using **Google Cloud Platform**.  
  → Gave folder access to the generated service account email.  
  → Used the API key to establish a secure connection between **Google Drive and Python**.  

- **Step 4: Python + API Integration**  
  → Wrote **Python scripts** to fetch all Excel/CSV files from Google Drive using the API.  
  → Merged the multiple files into a single dataset using **pandas** in Python.  
  → Passed the merged dataset to **Power BI** for further processing.  

- **Step 5: Data Cleaning & Transformation in Power Query (Power BI)**  
  → Performed complete **data cleaning and transformation using Power Query** in Power BI.  
  → Tasks included:  
    ▪ Removing duplicates  
    ▪ Handling null or missing values  
    ▪ Formatting and standardizing data types  
    ▪ Renaming columns for clarity  
    ▪ Filtering irrelevant data  
    ▪ Creating additional calculated columns for analysis  

- **Step 6: Dashboard Development in Power BI**  
  → Created a dynamic and interactive **dashboard with real-time data refresh**.  
  → Built **KPIs, charts, visuals, and slicers for deep analysis**.  
  → Applied **DAX** for advanced measures and business logic.  

---

## 🔥 Key Metrics Displayed
- **Average Annual Income**
- **Average Monthly Balance**
- **Average Number of Delays in Payment**
- **Average Credit Utilization**

---

## 🛠 Tools Used
- Power BI  
- Power Query  
- DAX  
- Python  
- Outlook  
- Power Automate  
- Google Drive API  

---

## 📂 Folder Structure

- **Finance-Project/**
  - [Financial_Project.pbix](./Financial_Project.pbix)
  - [Financial Data Problem Statements.pdf](./Financial%20Data%20Problem%20Statements.pdf)
  - [Gdrive_to_powerbi.py](./Gdrive_to_powerbi.py)
  - **Data/**
    - [combined_part_1.xlsx](./Data/combined_part_1.xlsx) — (Similar 25 files)
  - **Screenshots/**
    - [Finance Project Page 1.png](./Screenshots/Finance%20Project%20Page%201.png)
    - [Finance Project Page 2.png](./Screenshots/Finance%20Project%20Page%202.png)
  - [API_Key.json](./API_Key.json)
  - README.md

---

## 🖼️ Dashboard Preview

### 🔹 Page 1
![Dashboard Screenshot 1](./Screenshots/Finance%20Project%20Page%201.png) 
### 🔹 Page 2 
![Dashboard Screenshot 2](./Screenshots/Finance%20Project%20Page%202.png)

---

## 🚀 How to Run This Project

Follow these steps to set up and run the complete financial data automation and dashboard system:

---

### 📨 Step 1: Receive and Filter Emails in Outlook

- Ensure all financial data files are sent to a **dedicated email inbox** (e.g., `finance@yourdomain.com`).
- In **Outlook**, create a **rule** that:
  - Filters emails based on subject keywords (e.g., “Daily Financial Report”).
  - Saves attachments or moves these emails into a dedicated folder (e.g., “Finance Reports”).

---

### 🔁 Step 2: Automate File Transfer with Power Automate

- Open [Power Automate](https://powerautomate.microsoft.com/).
- Create a **new flow** that:
  1. **Triggers** when a new email with attachments is received in the specified Outlook folder.
  2. **Extracts attachments** from the email.
  3. **Saves files to a Google Drive folder** (e.g., `/Finance-Reports/`).
- Ensure the same folder is shared with your **Google Cloud service account** email (you’ll get this after creating the API key).

---

### ☁️ Step 3: Set Up Google Cloud API

- Visit [Google Cloud Console](https://console.cloud.google.com/).
- Create a project and enable the **Google Drive API**.
- Create a **service account** and download the **JSON key**.
- Share your Google Drive folder with the **service account email** (`xxx@project-id.iam.gserviceaccount.com`).

---

### 🐍 Step 4: Configure and Run Python Script

- Open the [Gdrive_to_powerbi.py](./Gdrive_to_powerbi.py) script.
- Replace the placeholders:
  - `your-json-path.json` → path to your downloaded service account key.
  - `folder_id` → your Google Drive folder ID.

- Install required Python packages if not already:
  ```bash
  pip install pandas google-api-python-client google-auth
  ```

- Run the script:
  ```bash
  python Gdrive_to_powerbi.py
  ```

---

### 📊 Step 5: Open and Refresh Power BI Dashboard

- Launch **Power BI Desktop**.
- Open [Financial_Project.pbix](./Financial_Project.pbix).
- Click **Refresh** to update all visuals and data models with the newly combined dataset.

---

This pipeline allows you to go from **email attachments → Google Drive → cleaned dataset → Power BI dashboard**, all automated and seamless

---

## 💡 Business Impact

- Saved **5 hours daily** of manual work.
- Eliminated **$12,000/month** in extra labor costs.
- Reduced error rates significantly.
- Improved reporting speed and accuracy.
- Enhanced decision-making with real-time insights.

---

## 🧠 Conclusion

This project demonstrates how automation combined with Power BI can transform traditional manual processes into efficient, scalable, and error-free workflows. The insights derived have been directly used by higher management to make data-driven decisions and offer customer-specific promotions.

---

## 🤝 Contact

**Shivam Wankhade**  
[LinkedIn](https://www.linkedin.com/in/shivam96k) | [GitHub](https://github.com/Shiva96k) | ✉️ Email: shivamwankhade180@gmail.com  

---

## ✨ Feedback Welcomed

Feel free to contribute or suggest improvements. Pull requests are open!

---