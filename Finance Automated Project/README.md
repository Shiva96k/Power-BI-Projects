
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

## 🚀 How to Run

1. Clone the repository.
2. Set up Google Cloud API credentials.
3. Replace the path to your JSON key in [Gdrive_to_powerbi.py](./Gdrive_to_powerbi.py).
4. Replace your Google Drive **folder ID** in [Gdrive_to_powerbi.py](./Gdrive_to_powerbi.py).
5. Run the Python script to fetch and combine data files.
6. Open [Financial_Project.pbix](./Financial_Project.pbix) in **Power BI Desktop**.
7. Refresh the dashboard to view the latest data.

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