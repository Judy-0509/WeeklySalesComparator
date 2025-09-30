# Weekly Excel Compare Tool  

A PyQt5-based desktop application for **comparing weekly Excel datasets**.  
It highlights data changes, computes **YoY (Year-over-Year)** and **WoW (Week-over-Week)** sales trends, visualizes regional performance with charts, and exports modified records back to Excel for further analysis.  

---

## ✨ Features
- 📂 **Drag & Drop or Select Excel Files** to load two datasets  
- 🗓️ **Automatic Week Detection** based on file name pattern (`Model_{week}WeeksYYYY`)  
- 🔎 **Data Comparison**  
  - Detects **deleted** or **changed** rows  
  - Highlights modifications with color-coded results  
- 📊 **Summary Tables**  
  - Regional monthly sales deltas  
  - Cumulative W1~Wn sales with YoY% and WoW% comparisons  
- 🏷️ **Vendor Analysis**  
  - Separate vendor-level YoY/WoW trends by region  
- 📈 **Interactive Charts**  
  - Click on a region to view 3-year weekly sales trends  
- 💾 **Export Results** as Excel (`.xlsx`)  

---

## 📦 Requirements
- Python 3.9+  
- Microsoft Excel (required for `xlwings`)  

Install dependencies:  
```bash
pip install pandas xlwings matplotlib PyQt5


<img width="671" height="307" alt="image" src="https://github.com/user-attachments/assets/7830b951-9623-4c9c-99bd-14680a19bf82" />
