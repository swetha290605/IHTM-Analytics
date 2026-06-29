# Manufacturing Analytics Dashboard - Setup & Usage Guide

## 📦 What You're Getting

This is a complete manufacturing analytics dashboard system with:

1. **7 Excel Spreadsheet Templates** (in `data/` folder)
2. **HTML Dashboard** with light/dark mode
3. **Python Data Processor** for converting Excel → JSON
4. **Node.js File Watcher** (chokidar) for auto-updates
5. **Development Server** for local testing
6. **Auto-Update System** - Dashboard refreshes when data changes

---

## ✅ Quick Start (5 Minutes)

### Step 1: Download & Extract
- Unzip all files to a folder (e.g., `C:\manufacturing-analytics\` or `~/manufacturing-analytics/`)

### Step 2: Install Dependencies
```bash
# Open terminal/command prompt in the folder

# Install Python packages (one-time)
pip install pandas openpyxl numpy

# Install Node packages (one-time, optional but recommended)
npm install
```

### Step 3: Run Dashboard
```bash
# Terminal 1: Start file watcher (optional)
npm run dev

# Terminal 2: Start web server
python3 server.py

# Terminal 3 (if using npm watch): Keep running
npm run build:watch
```

### Step 4: Open in Browser
Open: **http://localhost:8000/dashboard.html**

---

## 📊 Excel Files Explained

All Excel files go in the `data/` folder.

### File 1: Query_Analytics.xlsx
**Purpose**: Track projects through different stages

**Required Columns:**
| Column | Example | Notes |
|--------|---------|-------|
| Project Name | Proj-001 | Unique identifier |
| Customer Name | Acme Corp | Client name |
| Plant | P1, P2, P3, Camry, Others | Manufacturing location |
| IHTM Leader | Raj Kumar | Project owner |
| Stage | Query, Estimation, Budget Confirmed, etc. | Current pipeline stage |
| Date of Stage | 2024-01-15 | When stage was reached |
| Tool No | T001 | Tool identifier |
| Remarks | Initial enquiry | Any notes |

**Dashboard Shows:**
- 📊 Pie chart: % of projects by stage
- 📊 Bar chart: % of projects by plant
- 📍 Summary card: Total number of queries

---

### File 2: Safety_Status.xlsx
**Purpose**: Track safety incidents

**Required Columns:**
| Column | Example | Notes |
|--------|---------|-------|
| Month | January, February, etc. | Month name |
| No of Accidents | 0, 1, 2, 3 | Number of incidents |

**Color Coding:**
- 🟢 **Green**: 0 accidents (No Accident)
- 🔵 **Blue**: 1 accident (NLWD - No Lost Work Days)
- 🔴 **Red**: 2+ accidents (LWD - Lost Work Days)

**Dashboard Shows:**
- 📊 Bar chart: Accidents by month
- 📅 Safety Calendar: 12-month grid with color coding
- 🎯 Clickable months to see details

---

### File 3: IHTM_Revenue.xlsx
**Purpose**: Track revenue against targets

**Required Columns:**
| Column | Example | Notes |
|--------|---------|-------|
| Month | January, February, etc. | Month name |
| Revenue (Million Rs) | 45.5 | Actual revenue |
| Target (Million Rs) | 50 | Monthly target |

**Dashboard Shows:**
- 📊 Bar chart: Revenue vs Target
- 📍 Cards: Total revenue, Yearly target
- 📈 Metric: % of yearly target achieved
- 💰 Metric: Required average per 3 months

**How It Works:**
- If yearly target is ₹700M and you've earned ₹400M
- Days remaining / 3 = number of 3-month periods left
- Calculates what you need to earn per quarter

---

### File 4: Field_Issues.xlsx
**Purpose**: Track issue identification and resolution

**Required Columns:**
| Column | Example | Notes |
|--------|---------|-------|
| Month | January, February, etc. | Month name |
| Identified | 12 | Issues found this month |
| Solved | 8 | Issues resolved this month |
| Cum Identified | 12 | Cumulative total identified |
| Cum Solved | 8 | Cumulative total solved |

**Dashboard Shows:**
- 📊 Hybrid chart: Bar chart (monthly) + Line chart (cumulative)
- Tracks both monthly and cumulative trends

---

### File 5: Cost_Competency.xlsx
**Purpose**: Compare costs across different metrics

**Required Columns:**
| Column | Example | Notes |
|--------|---------|-------|
| Month | January, February, etc. | Month name |
| Local Value | 35.2 | Local manufacturing cost |
| Target Value | 36 | Target cost |
| IHTM Value | 37.5 | IHTM cost |
| Cost in Manufacturing | 28.5 | Direct manufacturing cost |

**Dashboard Shows:**
- 📊 Line chart: All 4 cost metrics
- Compare trends across the year
- Identify cost patterns

---

### File 6: Project_Punch_Points.xlsx
**Purpose**: Track items and punch points

**Required Columns:**
| Column | Example | Notes |
|--------|---------|-------|
| Month | January, February, etc. | Month name |
| Item | Design Review, Manufacturing, etc. | Item category |
| Item Count | 8 | Number of items |
| Punch Points | 24 | Punch points (usually 3x item count) |

**Example Data:**
```
January  | Design Review | 8  | 24
January  | Manufacturing | 6  | 18
February | Design Review | 12 | 36
```

**Dashboard Shows:**
- 📊 Grouped bar chart by item type
- Monthly breakdown

---

### File 7: Tools_Under_Progress.xlsx
**Purpose**: Track tools in development and their budget allocation

**Required Columns:**
| Column | Example | Notes |
|--------|---------|-------|
| Project | TOOL-001 | Tool identifier |
| Plant | P1, P2, P3, Camry, Others | Plant location |
| Shop | Press, Weld, Assy, Camry, QI, QC, Others | Workshop type |
| Budget Type | Capex, Revenue | Type of budget |
| Month | January, February, etc. | Month |
| In Mill | 2.5 | Cost in millions |
| No of Tools | 5 | Number of tools |

**Dashboard Shows:**
- 📍 Card: Total cost (in millions)
- 📍 Card: Total number of tools
- 📊 Pie chart: % by budget type (Capex vs Revenue)
- 📊 Pie chart: % by plant (P1, P2, P3, Camry, Others)
- 📊 Pie chart: % by plant + shop combination

---

## 🎮 How to Use the Dashboard

### Main Interface

**Top Bar:**
- 🏭 Logo and title
- 🌙/☀️ **Dark Mode Toggle** - Click to switch themes

**KPI Strip (4 Cards):**
- Total Queries
- Tools Progress
- Safety Incidents
- Revenue Target %

**Tabs Below KPI:**
- 📋 **Query Analytics** - First section (enabled by default)
- 🛡️ **Safety Status** - Second section
- ⚙️ **Others** - All remaining analytics (Revenue, Issues, Cost, Punch, Tools)

### Tab Navigation

Click any tab to switch between sections:
```
[📋 Query Analytics] [🛡️ Safety Status] [⚙️ Others]
                ↑
        Currently active
```

### Interactive Elements

**Safety Calendar:**
- Shows 12 months in a grid
- Click any month to see accident details
- Colors change based on incident count

**Dark Mode:**
- Click 🌙 icon to enable dark mode
- Your preference is saved automatically
- Works on all devices

---

## 🔄 Auto-Update Process

### How File Monitoring Works

1. **You edit an Excel file** in the `data/` folder
   ```
   data/Query_Analytics.xlsx ← You modify this
   ```

2. **File Watcher detects change** (within 2 seconds)
   ```
   watch.js → "Change detected!"
   ```

3. **Python processes the data**
   ```
   process_analytics.py → Reads all Excel files
   ```

4. **JSON is regenerated**
   ```
   analytics_data.json ← Updated with new data
   ```

5. **Dashboard auto-refreshes**
   ```
   JavaScript polls every 3 seconds
   ↓
   New data detected
   ↓
   Green banner: "🔄 Data updated — refreshing…"
   ↓
   Charts refresh with new data
   ```

### Manual Update

If auto-update doesn't work:

**Option 1: Run Python directly**
```bash
python3 process_analytics.py
```
Then refresh browser (Ctrl+R / Cmd+R)

**Option 2: Manually stop and restart services**
```bash
# Stop current services (Ctrl+C in terminals)

# Rebuild data
python3 process_analytics.py

# Restart server
python3 server.py

# Refresh browser
```

---

## 🎨 Customizing the Dashboard

### Change Colors

Edit `dashboard.html`, find the CSS section at the top:

```css
:root {
  --primary: #FF385C;        /* Main red - change to your brand color */
  --success: #007a3d;        /* Green */
  --warning: #854F0B;        /* Orange */
}
```

Examples:
- Blue theme: `--primary: #0066FF;`
- Green theme: `--primary: #00AA44;`
- Purple theme: `--primary: #9933FF;`

### Add New Section

1. Open `dashboard.html`
2. Add new `<button class="tab-btn">` in tab navigation
3. Add corresponding `<div class="panel">`
4. Add render function in JavaScript
5. Add data processing in `process_analytics.py`

### Change Chart Types

In `dashboard.html`, JavaScript section:
- `'bar'` → Bar chart
- `'line'` → Line chart
- `'doughnut'` → Pie/donut chart
- `'scatter'` → Scatter plot
- `'radar'` → Radar chart

---

## 📱 Using on Different Devices

### Desktop (1200px+)
- Full 4-column KPI strip
- Side-by-side charts
- Best for detailed analysis

### Tablet (700-1200px)
- 2-column layouts
- Responsive charts
- Touch-friendly controls

### Mobile (<700px)
- Single column layout
- Full-width charts
- Optimized for small screens
- Swipeable tabs

---

## 🔧 Troubleshooting

### Problem: Port 8000 is already in use
**Solution:**
```bash
# Use a different port
python3 -m http.server 8001

# Visit: http://localhost:8001/dashboard.html
```

### Problem: "Failed to load data" message
**Solutions:**
1. Check if `analytics_data.json` exists
2. Run `python3 process_analytics.py`
3. Check browser console (F12 → Console tab)
4. Look for error messages

### Problem: Charts not showing data
**Solutions:**
1. Hard refresh: Ctrl+Shift+R (Windows) or Cmd+Shift+R (Mac)
2. Clear cache: Open DevTools → Storage → Clear all
3. Check if Excel files are in `data/` folder
4. Verify Excel file names match exactly

### Problem: File watcher not auto-updating
**Solutions:**
1. Ensure Node.js is installed: `node --version`
2. Check chokidar installed: `npm list chokidar`
3. Try manual rebuild: `python3 process_analytics.py`
4. Restart services (Ctrl+C and run again)

### Problem: Excel file shows as modified but data didn't update
**Solutions:**
1. Save file completely (make sure editor finished)
2. Wait 2-3 seconds for watcher to detect
3. Check if file saved in correct format (.xlsx, not .xls)
4. Look at terminal output for errors

---

## 📊 Data Tips

### Organizing Data
- Keep files in `data/` folder only
- Use consistent month names (January, February, etc.)
- Avoid special characters in names
- Use numbers for counts/amounts

### Data Entry Best Practices
- One project per row in Query_Analytics
- Always include all required columns
- Use exact stage names (copy-paste to avoid typos)
- Format dates consistently

### Performance
- Dashboard handles up to 1000s of rows
- Charts show most recent 12-20 items
- Large datasets (10k+ rows) may slow down

---

## 🚀 Advanced Usage

### Integrating with Other Tools

**Excel VBA Macro** - Auto-save and run Python:
```vba
Sub UpdateDashboard()
    Shell "python3 process_analytics.py"
End Sub
```

**Scheduling Updates** - Run Python on schedule:

**Windows Task Scheduler:**
- Create task to run `process_analytics.py`
- Set frequency (hourly, daily, etc.)

**Mac/Linux Cron:**
```bash
# Edit crontab
crontab -e

# Add line (runs every hour)
0 * * * * /usr/bin/python3 /path/to/process_analytics.py
```

### Hosting Online

**Simple Hosting (GitHub Pages):**
1. Put files on GitHub
2. Enable GitHub Pages
3. Share public URL

**VPS/Cloud Server:**
1. Upload files to server
2. Install Python packages
3. Use proper web server (Nginx, Apache)
4. Set up SSL certificate

---

## 📞 Support Checklist

If something doesn't work:

- [ ] Python 3.8+ installed? `python3 --version`
- [ ] pandas/openpyxl installed? `pip show pandas`
- [ ] Node.js installed? `node --version`
- [ ] All Excel files in `data/` folder?
- [ ] analytics_data.json exists?
- [ ] Server running? Terminal shows "Server is running"
- [ ] Correct URL? `http://localhost:8000/dashboard.html`
- [ ] Browser console clear? F12 → Console tab

---

## 📚 File Descriptions

### Core Files

**dashboard.html** (31 KB)
- Complete UI, charts, and JavaScript
- No external dependencies except Chart.js
- Self-contained styling and functionality

**process_analytics.py** (13 KB)
- Reads all Excel files
- Aggregates and transforms data
- Outputs JSON for dashboard
- Run: `python3 process_analytics.py`

**server.py** (2.2 KB)
- Simple Python web server
- Serves dashboard and data
- Run: `python3 server.py`

**watch.js** (1.8 KB)
- File watcher for Excel changes
- Requires Node.js
- Run: `node watch.js`

**package.json** (595 B)
- Node.js dependencies list
- Install with: `npm install`

**README.md** (11 KB)
- Complete documentation
- Reference guide
- Troubleshooting

### Data Files

**data/Query_Analytics.xlsx**
- Sample project data
- Can be replaced with real data

**data/Safety_Status.xlsx**
- Sample safety incidents
- Update with real data

**data/IHTM_Revenue.xlsx**
- Sample revenue data
- Add your actual numbers

**data/Field_Issues.xlsx**
- Sample issue tracking
- Update as needed

**data/Cost_Competency.xlsx**
- Sample cost data
- Replace with actual costs

**data/Project_Punch_Points.xlsx**
- Sample punch point data
- Update monthly

**data/Tools_Under_Progress.xlsx**
- Sample tools data
- Track tools in development

**analytics_data.json** (13 KB)
- Generated automatically
- Do not edit manually
- Regenerated on each run of Python script

---

## 🎯 Next Steps

1. **Extract all files** to a folder
2. **Install dependencies**: `pip install pandas openpyxl`
3. **Replace sample data** with your real Excel files
4. **Start server**: `python3 server.py`
5. **Open dashboard**: http://localhost:8000/dashboard.html
6. **Toggle dark mode** using the 🌙 icon
7. **Explore all tabs** to see your data

---

## 📝 Notes

- Dashboard works offline (no internet required)
- All data stays on your computer
- Browser theme preference is saved
- Charts update in real-time
- Mobile-friendly design
- Works in all modern browsers

---

**Version**: 1.0.0  
**Last Updated**: June 2026  
**Status**: Production Ready ✅
