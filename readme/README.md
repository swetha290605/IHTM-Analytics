# 🏭 Manufacturing Analytics Dashboard

A comprehensive, real-time manufacturing analytics dashboard with support for multiple Excel data sources. Features light/dark mode, auto-update on file changes, and beautiful visualizations.

## 📋 Features

✅ **Multiple Analytics Sections:**
- Query Analytics (stage distribution, plant-wise breakdown)
- Safety Status (accident tracking, safety calendar)
- IHTM Revenue (monthly tracking with targets)
- Field Issues (identification and resolution tracking)
- Cost Competency (cost comparison analysis)
- Project Punch Points (item tracking)
- Tools Under Progress (budget and plant distribution)

✅ **Smart UI/UX:**
- Light and Dark mode toggle
- Real-time KPI dashboard
- Responsive design (mobile, tablet, desktop)
- Smooth transitions and animations
- Interactive charts with Chart.js

✅ **Auto-Update:**
- File watcher monitors Excel files
- Automatic data processing on changes
- Live reload dashboard (every 3 seconds)
- Visual notification banner

---

## 📦 Installation

### Prerequisites
- Python 3.8+
- Node.js 14+ (optional, for file watching with chokidar)
- pandas, openpyxl (Python libraries)

### Step 1: Install Python Dependencies
```bash
pip install pandas openpyxl numpy
```

### Step 2: Install Node Dependencies (Optional but Recommended)
```bash
npm install
# or
yarn install
```

---

## 🚀 Quick Start

### Option 1: Python Only (Simple Setup)
```bash
# Generate JSON from Excel files
python3 process_analytics.py

# Start local server
python3 server.py

# Open browser
# Visit: http://localhost:8000/dashboard.html
```

### Option 2: With Node File Watching (Recommended)
```bash
# Install dependencies
npm install

# Start watcher + server
npm run dev

# In another terminal:
python3 server.py

# Open browser
# Visit: http://localhost:8000/dashboard.html
```

### Option 3: Full Auto-Build + Watch
```bash
npm run build:watch

# In another terminal:
python3 server.py
```

---

## 📊 Excel File Format

All Excel files must be placed in the `data/` folder with the following structure:

### 1. Query_Analytics.xlsx
Required columns:
- `Project Name` - Project identifier
- `Customer Name` - Customer/client name
- `Plant` - Plant location (P1, P2, P3, Camry, Others)
- `IHTM Leader` - Project leader name
- `Stage` - Query, Speculation, Concept, Estimation, Waiting for Budget, Budget Confirmed, Waiting for User Confirmation
- `Date of Stage` - Date of stage change
- `Tool No` - Tool identifier
- `Remarks` - Additional notes

### 2. Safety_Status.xlsx
Required columns:
- `Month` - Month name (January, February, etc.)
- `No of Accidents` - Number of incidents (0 = green, 1 = NLWD, >1 = LWD)

### 3. IHTM_Revenue.xlsx
Required columns:
- `Month` - Month name
- `Revenue (Million Rs)` - Revenue in millions
- `Target (Million Rs)` - Monthly target

### 4. Field_Issues.xlsx
Required columns:
- `Month` - Month name
- `Identified` - Issues identified
- `Solved` - Issues resolved
- `Cum Identified` - Cumulative identified
- `Cum Solved` - Cumulative solved

### 5. Cost_Competency.xlsx
Required columns:
- `Month` - Month name
- `Local Value` - Local cost value
- `Target Value` - Target cost
- `IHTM Value` - IHTM cost
- `Cost in Manufacturing` - Manufacturing cost

### 6. Project_Punch_Points.xlsx
Required columns:
- `Month` - Month name
- `Item` - Item category (Design Review, Manufacturing, etc.)
- `Item Count` - Count of items
- `Punch Points` - Number of punch points

### 7. Tools_Under_Progress.xlsx
Required columns:
- `Project` - Project identifier
- `Plant` - Plant (P1, P2, P3, Camry, Others)
- `Shop` - Shop type (Press, Weld, Assy, Camry, QI, QC, Others)
- `Budget Type` - Capex or Revenue
- `Month` - Month name
- `In Mill` - Cost in millions
- `No of Tools` - Number of tools

---

## 📁 File Structure

```
manufacturing-analytics/
├── dashboard.html              # Main dashboard (HTML + JS + CSS)
├── process_analytics.py        # Python data processor
├── server.py                   # Development server
├── watch.js                    # File watcher script
├── package.json               # Node dependencies
├── analytics_data.json        # Generated data (auto-created)
├── data/
│   ├── Query_Analytics.xlsx
│   ├── Safety_Status.xlsx
│   ├── IHTM_Revenue.xlsx
│   ├── Field_Issues.xlsx
│   ├── Cost_Competency.xlsx
│   ├── Project_Punch_Points.xlsx
│   └── Tools_Under_Progress.xlsx
└── README.md
```

---

## 🎨 Dashboard Sections

### 🔴 Query Analytics
- **% Projects by Stage** - Pie chart showing distribution across pipeline stages
- **% Projects by Plant** - Plant-wise distribution
- **Total Queries** - Summary card

### 🛡️ Safety Status
- **Monthly Accidents** - Bar chart with accident trends
- **Safety Calendar** - Interactive 12-month grid
  - 🟢 Green = No accidents (0)
  - 🔵 Blue = NLWD (1 accident)
  - 🔴 Red = LWD (2+ accidents)

### ⚙️ Others (Expandable Section)

#### 💰 IHTM Revenue
- Revenue vs Target bar chart
- Total revenue and yearly target cards
- Target achievement percentage
- Required revenue per 3 months calculation

#### 🐛 Field Issues
- Multi-metric bar/line hybrid chart
- Tracks: Identified, Solved, Cumulative Identified, Cumulative Solved

#### 📈 Cost Competency
- Line chart comparing:
  - Local Value
  - Target Value
  - IHTM Value
  - Cost in Manufacturing

#### 📌 Project Punch Points
- Grouped bar chart by item type
- Monthly punch point tracking

#### 🔧 Tools Under Progress
- Total cost and tools summary cards
- % by Budget Type (pie)
- % by Plant (pie)
- % by Plant + Shop (pie)

---

## 🎭 Theme Toggle

Click the **moon icon** (☀️/🌙) in the top-right corner to toggle between:
- **Light Mode** - Clean, bright interface (default)
- **Dark Mode** - Easy on the eyes in low-light conditions

Theme preference is saved in browser localStorage.

---

## 🔄 Auto-Update Process

### File Change Detection
1. Excel files in `data/` folder are monitored
2. When a file is modified, the watcher detects it
3. `process_analytics.py` is automatically triggered
4. Data is processed and `analytics_data.json` is regenerated
5. Dashboard polls for updates every 3 seconds
6. When new data is detected, charts auto-refresh
7. Green notification banner appears: "🔄 Data updated"

### Manual Rebuild
```bash
python3 process_analytics.py
```

---

## 💻 Development

### Modifying Dashboard Elements

#### Add a New Chart
1. Open `dashboard.html`
2. Find the relevant panel
3. Add a new `<div class="card">` with a canvas element
4. Add the render function in the JavaScript section
5. Ensure the chart ID matches the canvas ID

#### Change Colors
Edit the CSS variables at the top of `dashboard.html`:
```css
:root {
  --primary: #FF385C;        /* Main red color */
  --success: #007a3d;        /* Green */
  --warning: #854F0B;        /* Orange */
  /* ... */
}
```

#### Update Data Processing
Edit `process_analytics.py`:
1. Add new functions for data aggregation
2. Update the result dictionary
3. Write the JSON output
4. Dashboard will auto-reload on next poll

---

## 🐛 Troubleshooting

### Issue: Port 8000 Already in Use
```bash
# Use a different port
python3 -m http.server 8001
# Then visit: http://localhost:8001/dashboard.html
```

### Issue: Excel File Not Loading
- Ensure file is in `data/` folder
- Check filename exactly matches the script (Query_Analytics.xlsx, etc.)
- Verify Excel format is `.xlsx` (not `.xls` or `.csv`)
- Run: `python3 process_analytics.py` to see errors

### Issue: Charts Not Displaying
- Check browser console (F12 → Console tab)
- Ensure `analytics_data.json` exists
- Try hard-refresh (Ctrl+Shift+R)
- Clear localStorage: `localStorage.clear()` in console

### Issue: File Watcher Not Triggering
- Ensure Node.js and chokidar are installed: `npm list chokidar`
- Check file permissions on data folder
- Try saving file with a new filename first

---

## 📱 Responsive Breakpoints

- **Desktop**: 1200px+ (4-column KPI strip)
- **Tablet**: 700px-1200px (2-column layouts)
- **Mobile**: <700px (1-column layouts)

---

## 🔒 Data Privacy

- All data processing happens locally
- No data is sent to external servers
- localStorage only stores theme preference

---

## 📝 API Reference

### JSON Output Structure (analytics_data.json)
```json
{
  "meta": {
    "generated_at": "ISO timestamp",
    "source": "Manufacturing Analytics"
  },
  "global_kpis": {
    "total_queries": number,
    "total_tools": number,
    "safety_accidents": number
  },
  "query_analytics": {
    "stage_distribution": [...],
    "plant_distribution": [...],
    "total_queries": number
  },
  "safety": {
    "month_accidents": [...],
    "safety_calendar": [...]
  },
  "revenue": {
    "revenue_by_month": [...],
    "total_revenue": number,
    "yearly_target": number,
    "target_achievement_pct": number,
    "needed_per_3months": number
  },
  "field_issues": {...},
  "cost_competency": {...},
  "punch_points": {...},
  "tools": {...}
}
```

---

## 🚀 Production Deployment

### Using a Production Server
Replace `server.py` with a proper web server:

**Option 1: Gunicorn + Flask**
```bash
pip install flask gunicorn
# Create app.py serving static files
gunicorn -w 4 -b 0.0.0.0:8000 app:app
```

**Option 2: Nginx**
```nginx
server {
  listen 80;
  server_name your-domain.com;
  root /path/to/manufacturing-analytics;
  
  location / {
    try_files $uri $uri/ =404;
  }
}
```

**Option 3: Docker**
```dockerfile
FROM python:3.9-slim
WORKDIR /app
COPY . .
RUN pip install pandas openpyxl
CMD ["python3", "server.py"]
```

---

## 📧 Support

For issues or questions:
1. Check the Troubleshooting section
2. Review browser console (F12)
3. Check Python script output
4. Verify Excel file format and column names

---

## 📄 License

MIT License - Feel free to use and modify

---

## 🎯 Roadmap

- [ ] Export dashboard as PDF
- [ ] Email scheduled reports
- [ ] Database backend instead of JSON
- [ ] User authentication & roles
- [ ] Advanced filtering by date range
- [ ] Custom dashboard builder
- [ ] API endpoint for third-party integration

---

**Last Updated**: June 2026
**Version**: 1.0.0
