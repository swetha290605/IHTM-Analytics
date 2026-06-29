# 🚀 Quick Reference Card

## Installation (One-Time)

```bash
# Python packages
pip install pandas openpyxl numpy

# Node packages (optional)
npm install
```

## Start Dashboard

```bash
# Terminal 1: File watcher (auto-rebuild Excel changes)
npm run dev

# Terminal 2: Web server
python3 server.py

# Terminal 3: Auto-build watcher (if using npm)
npm run build:watch
```

**Then open:** http://localhost:8000/dashboard.html

## File Locations

```
your-folder/
├── dashboard.html              ← Open this in browser
├── process_analytics.py        ← Data processor
├── server.py                   ← Web server
├── watch.js                    ← File watcher
├── package.json               ← Node dependencies
├── analytics_data.json        ← Generated data (auto-updated)
└── data/                      ← PUT YOUR EXCEL FILES HERE
    ├── Query_Analytics.xlsx
    ├── Safety_Status.xlsx
    ├── IHTM_Revenue.xlsx
    ├── Field_Issues.xlsx
    ├── Cost_Competency.xlsx
    ├── Project_Punch_Points.xlsx
    └── Tools_Under_Progress.xlsx
```

## Common Commands

```bash
# Rebuild data manually
python3 process_analytics.py

# Change port (if 8000 is in use)
python3 -m http.server 8001

# Install npm dependencies
npm install

# Start file watcher
npm run dev

# Update specific tool
npm run build
```

## Excel Files at a Glance

| File | Purpose | Key Columns |
|------|---------|-------------|
| Query_Analytics.xlsx | Project pipeline | Project Name, Stage, Plant |
| Safety_Status.xlsx | Safety incidents | Month, No of Accidents |
| IHTM_Revenue.xlsx | Revenue tracking | Month, Revenue, Target |
| Field_Issues.xlsx | Issue tracking | Month, Identified, Solved |
| Cost_Competency.xlsx | Cost comparison | Month, Local Value, Target Value |
| Project_Punch_Points.xlsx | Punch points | Month, Item, Punch Points |
| Tools_Under_Progress.xlsx | Tool tracking | Plant, Shop, Budget Type, Cost |

## Dashboard Navigation

```
🏭 Manufacturing Analytics Dashboard

[🌙 Dark Mode]

TABS:
├─ 📋 Query Analytics
│  ├─ % by Stage (pie)
│  ├─ % by Plant (pie)
│  └─ Total Queries (card)
│
├─ 🛡️ Safety Status
│  ├─ Monthly Accidents (bar)
│  └─ Safety Calendar (grid)
│
└─ ⚙️ Others (Expandable)
   ├─ 💰 IHTM Revenue
   │  ├─ Revenue vs Target (bar)
   │  └─ Metrics (cards)
   │
   ├─ 🐛 Field Issues (bar/line)
   ├─ 📈 Cost Competency (line)
   ├─ 📌 Punch Points (bar)
   └─ 🔧 Tools Progress
      ├─ Summary (cards)
      └─ Distribution (pies)
```

## Data Update Flow

```
You edit Excel → File detected → Python processes → JSON updated → Dashboard refreshes
     ↓                ↓              ↓                    ↓              ↓
  (2-3 sec)      watch.js     process_analytics.py   auto-saved    Green banner
```

## Troubleshooting Quick Fixes

| Problem | Solution |
|---------|----------|
| Port 8000 in use | `python3 -m http.server 8001` |
| Data not updating | `python3 process_analytics.py` |
| Charts not showing | Hard refresh: Ctrl+Shift+R |
| File not recognized | Check filename matches exactly |
| No green notification | Browser may have cached data |
| Dark mode not working | Clear cache (DevTools → Storage) |

## Browser Shortcuts

| Action | Keys |
|--------|------|
| Hard Refresh | Ctrl+Shift+R (Win) / Cmd+Shift+R (Mac) |
| Developer Tools | F12 or Ctrl+Shift+I |
| Toggle Dark Mode | Click 🌙 in top right |
| Switch Tab | Click [📋] [🛡️] [⚙️] buttons |

## Excel Column Names (Must Match Exactly)

### Query_Analytics.xlsx
```
Project Name | Customer Name | Plant | IHTM Leader | Stage | Date of Stage | Tool No | Remarks
```

### Safety_Status.xlsx
```
Month | No of Accidents
```

### IHTM_Revenue.xlsx
```
Month | Revenue (Million Rs) | Target (Million Rs)
```

### Field_Issues.xlsx
```
Month | Identified | Solved | Cum Identified | Cum Solved
```

### Cost_Competency.xlsx
```
Month | Local Value | Target Value | IHTM Value | Cost in Manufacturing
```

### Project_Punch_Points.xlsx
```
Month | Item | Item Count | Punch Points
```

### Tools_Under_Progress.xlsx
```
Project | Plant | Shop | Budget Type | Month | In Mill | No of Tools
```

## Data Examples

### Query_Analytics.xlsx
```
Project Name | Customer Name | Plant | Stage
Proj-001     | Acme Corp     | P1    | Query
Proj-002     | TechFlow      | P2    | Estimation
Proj-003     | GlobalInd     | P3    | Budget Confirmed
```

### Safety_Status.xlsx
```
Month    | No of Accidents
January  | 2
February | 1
March    | 0
```
Color Codes: 0=🟢, 1=🔵, 2+=🔴

### Revenue.xlsx
```
Month    | Revenue (Million Rs) | Target (Million Rs)
January  | 45.5                 | 50
February | 52.3                 | 50
March    | 48.7                 | 50
```

## Customization Tips

**Change Colors:** Edit CSS in `dashboard.html`
```css
--primary: #FF385C;      /* Change to your brand color */
--success: #007a3d;      /* Green */
--warning: #854F0B;      /* Orange */
```

**Add New Chart:** 
1. Add canvas: `<canvas id="c-chart-name"></canvas>`
2. Add render function in JavaScript
3. Add data in Python script
4. Call render function in `updateAll()`

**Change Chart Type:**
- `mk(id, 'bar', ...)` → Bar chart
- `mk(id, 'line', ...)` → Line chart
- `donut(id, ...)` → Pie/donut chart

## Performance Tips

- Keep Excel files under 1000 rows for best performance
- Use consistent data formats
- Archive old data periodically
- Clear browser cache monthly
- Restart server weekly

## Production Checklist

- [ ] All Excel files have real data
- [ ] Column names match exactly
- [ ] Port 8000 is accessible
- [ ] File permissions are correct
- [ ] Python dependencies installed
- [ ] Analytics_data.json generates without errors
- [ ] All charts display correctly
- [ ] Dark mode works
- [ ] Mobile responsive verified
- [ ] Test auto-update (edit Excel, watch refresh)

---

**Need Help?** Check README.md or SETUP_GUIDE.md for detailed documentation.
