# Brand ROI Analysis - Getting Started Guide

Welcome to the Brand ROI Analysis Toolkit! This guide will help you get started quickly.

## 📁 What's in This Folder?

### 🚀 Executable Scripts (Run These!)

1. **`dashboard.py`** - 🎯 **START HERE!**
   - Interactive menu-driven interface
   - Easiest way to use all features
   - No coding required
   ```bash
   python dashboard.py
   ```

2. **`roi_calculator.py`** - Calculate ROI for a single campaign
   ```bash
   python roi_calculator.py
   ```

3. **`comparison.py`** - Compare multiple campaigns
   ```bash
   python comparison.py
   ```

4. **`visualization.py`** - View charts and insights
   ```bash
   python visualization.py
   ```

5. **`examples.py`** - Learn through 6 real-world examples
   ```bash
   python examples.py
   ```

6. **`export.py`** - Export data to files
   ```bash
   python export.py
   ```

### 📚 Documentation

1. **`README.md`** - Complete user guide (read this first!)
2. **`QUICK_REF.md`** - Quick reference for commands and formulas
3. **`metrics.md`** - Detailed explanation of all metrics
4. **`IMPLEMENTATION_SUMMARY.md`** - Technical implementation details
5. **`CHANGELOG.md`** - Version history and features
6. **`INDEX.md`** - This file!

### 📊 Data Files

1. **`sample_data.csv`** - 10 sample campaigns for testing
2. **`requirements.txt`** - Python dependencies (optional)

---

## 🎯 Quick Start (30 Seconds)

**Option 1: Interactive Dashboard (Recommended)**
```bash
cd c:\kshipra-codebase\roi-analysis\brand
python dashboard.py
```

**Option 2: Run Sample Analysis**
```bash
cd c:\kshipra-codebase\roi-analysis\brand
python comparison.py
```

**Option 3: Learn by Examples**
```bash
cd c:\kshipra-codebase\roi-analysis\brand
python examples.py
```

---

## 📖 How to Use This Toolkit

### Scenario 1: "I want to analyze one campaign"

**Using Dashboard:**
1. Run `python dashboard.py`
2. Choose option `1` (Analyze Single Campaign)
3. Enter your campaign data
4. View results instantly

**Using Code:**
```python
from roi_calculator import BrandROICalculator

campaign = {
    'total_investment': 10000,
    'total_revenue': 25000,
    'new_customers': 150,
    'impressions': 250000,
    'engagements': 12500,
    'clicks': 5000,
    'conversions': 300
}

calculator = BrandROICalculator(campaign)
report = calculator.get_full_report()
print(f"ROI: {report['roi_percentage']}%")
```

### Scenario 2: "I want to compare multiple campaigns"

**Using Dashboard:**
1. Run `python dashboard.py`
2. Choose option `2` (Compare All Campaigns)
3. View comparison report

**Using Code:**
```python
from comparison import CampaignComparison

comparison = CampaignComparison()
comparison.load_from_csv('sample_data.csv')
print(comparison.generate_comparison_report())
```

### Scenario 3: "I want to export reports"

**Using Dashboard:**
1. Run `python dashboard.py`
2. Choose option `5` (Export Data)
3. Select format (JSON, CSV, or Text)

**Using Code:**
```python
from export import ROIDataExporter
from comparison import CampaignComparison

comparison = CampaignComparison()
comparison.load_from_csv('sample_data.csv')

exporter = ROIDataExporter()
comparisons = comparison.compare_all()
best = comparison.find_best_performers()
averages = comparison.calculate_averages()

file = exporter.export_comparison_report(comparisons, best, averages)
print(f"Report saved: {file}")
```

### Scenario 4: "I want to use my own data"

**Step 1:** Create a CSV file with these columns:
```
campaign_id,campaign_name,start_date,end_date,total_investment,
total_revenue,new_customers,impressions,engagements,clicks,
conversions,ad_spend,creative_cost,platform_fee
```

**Step 2:** Use the dashboard:
1. Run `python dashboard.py`
2. Choose option `7` (Load Custom Data File)
3. Enter path to your CSV file
4. Use any analysis option

---

## 🔑 Key Features

### What Can You Calculate?

✅ **ROI** - Return on Investment (%)
✅ **CAC** - Customer Acquisition Cost ($)
✅ **ROAS** - Return on Ad Spend (multiplier)
✅ **CTR** - Click-Through Rate (%)
✅ **CVR** - Conversion Rate (%)
✅ **CPC** - Cost Per Click ($)
✅ **CPA** - Cost Per Acquisition ($)
✅ **Engagement Rate** (%)
✅ **Profit Margin** (%)
✅ **Revenue Per Customer** ($)
✅ **Break-even Point** (number of customers)
✅ **Performance Grade** (A+ to F)

### What Can You Compare?

✅ Multiple campaigns side-by-side
✅ Best vs worst performers
✅ Above/below average campaigns
✅ Metric trends

### What Can You Export?

✅ JSON reports (detailed)
✅ CSV files (Excel-friendly)
✅ Text summaries
✅ Comparison reports

---

## 📊 Understanding Your Results

### Performance Grades

| Grade | ROI | What It Means |
|-------|-----|---------------|
| A+ | 500%+ | 🌟 Outstanding! Keep doing this! |
| A | 400-499% | 🌟 Excellent performance |
| B+ | 200-299% | ✅ Good, profitable campaign |
| B | 100-149% | ✅ Profitable, room to improve |
| C | 50-99% | ⚠️ Weak ROI, needs optimization |
| D | 0-49% | ⚠️ Poor performance |
| F | Negative | ❌ Losing money, stop immediately |

### What's a Good ROI?

- **Minimum:** 100% (2:1 return)
- **Good:** 200% (3:1 return)
- **Excellent:** 300%+ (4:1+ return)

### What's a Good CAC?

- **Rule of Thumb:** CAC should be < 1/3 of Customer Lifetime Value
- **Target:** $50 or less for most campaigns
- **Compare:** Look at your best-performing campaigns

---

## 🎓 Learning Path

**Beginner (Day 1):**
1. ✅ Run `python dashboard.py` and explore
2. ✅ Try option 2 (Compare All Campaigns)
3. ✅ Read the results and insights
4. ✅ Review `QUICK_REF.md` for formulas

**Intermediate (Day 2):**
1. ✅ Run `python examples.py` (all 6 scenarios)
2. ✅ Analyze your own campaign using option 1
3. ✅ Export a report using option 5
4. ✅ Read `metrics.md` for deeper understanding

**Advanced (Day 3+):**
1. ✅ Load your actual campaign data (option 7)
2. ✅ Compare your campaigns
3. ✅ Identify optimization opportunities
4. ✅ Modify the code for custom calculations

---

## 💡 Pro Tips

1. **Start with sample data** - Get familiar with the tools first
2. **Use the dashboard** - Easiest way for most tasks
3. **Export regularly** - Save your analysis for later
4. **Compare often** - Learn from top performers
5. **Read the insights** - The tool provides recommendations
6. **Track trends** - Run analysis weekly/monthly
7. **Share reports** - Export and share with stakeholders

---

## ❓ Common Questions

**Q: Do I need to install anything?**
A: Just Python 3.8+. No external libraries required!

**Q: Can I use my own data?**
A: Yes! Use option 7 in the dashboard or create a CSV file.

**Q: What format should my CSV be?**
A: See `sample_data.csv` for the exact format.

**Q: Can I add more metrics?**
A: Yes! Edit `roi_calculator.py` to add custom calculations.

**Q: How do I export to Excel?**
A: Use option 5 in dashboard, choose CSV format, then open in Excel.

**Q: What if I get division by zero errors?**
A: The calculator handles this automatically, returning 0 for invalid calculations.

---

## 🆘 Need Help?

1. **Check documentation:**
   - `README.md` - Full guide
   - `QUICK_REF.md` - Quick answers
   - `metrics.md` - Metric definitions

2. **Run examples:**
   ```bash
   python examples.py
   ```

3. **Check the code:**
   - All functions have docstrings
   - Comments explain complex logic

---

## 🚀 Next Steps

1. **Run the dashboard** - Start exploring now!
2. **Analyze sample data** - See what insights you can find
3. **Try your own data** - Get real value from your campaigns
4. **Share results** - Export reports for your team
5. **Optimize campaigns** - Use insights to improve ROI

---

## 📞 Support Files

- **README.md** - Detailed usage guide
- **QUICK_REF.md** - Commands and formulas
- **metrics.md** - Metric explanations
- **IMPLEMENTATION_SUMMARY.md** - Technical details
- **CHANGELOG.md** - Version history

---

**🎉 You're ready to go! Start with:**

```bash
python dashboard.py
```

**Have fun analyzing your campaigns! 📊**

---

*Brand ROI Analysis Toolkit v1.0*
*Created: November 8, 2025*
