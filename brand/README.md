# Brand ROI Analysis

This directory contains comprehensive ROI (Return on Investment) analysis tools and calculations specifically for brands.

## Purpose

Track and analyze the return on investment for brands participating in the platform, including:
- Campaign performance metrics
- Customer acquisition costs
- Engagement rates
- Revenue attribution
- Brand awareness impact
- Multi-campaign comparison
- Performance visualization

## Files

- **`roi_calculator.py`** - Python script for calculating brand ROI metrics
  - ROI percentage
  - Customer Acquisition Cost (CAC)
  - Click-Through Rate (CTR)
  - Conversion Rate (CVR)
  - Cost per Click (CPC)
  - Cost per Acquisition (CPA)
  - Return on Ad Spend (ROAS)
  - Profit margins
  - Revenue per customer
  - Break-even analysis
  - Performance grading

- **`comparison.py`** - Compare multiple campaigns
  - Identify best and worst performers
  - Calculate averages across campaigns
  - Generate comparison reports
  - Provide optimization recommendations

- **`visualization.py`** - Data visualization tools
  - ASCII bar charts
  - Summary tables
  - Performance insights
  - Campaign comparisons

- **`examples.py`** - Real-world usage scenarios
  - Single campaign analysis
  - Multi-campaign comparison
  - Performance optimization
  - CLV calculations
  - Budget planning
  - Seasonal campaign strategies

- **`sample_data.csv`** - Sample campaign data template
- **`metrics.md`** - Documentation of key metrics and KPIs
- **`requirements.txt`** - Python dependencies (optional)

## Usage

### Quick Start - Single Campaign Analysis

```bash
python roi_calculator.py
```

This will run the example campaign and display a comprehensive ROI report.

### Compare Multiple Campaigns

```bash
python comparison.py
```

Analyzes all campaigns in `sample_data.csv` and generates comparison reports.

### Visualize Campaign Performance

```bash
python visualization.py
```

Creates ASCII charts and performance summaries.

### Run All Examples

```bash
python examples.py
```

Demonstrates 6 real-world scenarios:
1. Single campaign analysis
2. Campaign comparison
3. Performance visualization
4. Customer Lifetime Value (CLV) calculation
5. Optimizing underperforming campaigns
6. Seasonal campaign budget planning

## Using Your Own Data

### Option 1: CSV File

Create a CSV file with the following columns:
```
campaign_id,campaign_name,start_date,end_date,total_investment,total_revenue,
new_customers,impressions,engagements,clicks,conversions,ad_spend,
creative_cost,platform_fee
```

Then load it:
```python
from comparison import CampaignComparison

comparison = CampaignComparison()
comparison.load_from_csv('your_data.csv')
print(comparison.generate_comparison_report())
```

### Option 2: Dictionary

```python
from roi_calculator import BrandROICalculator

campaign = {
    'campaign_name': 'Your Campaign',
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
print(report)
```

## Key Metrics Explained

- **ROI**: Return on Investment - Overall profitability percentage
- **CAC**: Customer Acquisition Cost - Cost to acquire each customer
- **ROAS**: Return on Ad Spend - Revenue per dollar spent
- **CTR**: Click-Through Rate - Percentage of impressions that resulted in clicks
- **CVR**: Conversion Rate - Percentage of clicks that resulted in conversions
- **CPC**: Cost Per Click - Average cost for each click
- **CPA**: Cost Per Acquisition - Average cost for each conversion

## Performance Grades

Campaigns are automatically graded based on ROI:
- **A+**: 500%+ ROI
- **A**: 400-499% ROI
- **A-**: 300-399% ROI
- **B+**: 200-299% ROI
- **B**: 150-199% ROI
- **B-**: 100-149% ROI
- **C**: 50-99% ROI
- **D**: 0-49% ROI
- **F**: Negative ROI

## Requirements

- Python 3.8+
- No external dependencies required for basic functionality
- Optional: matplotlib, pandas for advanced visualizations (not currently implemented)

## Next Steps

1. Replace `sample_data.csv` with your actual campaign data
2. Run analyses to identify top performers
3. Use insights to optimize underperforming campaigns
4. Plan future campaigns based on historical performance
