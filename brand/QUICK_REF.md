# Brand ROI Quick Reference Guide

## Quick Start Commands

```bash
# Run interactive dashboard
python dashboard.py

# Single campaign analysis
python roi_calculator.py

# Compare all campaigns
python comparison.py

# View visualizations
python visualization.py

# Run all examples
python examples.py

# Export data
python export.py
```

## Common Use Cases

### 1. Calculate ROI for a New Campaign

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
roi = calculator.calculate_roi()
print(f"ROI: {roi}%")
```

### 2. Compare Multiple Campaigns

```python
from comparison import CampaignComparison

comparison = CampaignComparison()
comparison.load_from_csv('sample_data.csv')
print(comparison.generate_comparison_report())
```

### 3. Find Best Performers

```python
from comparison import CampaignComparison

comparison = CampaignComparison()
comparison.load_from_csv('sample_data.csv')
best = comparison.find_best_performers()

print(f"Best ROI: {best['highest_roi']['campaign_name']}")
print(f"ROI: {best['highest_roi']['roi']}%")
```

### 4. Export Reports

```python
from export import ROIDataExporter
from comparison import CampaignComparison

comparison = CampaignComparison()
comparison.load_from_csv('sample_data.csv')

exporter = ROIDataExporter()
comparisons = comparison.compare_all()
best = comparison.find_best_performers()
averages = comparison.calculate_averages()

# Export to JSON
exporter.export_comparison_report(comparisons, best, averages)

# Export to CSV
exporter.create_excel_friendly_csv(comparisons, 'report')
```

## Key Metrics Formulas

### ROI (Return on Investment)
```
ROI = ((Revenue - Investment) / Investment) × 100
```

### CAC (Customer Acquisition Cost)
```
CAC = Total Investment / Number of New Customers
```

### ROAS (Return on Ad Spend)
```
ROAS = Total Revenue / Total Investment
```

### CTR (Click-Through Rate)
```
CTR = (Clicks / Impressions) × 100
```

### CVR (Conversion Rate)
```
CVR = (Conversions / Clicks) × 100
```

### Engagement Rate
```
Engagement Rate = (Engagements / Impressions) × 100
```

### CPC (Cost Per Click)
```
CPC = Total Investment / Clicks
```

### CPA (Cost Per Acquisition)
```
CPA = Total Investment / Conversions
```

### Profit Margin
```
Profit Margin = ((Revenue - Investment) / Revenue) × 100
```

### Revenue Per Customer
```
RPC = Total Revenue / Number of Customers
```

### Customer Lifetime Value (CLV)
```
CLV = Avg Purchase Value × Purchase Frequency × Customer Lifespan
```

## Performance Benchmarks

| Metric | Poor | Good | Excellent |
|--------|------|------|-----------|
| ROI | < 100% | 100-300% | > 300% |
| ROAS | < 2x | 2-4x | > 4x |
| CAC | > $100 | $50-$100 | < $50 |
| Engagement Rate | < 2% | 2-5% | > 5% |
| CTR | < 1% | 1-3% | > 3% |
| CVR | < 1% | 1-3% | > 3% |

## Performance Grades

- **A+ (500%+)**: Outstanding performance
- **A (400-499%)**: Excellent results
- **A- (300-399%)**: Very good campaign
- **B+ (200-299%)**: Good ROI
- **B (150-199%)**: Above average
- **B- (100-149%)**: Average performance
- **C (50-99%)**: Below average, needs optimization
- **D (0-49%)**: Poor performance
- **F (Negative)**: Losing money

## Data File Format

CSV file should include these columns:

```
campaign_id          - Unique identifier
campaign_name        - Campaign name
start_date          - Start date (YYYY-MM-DD)
end_date            - End date (YYYY-MM-DD)
total_investment    - Total spent on campaign
total_revenue       - Total revenue generated
new_customers       - Number of new customers acquired
impressions         - Number of ad impressions
engagements         - Total engagements (likes, shares, etc.)
clicks              - Number of clicks
conversions         - Number of conversions/sales
ad_spend            - Amount spent on ads
creative_cost       - Cost of creating campaign materials
platform_fee        - Platform/service fees
```

## Troubleshooting

### "No data loaded" error
- Ensure `sample_data.csv` exists in the brand folder
- Check file path when loading custom CSV
- Verify CSV format matches template

### Division by zero warnings
- Occurs when metrics are 0 (e.g., no clicks)
- Calculator returns 0 in these cases
- Check input data for missing values

### Import errors
- Ensure you're running from the brand folder
- All modules should be in the same directory
- No external dependencies required

## Tips for Better ROI

1. **Improve Conversion Rate**
   - Better targeting
   - Optimize landing pages
   - A/B test creative

2. **Reduce CAC**
   - Focus on high-performing channels
   - Improve audience targeting
   - Optimize ad creative

3. **Increase Customer Lifetime Value**
   - Encourage repeat purchases
   - Upsell/cross-sell
   - Improve customer retention

4. **Optimize Ad Spend**
   - Focus budget on top performers
   - Pause underperforming campaigns
   - Test new channels with small budgets

5. **Track and Analyze**
   - Monitor metrics regularly
   - Compare campaigns
   - Learn from top performers

## Support

For questions or issues:
1. Check the main README.md
2. Review metrics.md for metric definitions
3. Run examples.py for usage demonstrations
4. Review sample_data.csv for data format
