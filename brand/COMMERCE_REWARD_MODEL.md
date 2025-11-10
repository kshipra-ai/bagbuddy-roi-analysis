# Kshipra Commerce-Reward Business Model ROI Calculator

## Overview
This calculator models Kshipra AI's new **commerce-reward business model** where users purchase bags and earn rewards by watching verified brand advertisements.

## Business Model

### How It Works
1. **Users Purchase Bags** - Pay retail price ($0.40 default)
2. **Watch Brand Ads** - Earn cash credits + reward points per view
3. **Tier-Based Unlock System** - Bronze, Silver, Gold, Platinum tiers limit monthly cash credit unlocks
4. **Redeem at Stores** - Use rewards at participating local merchants
5. **Brands Pay Per View** - CPV model ($0.25 per verified ad view)

### Revenue Streams
- **Bag Sales**: Direct retail revenue from users
- **Brand Advertising**: CPV-based revenue from advertisers
- **Total Margin**: Bag revenue + Kshipra's portion of ad revenue

### User Economics
- Users watch ~3 ads to recover bag cost
- Net user cost per bag: **$-0.05** (bags are FREE after watching ads!)
- Monthly rewards: $1.80 per active user (base scenario)

### Tier System
| Tier | Cash Credit Limit | User Distribution |
|------|------------------|-------------------|
| Bronze | 1 bag value/month | 60% |
| Silver | 3 bag values/month | 25% |
| Gold | 7 bag values/month | 12% |
| Platinum | $1/day (~$30/month) | 3% |

## Files

### Python Scripts
- **`commerce_reward_calculator.py`** - Core calculation engine
  - Revenue per user metrics
  - Revenue per bag metrics
  - Tier economics
  - Store value calculations
  - Brand ROI comparisons
  - Scenario generation (best/base/worst)

- **`export_commerce_reward_excel.py`** - Excel export with formulas
  - Interactive workbook with editable assumptions
  - Dashboard, Assumptions, Calculations sheets
  - All metrics formula-based for dynamic recalculation

### Excel Output
- **`kshipra_commerce_reward_model.xlsx`**
  - **Dashboard**: Key metrics summary
  - **Assumptions**: Editable inputs (YELLOW cells)
  - **Calculations**: Detailed formulas showing how metrics are calculated

## Usage

### Run Python Calculator
```bash
python commerce_reward_calculator.py
```

Shows three scenarios:
- 🎯 Base Scenario (default assumptions)
- 🚀 Best Case (high engagement, premium users)
- ⚠️ Worst Case (low engagement, conservative)

### Generate Excel File
```bash
python export_commerce_reward_excel.py
```

Creates interactive Excel workbook with editable assumptions.

### Customize Assumptions
```python
from commerce_reward_calculator import CommerceRewardCalculator

# Custom assumptions
custom = {
    'bag_retail_price': 0.50,
    'bags_sold_per_month': 20000,
    'avg_monthly_ad_views': 15
}

calc = CommerceRewardCalculator(custom)
calc.print_summary()
```

## Key Metrics

### Base Scenario Results
- **Monthly Revenue**: $34,000
- **Monthly Margin**: $16,000 (47.1%)
- **User Net Bag Cost**: $-0.05 (FREE!)
- **Store Value Created**: $28,500/month
- **Brand CPM**: $250 vs Meta $10 (but 100% verified engagement)

### Per User (Monthly)
- Brand Spend: $3.00
- User Rewards: $1.80
- Kshipra Margin: $1.20

### Per Bag
- Total Revenue: $1.15
- Net Margin: $0.70
- User Earnings: $0.45

## Editable Assumptions

### Pricing & Volume
- Bag retail price ($0.40)
- Bags sold per month (10,000)

### Ad Economics
- CPV - brand pays per view ($0.25)
- Cash credit per view ($0.08)
- Reward points per view ($0.07)
- Kshipra margin per view ($0.10)

### User Behavior
- Avg ads to recover bag cost (3)
- Avg monthly ad views (12)

### Tier Limits
- Bronze/Silver/Gold bag value limits
- Platinum daily cap ($1.00)

### Store Metrics
- Reward redemption rate (75%)
- Repeat visit increase (15%)
- Basket size uplift (8%)
- Average basket value ($25)

### Brand Benchmarks
- Meta CPM ($10)
- TikTok CPM ($8.50)
- Industry avg CTR (1%)

## Scenarios

### Best Case
- High engagement (20 views/user/month)
- More premium users (8% Platinum, 20% Gold)
- High redemption (85%)
- Revenue: $54,000, Margin: $24,000

### Worst Case
- Low engagement (6 views/user/month)
- Fewer premium users (1% Platinum, 5% Gold)
- Low redemption (50%)
- Revenue: $19,000, Margin: $10,000

## Value Propositions

### For Users
- **FREE Bags**: Watch 3 ads → bag is free
- **Earn Rewards**: $1.80/month in cash + points
- **Support Local**: Redeem at neighborhood stores

### For Brands
- **Verified Views**: 100% engagement (vs 1% CTR on social)
- **Cost Effective**: $0.25 per engaged customer
- **Trackable**: Every view linked to redemption behavior

### For Stores
- **$28,500/month** in reward redemptions driving foot traffic
- **15% increase** in repeat visits
- **8% basket uplift** when customers redeem

## Next Steps
1. ✅ Test base calculator
2. ✅ Generate Excel with formulas
3. ✅ Add scenario comparison sheet to Excel
4. ✅ Build sensitivity analysis (vary multiple inputs)
5. ✅ Add investor-specific metrics (LTV, CAC, payback period)

## Excel File Structure

### Dashboard Sheet
Quick overview with key metrics updated in real-time from assumptions.

### Assumptions Sheet (YELLOW = Editable)
All input parameters that drive the model:
- Pricing & Volume
- Ad Economics
- User Behavior
- Tier Limits & Distribution
- Store & Redemption Metrics
- Brand Benchmarks

### Calculations Sheet
Detailed formulas showing how all metrics are calculated:
- Monthly business summary
- Per user metrics
- Per bag metrics
- Store value calculations
- Brand ROI comparisons

### Scenarios Sheet
Pre-built comparison of Best/Base/Worst cases showing:
- Revenue and margin impacts
- Store value created
- Key assumption differences

### Investor Metrics Sheet
Startup-focused metrics with formulas:
- **LTV**: $22.80 (base case)
- **CAC**: $2.00
- **LTV:CAC Ratio**: 11.4x ✅ (Target: >3x)
- **Payback Period**: 1.2 months ✅ (Target: <6 months)
- Annual projections

### Sensitivity Analysis Sheet
Tables showing impact of varying:
1. Monthly ad views (3-25 views)
2. Bag retail price ($0.25-$0.75)
3. Monthly volume (5K-50K bags)

## Demo Script

Run comprehensive demo:
```bash
python demo_commerce_reward.py
```

Shows:
- Basic model calculations
- Scenario comparison
- Investor metrics
- Sensitivity analysis
- Tier-based economics

## Questions This Model Answers
- ✅ How much margin do we make per bag?
- ✅ What's the user's net cost after watching ads?
- ✅ How much value flows back to local stores?
- ✅ What's our CPM vs Meta/TikTok?
- ✅ What happens in best/worst case scenarios?
- ✅ How do tier limits affect economics?

---

**Version**: 1.0  
**Last Updated**: November 2025  
**Author**: Kshipra AI
