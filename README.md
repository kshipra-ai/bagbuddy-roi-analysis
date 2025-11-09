# ROI Analysis Repository

This repository contains ROI (Return on Investment) analysis tools and calculations for two key stakeholder groups: **Brands** and **Investors**.

## Repository Structure

```
roi-analysis/
├── brand/              # ROI analysis for brands
│   ├── README.md
│   ├── roi_calculator.py
│   └── metrics.md
├── investor/           # ROI analysis for investors
│   ├── README.md
│   ├── roi_calculator.py
│   └── metrics.md
└── README.md          # This file
```

## Purpose

This repository provides separate ROI calculation frameworks tailored to the distinct needs of:

### 1. Brands
Track marketing campaign performance, customer acquisition costs, engagement rates, and revenue attribution.

**Key Metrics:**
- Campaign ROI
- Customer Acquisition Cost (CAC)
- Engagement Rate
- Conversion Rate
- Customer Lifetime Value (CLV)

### 2. Investors
Analyze investment returns, portfolio performance, and growth projections.

**Key Metrics:**
- ROI Percentage
- Annualized Return
- Multiple on Invested Capital (MOIC)
- Internal Rate of Return (IRR)
- Net Present Value (NPV)

## Getting Started

### For Brand Analysis

**Quick Start - Interactive Dashboard:**
```bash
cd brand
python dashboard.py
```

**Or run individual tools:**
```bash
cd brand
python roi_calculator.py    # Single campaign analysis
python comparison.py         # Compare multiple campaigns
python visualization.py      # View charts and insights
python examples.py           # Learn through examples
```

**📚 Documentation:**
- `brand/INDEX.md` - **Start here!** Getting started guide
- `brand/README.md` - Complete user guide
- `brand/QUICK_REF.md` - Quick reference
- `brand/metrics.md` - Metric definitions

**✅ Status:** Fully implemented with 14 files including:
- 6 executable Python scripts
- 10 sample campaigns
- Comprehensive documentation
- Interactive dashboard
- Export capabilities

### For Investor Analysis
```bash
cd investor
python roi_calculator.py
```

See `investor/README.md` for detailed usage instructions.

**⏳ Status:** Basic implementation (will be enhanced next)

## Requirements

- Python 3.8+
- No external dependencies required for basic calculations

## Contributing

When adding new metrics or calculations:
1. Update the relevant calculator file (`roi_calculator.py`)
2. Document new metrics in `metrics.md`
3. Update the README with usage examples

## License

TBD

## Contact

TBD
