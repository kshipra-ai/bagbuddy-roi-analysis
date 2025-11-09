"""
Digital Ads ROI Calculator

This module calculates ROI metrics for digital advertising campaigns.

RESEARCH-BASED BENCHMARKS (2024-2025):
Sources to reference:
1. WordStream Google Ads Benchmarks (2024)
2. Meta (Facebook/Instagram) Ads Benchmarks Report (2024)
3. Google Ads Industry Benchmark Report (2024)
4. Smartly.io Digital Advertising Benchmarks (2024)
5. eMarketer Digital Ad Spending & Performance Report (2024)

Realistic Industry Benchmarks by Platform & Industry:

GOOGLE ADS (Search):
- Average CPM: $10 - $50 (varies by industry/competition)
- Average CPC: $1 - $3 (most industries)
- Average CTR: 3% - 5% (search ads - intent-based)
- Average Conversion Rate: 2.5% - 5% (landing page quality dependent)

FACEBOOK/INSTAGRAM (Meta):
- Average CPM: $5 - $15 (Meta 2024 data)
- Average CPC: $0.50 - $2.00
- Average CTR: 0.9% - 1.5% (interruption-based ads)
- Average Conversion Rate: 2% - 4% (post-click)

DISPLAY ADS (GDN, Programmatic):
- Average CPM: $2 - $10
- Average CPC: $0.30 - $1.50
- Average CTR: 0.1% - 0.5% (very low for display)
- Average Conversion Rate: 0.5% - 1.5%

CONSERVATIVE REALISTIC NUMBERS (for small business):
We'll use Facebook/Instagram benchmarks as they're most common for local businesses:
- CPM: $10.00 (realistic for 2024-2025, was lower in past)
- CTR: 1.0% (realistic for decent ad creative)
- Conversion Rate: 2.5% (realistic for decent landing page)

NOTE: These assume:
- Decent ad creative and copy
- Targeted audience (not broad)
- Optimized landing page
- Clear call-to-action

Without these, expect 50% worse performance.
"""

from typing import Dict, Any


class DigitalAdsROICalculator:
    """Calculate ROI metrics for digital advertising campaigns."""
    
    # Research-based industry benchmarks (Meta 2024, WordStream 2024)
    DEFAULT_CPM = 10.00  # $10 per 1000 impressions (Facebook/Instagram 2024-2025)
    DEFAULT_CTR = 1.0  # 1% click-through rate (realistic for decent creative)
    DEFAULT_CONVERSION_RATE = 2.5  # 2.5% conversion rate (realistic for decent LP)
    
    def __init__(self, campaign_data: Dict[str, Any]):
        """
        Initialize the calculator with campaign data.
        
        Args:
            campaign_data: Dictionary containing campaign metrics
                Required keys:
                - ad_budget: Total budget for digital ads
                - cpm: Cost per 1000 impressions (optional, defaults to $10)
                - ctr_percent: Click-through rate percentage (optional, defaults to 1%)
                - conversion_rate_percent: Conversion rate (optional, defaults to 2.5%)
                - avg_revenue_per_conversion: Average revenue per sale
        """
        self.campaign_data = campaign_data
    
    def calculate_campaign_cost(self) -> float:
        """Get the total ad budget."""
        return self.campaign_data.get('ad_budget', 0)
    
    def calculate_cpm(self) -> float:
        """Get Cost per 1000 impressions."""
        return self.campaign_data.get('cpm', self.DEFAULT_CPM)
    
    def calculate_total_impressions(self) -> int:
        """
        Calculate total impressions based on budget and CPM.
        
        Formula: (Budget / CPM) × 1000
        """
        budget = self.calculate_campaign_cost()
        cpm = self.calculate_cpm()
        
        if cpm == 0:
            return 0
        
        impressions = (budget / cpm) * 1000
        return int(impressions)
    
    def calculate_ctr(self) -> float:
        """Get click-through rate percentage."""
        return self.campaign_data.get('ctr_percent', self.DEFAULT_CTR)
    
    def calculate_total_clicks(self) -> int:
        """
        Calculate total clicks.
        
        Formula: Impressions × (CTR / 100)
        """
        impressions = self.calculate_total_impressions()
        ctr = self.calculate_ctr() / 100
        
        clicks = int(impressions * ctr)
        return clicks
    
    def calculate_conversion_rate(self) -> float:
        """Get conversion rate percentage."""
        return self.campaign_data.get('conversion_rate_percent', self.DEFAULT_CONVERSION_RATE)
    
    def calculate_conversions(self) -> int:
        """
        Calculate total conversions.
        
        Formula: Clicks × (Conversion Rate / 100)
        """
        clicks = self.calculate_total_clicks()
        conversion_rate = self.calculate_conversion_rate() / 100
        
        conversions = int(clicks * conversion_rate)
        return conversions
    
    def calculate_avg_revenue_per_conversion(self) -> float:
        """Get average revenue per conversion."""
        return self.campaign_data.get('avg_revenue_per_conversion', 0)
    
    def calculate_total_sales(self) -> float:
        """
        Calculate total sales generated.
        
        Formula: Conversions × Avg Revenue per Conversion
        """
        conversions = self.calculate_conversions()
        avg_revenue = self.calculate_avg_revenue_per_conversion()
        
        total_sales = conversions * avg_revenue
        return round(total_sales, 2)
    
    def calculate_roi(self) -> float:
        """Calculate ROI percentage."""
        total_sales = self.calculate_total_sales()
        cost = self.calculate_campaign_cost()
        
        if cost == 0:
            return 0.0
        
        roi = ((total_sales - cost) / cost) * 100
        return round(roi, 2)
    
    def calculate_cost_per_click(self) -> float:
        """Calculate actual CPC."""
        cost = self.calculate_campaign_cost()
        clicks = self.calculate_total_clicks()
        
        if clicks == 0:
            return 0.0
        
        cpc = cost / clicks
        return round(cpc, 2)
    
    def calculate_cost_per_conversion(self) -> float:
        """Calculate CPA (Cost per Acquisition)."""
        cost = self.calculate_campaign_cost()
        conversions = self.calculate_conversions()
        
        if conversions == 0:
            return 0.0
        
        cpa = cost / conversions
        return round(cpa, 2)
    
    def calculate_roas(self) -> float:
        """Calculate Return on Ad Spend."""
        total_sales = self.calculate_total_sales()
        cost = self.calculate_campaign_cost()
        
        if cost == 0:
            return 0.0
        
        roas = total_sales / cost
        return round(roas, 2)
    
    def get_full_report(self) -> Dict[str, Any]:
        """Generate comprehensive digital ads ROI report."""
        return {
            # Campaign Setup
            'campaign_cost': self.calculate_campaign_cost(),
            'cpm': self.calculate_cpm(),
            'ctr_percent': self.calculate_ctr(),
            'conversion_rate_percent': self.calculate_conversion_rate(),
            
            # Performance Metrics
            'total_impressions': self.calculate_total_impressions(),
            'total_clicks': self.calculate_total_clicks(),
            'total_conversions': self.calculate_conversions(),
            
            # Revenue Metrics
            'avg_revenue_per_conversion': self.calculate_avg_revenue_per_conversion(),
            'total_sales': self.calculate_total_sales(),
            
            # ROI & Cost Metrics
            'roi_percent': self.calculate_roi(),
            'cost_per_click': self.calculate_cost_per_click(),
            'cost_per_conversion': self.calculate_cost_per_conversion(),
            'roas': self.calculate_roas()
        }


if __name__ == "__main__":
    # Example: Digital Ads Campaign with REALISTIC industry numbers
    sample_campaign = {
        'campaign_name': 'Facebook Ads - Coffee Shop',
        'ad_budget': 1000,  # $1,000 ad budget
        'cpm': 10.00,  # $10.00 per 1000 impressions (Meta 2024-2025 realistic)
        'ctr_percent': 1.0,  # 1.0% click-through rate (WordStream 2024)
        'conversion_rate_percent': 2.5,  # 2.5% conversion rate (realistic for SMB)
        'avg_revenue_per_conversion': 25  # $25 average purchase
    }
    
    calculator = DigitalAdsROICalculator(sample_campaign)
    report = calculator.get_full_report()
    
    # Calculate breakdown components
    impressions = report['total_impressions']
    clicks = report['total_clicks']
    conversions = report['total_conversions']
    
    print("\n" + "=" * 80)
    print(f"DIGITAL ADS ROI REPORT: {sample_campaign['campaign_name']}")
    print("=" * 80 + "\n")
    
    print("CAMPAIGN SETUP")
    print("-" * 80)
    print(f"Ad Budget: ${report['campaign_cost']:,.2f}")
    print(f"CPM (Cost per 1000 impressions): ${report['cpm']:.2f}")
    print(f"Target CTR: {report['ctr_percent']}%")
    print(f"Target Conversion Rate: {report['conversion_rate_percent']}%\n")
    
    print("PERFORMANCE FUNNEL (DETAILED BREAKDOWN)")
    print("-" * 80)
    print(f"Step 1 - Impressions:")
    print(f"  Budget ${report['campaign_cost']:,.2f} ÷ CPM ${report['cpm']:.2f} × 1000")
    print(f"  = {impressions:,} impressions\n")
    
    print(f"Step 2 - Clicks:")
    print(f"  {impressions:,} impressions × {report['ctr_percent']}% CTR")
    print(f"  = {clicks:,} clicks\n")
    
    print(f"Step 3 - Conversions:")
    print(f"  {clicks:,} clicks × {report['conversion_rate_percent']}% conversion rate")
    print(f"  = {conversions:,} conversions\n")
    
    print(f"Step 4 - Revenue:")
    print(f"  {conversions:,} conversions × ${report['avg_revenue_per_conversion']:.2f} avg sale")
    print(f"  = ${report['total_sales']:,.2f} total sales\n")
    
    print("ROI & COST METRICS")
    print("-" * 80)
    print(f"Total Sales: ${report['total_sales']:,.2f}")
    print(f"Total Cost: ${report['campaign_cost']:,.2f}")
    print(f"Net Profit: ${report['total_sales'] - report['campaign_cost']:,.2f}")
    print(f"ROI: {report['roi_percent']}%")
    print(f"ROAS: {report['roas']}x (${report['roas']:.2f} revenue per $1 spent)\n")
    
    print("COST EFFICIENCY")
    print("-" * 80)
    print(f"Cost per Click (CPC): ${report['cost_per_click']:.2f}")
    print(f"Cost per Conversion (CPA): ${report['cost_per_conversion']:.2f}\n")
    
    print("=" * 80)
    print(f"\n💡 Summary:")
    print(f"   A ${report['campaign_cost']:,.2f} digital ad campaign generated")
    print(f"   {impressions:,} impressions, {clicks:,} clicks, and {conversions:,} conversions,")
    print(f"   resulting in ${report['total_sales']:,.2f} in sales ({report['roi_percent']}% ROI).\n")
    print("=" * 80)
