"""
Flyer Campaign ROI Calculator

This module calculates ROI metrics for physical flyer/print advertising campaigns.

RESEARCH-BASED BENCHMARKS (2023-2024):
Sources to reference:
1. USPS "Mail Moment" studies (2023) - Direct mail response rates
2. Data & Marketing Association (DMA) Response Rate Report (2023)
3. PostGrid Direct Mail Marketing Statistics (2024)
4. Gunderson Direct & Digital Marketing Research (2023)

Realistic Industry Benchmarks:
- Flyer printing cost: $0.05 - $0.25 per flyer (depends on quality/quantity)
  * Budget flyers (B&W, standard paper): $0.03 - $0.08
  * Premium flyers (color, glossy): $0.10 - $0.25
  * We'll use: $0.12 (full color, standard glossy)

- Distribution cost: $0.05 - $0.20 per flyer
  * Bulk mail rate: $0.05 - $0.08
  * Door-to-door service: $0.10 - $0.20
  * We'll use: $0.10 (door-to-door distribution)

- Response rate (DMA 2023 data):
  * Prospect lists (cold audience): 0.5% - 1.2%
  * House lists (existing customers): 3.5% - 5%
  * We'll use: 0.8% (cold/prospect flyers - most common scenario)

- Conversion rate from response:
  * Low-involvement products: 5% - 10%
  * High-involvement products: 10% - 20%
  * We'll use: 10% (conservative, realistic)

- Effective conversion rate: 0.8% × 10% = 0.08%
  * This means: ~8 sales per 10,000 flyers distributed
  * Or: ~4 sales per 5,000 flyers

- Typical distribution volume: 5,000 - 50,000 flyers per campaign

NOTE: These numbers reflect REAL INDUSTRY DATA, not aspirational goals.
Most untargeted flyer campaigns have negative or minimal ROI unless:
- Average order value is very high ($100+)
- Targeting is excellent (local, relevant audience)
- Offer is compelling with clear call-to-action
"""

from typing import Dict, Any


class FlyerCampaignROICalculator:
    """Calculate ROI metrics for flyer/print advertising campaigns."""
    
    # Research-based industry benchmarks (DMA 2023, USPS 2024)
    DEFAULT_PRINT_COST = 0.12  # $0.12 per flyer (full color, standard glossy)
    DEFAULT_DISTRIBUTION_COST = 0.10  # $0.10 per flyer (door-to-door)
    DEFAULT_RESPONSE_RATE = 0.8  # 0.8% response rate (prospect lists - DMA 2023)
    DEFAULT_CONVERSION_RATE = 10.0  # 10% of responses convert (realistic)
    
    def __init__(self, campaign_data: Dict[str, Any]):
        """
        Initialize the calculator with campaign data.
        
        Args:
            campaign_data: Dictionary containing campaign metrics
                Required keys:
                - num_flyers: Number of flyers to distribute
                - print_cost_per_flyer: Cost to print each flyer (optional, defaults to $0.12)
                - distribution_cost_per_flyer: Cost to distribute each (optional, defaults to $0.10)
                - response_rate_percent: Percentage who respond (optional, defaults to 0.8%)
                - conversion_rate_percent: Percentage of responses that convert (optional, defaults to 10%)
                - avg_revenue_per_conversion: Average revenue per sale
        """
        self.campaign_data = campaign_data
    
    def calculate_num_flyers(self) -> int:
        """Get total number of flyers distributed."""
        return self.campaign_data.get('num_flyers', 0)
    
    def calculate_print_cost_per_flyer(self) -> float:
        """Get print cost per flyer."""
        return self.campaign_data.get('print_cost_per_flyer', self.DEFAULT_PRINT_COST)
    
    def calculate_distribution_cost_per_flyer(self) -> float:
        """Get distribution cost per flyer."""
        return self.campaign_data.get('distribution_cost_per_flyer', self.DEFAULT_DISTRIBUTION_COST)
    
    def calculate_total_print_cost(self) -> float:
        """Calculate total printing cost."""
        num_flyers = self.calculate_num_flyers()
        cost_per_flyer = self.calculate_print_cost_per_flyer()
        
        total = num_flyers * cost_per_flyer
        return round(total, 2)
    
    def calculate_total_distribution_cost(self) -> float:
        """Calculate total distribution cost."""
        num_flyers = self.calculate_num_flyers()
        cost_per_flyer = self.calculate_distribution_cost_per_flyer()
        
        total = num_flyers * cost_per_flyer
        return round(total, 2)
    
    def calculate_campaign_cost(self) -> float:
        """Calculate total campaign cost (print + distribution)."""
        print_cost = self.calculate_total_print_cost()
        distribution_cost = self.calculate_total_distribution_cost()
        
        total = print_cost + distribution_cost
        return round(total, 2)
    
    def calculate_total_impressions(self) -> int:
        """
        Total impressions = number of flyers distributed.
        Each flyer is assumed to be seen by 1 person.
        """
        return self.calculate_num_flyers()
    
    def calculate_response_rate(self) -> float:
        """Get response rate percentage."""
        return self.campaign_data.get('response_rate_percent', self.DEFAULT_RESPONSE_RATE)
    
    def calculate_responses(self) -> int:
        """
        Calculate number of responses (people who engage).
        
        Formula: Flyers × (Response Rate / 100)
        """
        num_flyers = self.calculate_num_flyers()
        response_rate = self.calculate_response_rate() / 100
        
        responses = int(num_flyers * response_rate)
        return responses
    
    def calculate_conversion_rate(self) -> float:
        """Get conversion rate percentage."""
        return self.campaign_data.get('conversion_rate_percent', self.DEFAULT_CONVERSION_RATE)
    
    def calculate_conversions(self) -> int:
        """
        Calculate total conversions.
        
        Formula: Responses × (Conversion Rate / 100)
        """
        responses = self.calculate_responses()
        conversion_rate = self.calculate_conversion_rate() / 100
        
        conversions = int(responses * conversion_rate)
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
    
    def calculate_cost_per_impression(self) -> float:
        """Calculate cost per impression (cost per flyer distributed)."""
        cost = self.calculate_campaign_cost()
        impressions = self.calculate_total_impressions()
        
        if impressions == 0:
            return 0.0
        
        cpi = cost / impressions
        return round(cpi, 4)
    
    def calculate_cost_per_response(self) -> float:
        """Calculate cost per response/engagement."""
        cost = self.calculate_campaign_cost()
        responses = self.calculate_responses()
        
        if responses == 0:
            return 0.0
        
        cpr = cost / responses
        return round(cpr, 2)
    
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
        """Generate comprehensive flyer campaign ROI report."""
        return {
            # Campaign Setup
            'num_flyers': self.calculate_num_flyers(),
            'print_cost_per_flyer': self.calculate_print_cost_per_flyer(),
            'distribution_cost_per_flyer': self.calculate_distribution_cost_per_flyer(),
            'total_print_cost': self.calculate_total_print_cost(),
            'total_distribution_cost': self.calculate_total_distribution_cost(),
            'campaign_cost': self.calculate_campaign_cost(),
            
            # Performance Metrics
            'total_impressions': self.calculate_total_impressions(),
            'response_rate_percent': self.calculate_response_rate(),
            'total_responses': self.calculate_responses(),
            'conversion_rate_percent': self.calculate_conversion_rate(),
            'total_conversions': self.calculate_conversions(),
            
            # Revenue Metrics
            'avg_revenue_per_conversion': self.calculate_avg_revenue_per_conversion(),
            'total_sales': self.calculate_total_sales(),
            
            # ROI & Cost Metrics
            'roi_percent': self.calculate_roi(),
            'cost_per_impression': self.calculate_cost_per_impression(),
            'cost_per_response': self.calculate_cost_per_response(),
            'cost_per_conversion': self.calculate_cost_per_conversion(),
            'roas': self.calculate_roas()
        }


if __name__ == "__main__":
    # Example: Flyer Campaign with REALISTIC industry numbers
    sample_campaign = {
        'campaign_name': 'Coffee Shop Grand Opening Flyers',
        'num_flyers': 5000,  # 5,000 flyers
        'print_cost_per_flyer': 0.12,  # $0.12 per flyer (full color, glossy)
        'distribution_cost_per_flyer': 0.10,  # $0.10 per flyer (door-to-door)
        'response_rate_percent': 0.8,  # 0.8% response rate (DMA 2023 - prospect lists)
        'conversion_rate_percent': 10,  # 10% of responses result in purchase
        'avg_revenue_per_conversion': 25  # $25 average purchase
    }
    
    calculator = FlyerCampaignROICalculator(sample_campaign)
    report = calculator.get_full_report()
    
    # Calculate breakdown components
    num_flyers = report['num_flyers']
    responses = report['total_responses']
    conversions = report['total_conversions']
    
    print("\n" + "=" * 80)
    print(f"FLYER CAMPAIGN ROI REPORT: {sample_campaign['campaign_name']}")
    print("=" * 80 + "\n")
    
    print("CAMPAIGN COST BREAKDOWN")
    print("-" * 80)
    print(f"Number of Flyers: {num_flyers:,}")
    print(f"Print Cost per Flyer: ${report['print_cost_per_flyer']:.2f}")
    print(f"Distribution Cost per Flyer: ${report['distribution_cost_per_flyer']:.2f}")
    print(f"Total Cost per Flyer: ${report['print_cost_per_flyer'] + report['distribution_cost_per_flyer']:.2f}\n")
    
    print(f"Total Printing Cost: ${report['total_print_cost']:,.2f}")
    print(f"Total Distribution Cost: ${report['total_distribution_cost']:,.2f}")
    print(f"Total Campaign Cost: ${report['campaign_cost']:,.2f}\n")
    
    print("PERFORMANCE FUNNEL (DETAILED BREAKDOWN)")
    print("-" * 80)
    print(f"Step 1 - Distribution (Impressions):")
    print(f"  {num_flyers:,} flyers distributed\n")
    
    print(f"Step 2 - Responses:")
    print(f"  {num_flyers:,} flyers × {report['response_rate_percent']}% response rate")
    print(f"  = {responses:,} responses (calls, visits, QR scans)\n")
    
    print(f"Step 3 - Conversions:")
    print(f"  {responses:,} responses × {report['conversion_rate_percent']}% conversion rate")
    print(f"  = {conversions:,} purchases\n")
    
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
    print(f"Cost per Impression: ${report['cost_per_impression']:.4f}")
    print(f"Cost per Response: ${report['cost_per_response']:.2f}")
    print(f"Cost per Conversion (CPA): ${report['cost_per_conversion']:.2f}\n")
    
    print("=" * 80)
    print(f"\n💡 Summary:")
    print(f"   A flyer campaign with {num_flyers:,} flyers (${report['campaign_cost']:,.2f})")
    print(f"   generated {responses:,} responses and {conversions:,} conversions,")
    print(f"   resulting in ${report['total_sales']:,.2f} in sales ({report['roi_percent']}% ROI).\n")
    print("=" * 80)
