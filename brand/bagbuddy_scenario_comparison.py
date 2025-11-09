"""
BagBuddy ROI Scenario Comparison

This module compares three scenarios for BagBuddy platform campaigns:
1. Conservative (baseline): No feature boost assumptions
2. Moderate (realistic boost): Platform features provide modest improvements
3. Optimistic (best case): Maximum realistic impact from platform features

BAGBUDDY UNIQUE FEATURES:
1. Platform-wide redemption: Users can save points and redeem with ANY brand
2. No single-brand lock-in: Flexibility increases perceived value
3. Environmental impact: For every $1000 in earned points, 1 tree planted

RESEARCH-BASED ASSUMPTIONS:

Conservative Scenario (Baseline):
- Scan rate: 2.0% (industry QR code benchmark - Statista 2024)
- Conversion rate: 20% (realistic for small purchases)
- No assumptions about feature benefits

Moderate Scenario (Platform Features Boost):
- Scan rate: 2.5% (+0.5% boost)
  * Research: Environmental messaging increases engagement 10-30% (Nielsen 2023)
  * Conservative estimate: +25% relative increase in scans
  
- Conversion rate: 25% (+5% boost)
  * Research: Coalition loyalty programs see 15-30% higher redemption (Bond Loyalty 2023)
  * Conservative estimate: +25% relative increase in conversion

Optimistic Scenario (Best Case):
- Scan rate: 3.0% (+1.0% boost)
  * Research: Top eco-conscious campaigns can achieve 30% engagement lift
  * Optimistic estimate: +50% relative increase in scans
  
- Conversion rate: 30% (+10% boost)
  * Research: Best coalition programs achieve 30%+ redemption rates
  * Optimistic estimate: +50% relative increase in conversion

NOTE: All scenarios use same cost structure and average order value.
"""

from typing import Dict, Any, List
from roi_calculator import BrandROICalculator


class BagBuddyScenarioComparison:
    """Compare different ROI scenarios for BagBuddy campaigns."""
    
    # Scenario configurations
    SCENARIOS = {
        'conservative': {
            'name': 'Conservative (Baseline)',
            'description': 'No feature boost - pure industry benchmarks',
            'scan_rate': 2.0,
            'conversion_rate': 20.0,
            'color': '🔴'
        },
        'moderate': {
            'name': 'Moderate (Platform Boost)',
            'description': 'Realistic impact from BagBuddy features',
            'scan_rate': 2.5,
            'conversion_rate': 25.0,
            'color': '🟡'
        },
        'optimistic': {
            'name': 'Optimistic (Best Case)',
            'description': 'Maximum realistic feature impact',
            'scan_rate': 3.0,
            'conversion_rate': 30.0,
            'color': '🟢'
        }
    }
    
    def __init__(self, base_campaign_data: Dict[str, Any]):
        """
        Initialize with base campaign parameters.
        
        Args:
            base_campaign_data: Dictionary with campaign basics
                Required keys:
                - num_quarters: Number of quarters
                - avg_revenue_per_conversion: Average sale value
                - impressions_per_bag: Physical impressions per bag
                - trees_planted: Optional environmental metric
        """
        self.base_campaign_data = base_campaign_data
    
    def run_scenario(self, scenario_name: str) -> Dict[str, Any]:
        """
        Run ROI calculation for a specific scenario.
        
        Args:
            scenario_name: One of 'conservative', 'moderate', 'optimistic'
            
        Returns:
            Full ROI report dictionary
        """
        if scenario_name not in self.SCENARIOS:
            raise ValueError(f"Unknown scenario: {scenario_name}")
        
        scenario = self.SCENARIOS[scenario_name]
        
        # Build campaign data with scenario-specific rates
        campaign_data = self.base_campaign_data.copy()
        campaign_data['scan_rate'] = scenario['scan_rate']
        campaign_data['conversion_rate'] = scenario['conversion_rate']
        
        # Calculate using existing ROI calculator
        calculator = BrandROICalculator(campaign_data)
        report = calculator.get_full_report()
        
        # Add scenario metadata
        report['scenario_name'] = scenario['name']
        report['scenario_description'] = scenario['description']
        report['scenario_color'] = scenario['color']
        
        return report
    
    def compare_all_scenarios(self) -> Dict[str, Dict[str, Any]]:
        """
        Run all three scenarios and return comparison.
        
        Returns:
            Dictionary with all scenario results
        """
        results = {}
        for scenario_name in ['conservative', 'moderate', 'optimistic']:
            results[scenario_name] = self.run_scenario(scenario_name)
        
        return results
    
    def print_comparison_table(self):
        """Print a formatted comparison table of all scenarios."""
        results = self.compare_all_scenarios()
        
        print("\n" + "=" * 100)
        print("BAGBUDDY ROI SCENARIO COMPARISON")
        print("=" * 100 + "\n")
        
        print("PLATFORM FEATURES:")
        print("  ✓ Platform-wide redemption (save points, redeem with ANY brand)")
        print("  ✓ No single-brand lock-in (flexibility increases value)")
        print("  ✓ Environmental impact ($1000 earned = 1 tree planted)\n")
        
        # Header
        print("-" * 100)
        print(f"{'METRIC':<30} {'CONSERVATIVE':<23} {'MODERATE':<23} {'OPTIMISTIC':<23}")
        print("-" * 100)
        
        # Extract data
        cons = results['conservative']
        mod = results['moderate']
        opt = results['optimistic']
        
        # Scenario descriptions
        print(f"{'Scenario Description':<30} {'Baseline - No Boost':<23} {'Platform Features':<23} {'Best Case Scenario':<23}")
        print()
        
        # Input assumptions
        print("INPUT ASSUMPTIONS:")
        print(f"{'Scan Rate':<30} {cons['scan_rate_percent']:.1f}%{'':<18} {mod['scan_rate_percent']:.1f}%{'':<18} {opt['scan_rate_percent']:.1f}%")
        print(f"{'Conversion Rate':<30} {cons['conversion_rate_percent']:.1f}%{'':<18} {mod['conversion_rate_percent']:.1f}%{'':<18} {opt['conversion_rate_percent']:.1f}%")
        print()
        
        # Key metrics
        print("CAMPAIGN COSTS:")
        print(f"{'Campaign Cost':<30} ${cons['campaign_cost']:>8,.2f}{'':<13} ${mod['campaign_cost']:>8,.2f}{'':<13} ${opt['campaign_cost']:>8,.2f}")
        print()
        
        print("PERFORMANCE:")
        print(f"{'Total Bags':<30} {cons['num_bags_distributed']:>8,}{'':<13} {mod['num_bags_distributed']:>8,}{'':<13} {opt['num_bags_distributed']:>8,}")
        print(f"{'Total Impressions':<30} {cons['total_impressions']:>8,}{'':<13} {mod['total_impressions']:>8,}{'':<13} {opt['total_impressions']:>8,}")
        print(f"{'QR Code Scans':<30} {cons['engagements_scans']:>8,}{'':<13} {mod['engagements_scans']:>8,}{'':<13} {opt['engagements_scans']:>8,}")
        print(f"{'Conversions':<30} {cons['conversions_redemptions']:>8,}{'':<13} {mod['conversions_redemptions']:>8,}{'':<13} {opt['conversions_redemptions']:>8,}")
        print()
        
        print("REVENUE:")
        print(f"{'Avg Revenue/Conversion':<30} ${cons['avg_revenue_per_conversion']:>8,.2f}{'':<13} ${mod['avg_revenue_per_conversion']:>8,.2f}{'':<13} ${opt['avg_revenue_per_conversion']:>8,.2f}")
        print(f"{'Total Sales':<30} ${cons['total_sales_generated']:>8,.2f}{'':<13} ${mod['total_sales_generated']:>8,.2f}{'':<13} ${opt['total_sales_generated']:>8,.2f}")
        print()
        
        # ROI - with color indicators
        print("ROI METRICS:")
        roi_cons = f"{cons['roi_percent']:>8.2f}%"
        roi_mod = f"{mod['roi_percent']:>8.2f}%"
        roi_opt = f"{opt['roi_percent']:>8.2f}%"
        
        print(f"{'ROI':<30} {self.SCENARIOS['conservative']['color']} {roi_cons:<20} {self.SCENARIOS['moderate']['color']} {roi_mod:<20} {self.SCENARIOS['optimistic']['color']} {roi_opt}")
        print(f"{'Cost per Impression':<30} ${cons['cost_per_impression']:>8.4f}{'':<13} ${mod['cost_per_impression']:>8.4f}{'':<13} ${opt['cost_per_impression']:>8.4f}")
        print(f"{'Cost per Engagement':<30} ${cons['cost_per_engagement']:>8.2f}{'':<13} ${mod['cost_per_engagement']:>8.2f}{'':<13} ${opt['cost_per_engagement']:>8.2f}")
        print(f"{'Cost per Conversion':<30} ${cons['cost_per_conversion']:>8.2f}{'':<13} ${mod['cost_per_conversion']:>8.2f}{'':<13} ${opt['cost_per_conversion']:>8.2f}")
        print()
        
        # Environmental impact
        print("ENVIRONMENTAL IMPACT:")
        trees_cons = cons.get('trees_planted', 0) + (cons['total_sales_generated'] / 1000)
        trees_mod = mod.get('trees_planted', 0) + (mod['total_sales_generated'] / 1000)
        trees_opt = opt.get('trees_planted', 0) + (opt['total_sales_generated'] / 1000)
        
        print(f"{'Trees Planted (sales/1000)':<30} {trees_cons:>8.2f}{'':<13} {trees_mod:>8.2f}{'':<13} {trees_opt:>8.2f}")
        print()
        
        print("-" * 100)
        
        # Summary insights
        print("\n💡 KEY INSIGHTS:")
        print()
        print("1. BASELINE (Conservative): Pure industry benchmarks, no feature advantage")
        print(f"   • ROI: {cons['roi_percent']:.2f}% - {self._roi_verdict(cons['roi_percent'])}")
        print()
        print("2. PLATFORM BOOST (Moderate): Realistic impact from BagBuddy unique features")
        print(f"   • ROI: {mod['roi_percent']:.2f}% - {self._roi_verdict(mod['roi_percent'])}")
        print(f"   • Improvement: {mod['roi_percent'] - cons['roi_percent']:+.2f} percentage points vs baseline")
        print()
        print("3. BEST CASE (Optimistic): Maximum realistic benefit from all features")
        print(f"   • ROI: {opt['roi_percent']:.2f}% - {self._roi_verdict(opt['roi_percent'])}")
        print(f"   • Improvement: {opt['roi_percent'] - cons['roi_percent']:+.2f} percentage points vs baseline")
        print()
        
        # Feature impact analysis
        scan_boost = ((mod['scan_rate_percent'] - cons['scan_rate_percent']) / cons['scan_rate_percent']) * 100
        conv_boost = ((mod['conversion_rate_percent'] - cons['conversion_rate_percent']) / cons['conversion_rate_percent']) * 100
        
        print("FEATURE IMPACT (Moderate Scenario):")
        print(f"  • Environmental messaging → +{scan_boost:.0f}% scan rate boost")
        print(f"  • Platform flexibility → +{conv_boost:.0f}% conversion rate boost")
        print(f"  • Combined effect → {mod['roi_percent'] - cons['roi_percent']:+.2f}pp ROI improvement")
        print()
        
        print("=" * 100 + "\n")
    
    def _roi_verdict(self, roi: float) -> str:
        """Return a verdict string based on ROI percentage."""
        if roi < -50:
            return "SEVERE LOSS"
        elif roi < 0:
            return "UNPROFITABLE"
        elif roi < 10:
            return "MARGINALLY PROFITABLE"
        elif roi < 25:
            return "MODEST PROFIT"
        elif roi < 50:
            return "GOOD PROFIT"
        else:
            return "EXCELLENT PROFIT"
    
    def get_recommendation(self) -> str:
        """
        Provide recommendation based on scenario comparison.
        
        Returns:
            Recommendation text
        """
        results = self.compare_all_scenarios()
        mod = results['moderate']
        
        if mod['roi_percent'] < 0:
            return """
RECOMMENDATION: IMPROVE AVERAGE ORDER VALUE

Even with platform features, ROI is negative with current $25 avg order value.
Action items:
1. Target brands with higher average transaction values ($40+)
2. Encourage bundle purchases or minimum order requirements
3. Focus on repeat customer lifetime value, not single transaction ROI
4. Consider subscription or membership models
"""
        elif mod['roi_percent'] < 15:
            return """
RECOMMENDATION: OPTIMIZE ENGAGEMENT

ROI is marginally positive. Focus on maximizing platform features:
1. Emphasize environmental impact in marketing (tree planting)
2. Promote platform-wide redemption flexibility
3. A/B test QR code placement and call-to-action messaging
4. Improve landing page conversion optimization
"""
        else:
            return """
RECOMMENDATION: SCALE CAMPAIGN

ROI is positive and healthy. Focus on growth:
1. Increase bag distribution volume
2. Expand to more brand partnerships
3. Invest in user education about platform benefits
4. Track and optimize based on actual performance data
"""


if __name__ == "__main__":
    # Example campaign
    base_campaign = {
        'num_quarters': 1,
        'avg_revenue_per_conversion': 25,  # $25 average sale
        'impressions_per_bag': 5,  # 5 impressions per bag per brand
        'trees_planted': 0  # Will calculate based on sales
    }
    
    comparison = BagBuddyScenarioComparison(base_campaign)
    comparison.print_comparison_table()
    
    print(comparison.get_recommendation())
