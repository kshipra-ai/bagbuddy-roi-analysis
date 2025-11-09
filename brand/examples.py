"""
Brand ROI Analysis - Example Usage

This script demonstrates how to use the brand ROI analysis tools
with real-world scenarios.
"""

from pathlib import Path
from roi_calculator import BrandROICalculator
from comparison import CampaignComparison
from visualization import BrandROIVisualizer


def example_1_single_campaign_analysis():
    """Example 1: Analyze a single campaign's ROI."""
    print("\n" + "="*80)
    print("EXAMPLE 1: Single Campaign Analysis")
    print("="*80 + "\n")
    
    # Sample campaign data
    campaign_data = {
        'campaign_name': 'Summer Sale 2025',
        'total_investment': 15000,
        'total_revenue': 45000,
        'new_customers': 250,
        'impressions': 500000,
        'engagements': 25000,
        'clicks': 10000,
        'conversions': 500
    }
    
    # Create calculator
    calculator = BrandROICalculator(campaign_data)
    
    # Get full report
    report = calculator.get_full_report()
    
    print(f"Campaign: {campaign_data['campaign_name']}")
    print(f"Investment: ${campaign_data['total_investment']:,}")
    print(f"Revenue: ${campaign_data['total_revenue']:,}")
    print(f"\nResults:")
    print(f"  ROI: {report['roi_percentage']}%")
    print(f"  Customer Acquisition Cost: ${report['customer_acquisition_cost']}")
    print(f"  Engagement Rate: {report['engagement_rate']}%")
    print(f"  New Customers: {report['new_customers']}")
    
    # Interpretation
    print(f"\n💡 Interpretation:")
    if report['roi_percentage'] > 200:
        print(f"   ✅ Excellent ROI! Campaign generated {report['roi_percentage']}% return.")
    elif report['roi_percentage'] > 100:
        print(f"   ✓ Good ROI at {report['roi_percentage']}%. Room for improvement.")
    else:
        print(f"   ⚠️  ROI of {report['roi_percentage']}% needs optimization.")


def example_2_compare_campaigns():
    """Example 2: Compare multiple campaigns."""
    print("\n" + "="*80)
    print("EXAMPLE 2: Campaign Comparison")
    print("="*80 + "\n")
    
    # Load campaigns from CSV
    current_dir = Path(__file__).parent
    data_file = current_dir / "sample_data.csv"
    
    if not data_file.exists():
        print("Sample data file not found. Please ensure sample_data.csv exists.")
        return
    
    # Create comparison
    comparison = CampaignComparison()
    comparison.load_from_csv(str(data_file))
    
    # Generate and print report
    print(comparison.generate_comparison_report())


def example_3_visualize_performance():
    """Example 3: Visualize campaign performance."""
    print("\n" + "="*80)
    print("EXAMPLE 3: Campaign Performance Visualization")
    print("="*80 + "\n")
    
    # Load data
    current_dir = Path(__file__).parent
    data_file = current_dir / "sample_data.csv"
    
    if not data_file.exists():
        print("Sample data file not found. Please ensure sample_data.csv exists.")
        return
    
    # Create visualizer
    viz = BrandROIVisualizer(str(data_file))
    
    # Show summary table
    print(viz.generate_summary_table())
    
    # Show insights
    print(viz.generate_performance_insights())
    
    # Show charts
    print(viz.generate_ascii_bar_chart('roi'))


def example_4_calculate_clv():
    """Example 4: Calculate Customer Lifetime Value."""
    print("\n" + "="*80)
    print("EXAMPLE 4: Customer Lifetime Value Calculation")
    print("="*80 + "\n")
    
    # Sample customer data
    average_purchase_value = 150
    purchase_frequency = 4  # purchases per year
    customer_lifespan = 3  # years
    
    clv = average_purchase_value * purchase_frequency * customer_lifespan
    
    print(f"Average Purchase Value: ${average_purchase_value}")
    print(f"Purchase Frequency: {purchase_frequency} times/year")
    print(f"Customer Lifespan: {customer_lifespan} years")
    print(f"\nCustomer Lifetime Value (CLV): ${clv}")
    
    # Compare with CAC
    cac = 60  # from campaign
    clv_to_cac_ratio = clv / cac
    
    print(f"\nCustomer Acquisition Cost: ${cac}")
    print(f"CLV to CAC Ratio: {clv_to_cac_ratio:.2f}:1")
    
    print(f"\n💡 Interpretation:")
    if clv_to_cac_ratio >= 3:
        print(f"   ✅ Excellent! CLV is {clv_to_cac_ratio:.1f}x higher than CAC.")
    elif clv_to_cac_ratio >= 2:
        print(f"   ✓ Good ratio. Aim for 3:1 or higher for optimal profitability.")
    else:
        print(f"   ⚠️  CAC is too high relative to CLV. Need to reduce acquisition costs.")


def example_5_optimize_underperforming_campaign():
    """Example 5: Identify and optimize underperforming campaign."""
    print("\n" + "="*80)
    print("EXAMPLE 5: Optimizing Underperforming Campaign")
    print("="*80 + "\n")
    
    # Underperforming campaign
    current_campaign = {
        'campaign_name': 'Q3 Brand Awareness',
        'total_investment': 12000,
        'total_revenue': 15000,
        'new_customers': 80,
        'impressions': 500000,
        'engagements': 15000,
        'clicks': 5000,
        'conversions': 120
    }
    
    calc = BrandROICalculator(current_campaign)
    current_roi = calc.calculate_roi()
    current_cac = calc.calculate_customer_acquisition_cost()
    current_cvr = (current_campaign['conversions'] / current_campaign['clicks'] * 100)
    
    print(f"Current Campaign Performance:")
    print(f"  ROI: {current_roi}%")
    print(f"  CAC: ${current_cac}")
    print(f"  Conversion Rate: {current_cvr:.2f}%")
    
    print(f"\n📊 Optimization Scenarios:\n")
    
    # Scenario 1: Improve conversion rate by 50%
    improved_conversions = int(current_campaign['conversions'] * 1.5)
    improved_customers = int(current_campaign['new_customers'] * 1.5)
    improved_revenue = current_campaign['total_revenue'] * 1.5
    
    scenario1 = current_campaign.copy()
    scenario1['conversions'] = improved_conversions
    scenario1['new_customers'] = improved_customers
    scenario1['total_revenue'] = improved_revenue
    
    calc1 = BrandROICalculator(scenario1)
    scenario1_roi = calc1.calculate_roi()
    
    print(f"Scenario 1: Improve Conversion Rate by 50%")
    print(f"  New ROI: {scenario1_roi:.2f}% (increase of {scenario1_roi - current_roi:.2f}%)")
    print(f"  Additional Revenue: ${improved_revenue - current_campaign['total_revenue']:,.2f}\n")
    
    # Scenario 2: Reduce investment by 20%
    reduced_investment = current_campaign['total_investment'] * 0.8
    scenario2 = current_campaign.copy()
    scenario2['total_investment'] = reduced_investment
    
    calc2 = BrandROICalculator(scenario2)
    scenario2_roi = calc2.calculate_roi()
    scenario2_cac = calc2.calculate_customer_acquisition_cost()
    
    print(f"Scenario 2: Reduce Investment by 20%")
    print(f"  New ROI: {scenario2_roi:.2f}% (increase of {scenario2_roi - current_roi:.2f}%)")
    print(f"  New CAC: ${scenario2_cac:.2f} (reduced by ${current_cac - scenario2_cac:.2f})")
    print(f"  Cost Savings: ${current_campaign['total_investment'] - reduced_investment:,.2f}\n")
    
    print(f"💡 Recommendation:")
    print(f"   Combine both strategies: Improve targeting (better CVR) while")
    print(f"   reducing spend on underperforming channels.")


def example_6_seasonal_campaign_planning():
    """Example 6: Plan seasonal campaign budget."""
    print("\n" + "="*80)
    print("EXAMPLE 6: Seasonal Campaign Budget Planning")
    print("="*80 + "\n")
    
    # Target metrics
    target_revenue = 100000
    target_roi = 300  # 300% or 3:1
    expected_cvr = 2.5  # 2.5%
    avg_order_value = 200
    
    # Calculate required metrics
    required_profit = target_revenue * (target_roi / (100 + target_roi))
    max_investment = target_revenue - required_profit
    required_conversions = target_revenue / avg_order_value
    required_clicks = required_conversions / (expected_cvr / 100)
    cpc = max_investment / required_clicks
    
    print(f"Campaign Goals:")
    print(f"  Target Revenue: ${target_revenue:,.2f}")
    print(f"  Target ROI: {target_roi}%")
    print(f"  Expected Conversion Rate: {expected_cvr}%")
    print(f"  Average Order Value: ${avg_order_value}\n")
    
    print(f"Required Metrics:")
    print(f"  Maximum Investment: ${max_investment:,.2f}")
    print(f"  Required Conversions: {required_conversions:.0f}")
    print(f"  Required Clicks: {required_clicks:.0f}")
    print(f"  Maximum CPC: ${cpc:.2f}\n")
    
    print(f"💡 Strategy:")
    print(f"   To achieve {target_roi}% ROI on ${target_revenue:,.0f} revenue:")
    print(f"   1. Keep total spend under ${max_investment:,.2f}")
    print(f"   2. Maintain conversion rate above {expected_cvr}%")
    print(f"   3. Keep cost per click below ${cpc:.2f}")
    print(f"   4. Target {required_conversions:.0f} conversions")


def main():
    """Run all examples."""
    examples = [
        example_1_single_campaign_analysis,
        example_2_compare_campaigns,
        example_3_visualize_performance,
        example_4_calculate_clv,
        example_5_optimize_underperforming_campaign,
        example_6_seasonal_campaign_planning
    ]
    
    print("\n" + "🎯 BRAND ROI ANALYSIS - EXAMPLE SCENARIOS" + "\n")
    
    for i, example in enumerate(examples, 1):
        try:
            example()
        except Exception as e:
            print(f"\n⚠️  Error in example {i}: {e}")
        
        if i < len(examples):
            input("\n\nPress Enter to continue to next example...")
    
    print("\n" + "="*80)
    print("✅ All examples completed!")
    print("="*80 + "\n")


if __name__ == "__main__":
    main()
