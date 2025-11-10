"""
Comprehensive Demo of Kshipra Commerce-Reward Model

Demonstrates:
1. Basic calculations
2. Scenario comparison
3. Investor metrics
4. Sensitivity analysis
"""

from commerce_reward_calculator import CommerceRewardCalculator

def demo_basic_model():
    """Demonstrate basic model calculations."""
    print("\n" + "="*80)
    print("DEMO 1: BASIC MODEL CALCULATIONS")
    print("="*80)
    
    calc = CommerceRewardCalculator()
    
    # Monthly summary
    summary = calc.calculate_monthly_summary()
    print("\n📊 Monthly Business Summary:")
    print(f"   Bags Sold: {summary['bags_sold']:,}")
    print(f"   Total Revenue: ${summary['total_revenue']:,.2f}")
    print(f"   Total Margin: ${summary['total_margin']:,.2f} ({summary['margin_percentage']:.1f}%)")
    print(f"   Revenue per Bag: ${summary['avg_revenue_per_bag']:.2f}")
    
    # Per user
    per_user = calc.calculate_revenue_per_user()
    print("\n👤 Per User Economics:")
    print(f"   Brand Spend: ${per_user['brand_spend_per_user']:.2f}/month")
    print(f"   User Rewards: ${per_user['total_user_rewards']:.2f}/month")
    print(f"   Kshipra Margin: ${per_user['kshipra_margin_per_user']:.2f}/month")
    
    # Per bag
    per_bag = calc.calculate_revenue_per_bag()
    print("\n🛍️  Per Bag Economics:")
    print(f"   Bag Retail Price: ${per_bag['bag_retail_revenue']:.2f}")
    print(f"   User Earnings: ${per_bag['user_earnings_per_bag']:.2f}")
    print(f"   User Net Cost: ${per_bag['user_net_cost']:.2f}")
    print(f"   → Bags are {'FREE' if per_bag['user_net_cost'] <= 0 else 'NOT free'} after watching {calc.assumptions['avg_ads_to_recover_bag']} ads!")

def demo_scenarios():
    """Demonstrate scenario comparison."""
    print("\n" + "="*80)
    print("DEMO 2: SCENARIO COMPARISON")
    print("="*80)
    
    calc_base = CommerceRewardCalculator()
    calc_best = calc_base.generate_scenario('best')
    calc_worst = calc_base.generate_scenario('worst')
    
    scenarios = [
        ('Worst Case', calc_worst),
        ('Base Case', calc_base),
        ('Best Case', calc_best)
    ]
    
    print("\n{:<15} {:<15} {:<15} {:<15}".format("Scenario", "Revenue", "Margin", "Store Value"))
    print("-" * 65)
    
    for name, calc in scenarios:
        summary = calc.calculate_monthly_summary()
        store = calc.calculate_store_value()
        print("{:<15} ${:<14,.0f} ${:<14,.0f} ${:<14,.0f}".format(
            name,
            summary['total_revenue'],
            summary['total_margin'],
            store['total_monthly_store_value']
        ))
    
    print("\n📈 Key Assumptions by Scenario:")
    print("{:<15} {:<10} {:<12} {:<15}".format("Scenario", "Ad Views", "Platinum %", "Redemption %"))
    print("-" * 55)
    for name, calc in scenarios:
        print("{:<15} {:<10} {:<12} {:<15}".format(
            name,
            calc.assumptions['avg_monthly_ad_views'],
            f"{calc.assumptions['pct_platinum']}%",
            f"{calc.assumptions['reward_redemption_rate']}%"
        ))

def demo_investor_metrics():
    """Demonstrate investor-specific metrics."""
    print("\n" + "="*80)
    print("DEMO 3: INVESTOR METRICS")
    print("="*80)
    
    calc = CommerceRewardCalculator()
    investor = calc.calculate_investor_metrics()
    
    print("\n💼 Unit Economics:")
    print(f"   Customer Acquisition Cost (CAC): ${investor['cac']:.2f}")
    print(f"   Customer Lifetime Value (LTV): ${investor['ltv']:.2f}")
    print(f"      - From Bag Sales: ${investor['bag_ltv']:.2f}")
    print(f"      - From Ad Revenue: ${investor['ad_ltv']:.2f}")
    
    print("\n📊 Key Ratios:")
    print(f"   LTV:CAC Ratio: {investor['ltv_cac_ratio']:.1f}x")
    print(f"      → {'✅ EXCELLENT' if investor['ltv_cac_ratio'] > 3 else '⚠️  NEEDS IMPROVEMENT'} (Target: >3x)")
    
    print(f"\n   Payback Period: {investor['payback_months']:.1f} months")
    print(f"      → {'✅ EXCELLENT' if investor['payback_months'] < 6 else '⚠️  NEEDS IMPROVEMENT'} (Target: <6 months)")
    
    print(f"\n   Contribution Margin: ${investor['contribution_margin_per_user']:.2f}/user ({investor['contribution_margin_pct']:.1f}%)")
    
    # Annual projections
    summary = calc.calculate_monthly_summary()
    annual_revenue = summary['total_revenue'] * 12
    annual_margin = summary['total_margin'] * 12
    
    print("\n📅 Annual Projections (12 months):")
    print(f"   Annual Revenue: ${annual_revenue:,.0f}")
    print(f"   Annual Margin: ${annual_margin:,.0f}")

def demo_sensitivity_analysis():
    """Demonstrate sensitivity analysis."""
    print("\n" + "="*80)
    print("DEMO 4: SENSITIVITY ANALYSIS")
    print("="*80)
    
    calc = CommerceRewardCalculator()
    
    # Sensitivity to ad views
    print("\n🔍 Sensitivity: Monthly Ad Views per User")
    print("{:<12} {:<15} {:<15} {:<12}".format("Ad Views", "Revenue", "Margin", "Margin %"))
    print("-" * 60)
    
    ad_view_values = [6, 9, 12, 15, 20]
    results = calc.calculate_sensitivity_analysis('avg_monthly_ad_views', ad_view_values)
    
    for result in results:
        print("{:<12} ${:<14,.0f} ${:<14,.0f} {:<12.1f}%".format(
            result['value'],
            result['total_revenue'],
            result['total_margin'],
            result['margin_pct']
        ))
    
    # Sensitivity to bag price
    print("\n🔍 Sensitivity: Bag Retail Price")
    print("{:<12} {:<15} {:<15} {:<15}".format("Bag Price", "Revenue", "Margin", "User Net Cost"))
    print("-" * 65)
    
    bag_prices = [0.30, 0.40, 0.50, 0.60]
    for price in bag_prices:
        calc.assumptions['bag_retail_price'] = price
        summary = calc.calculate_monthly_summary()
        per_bag = calc.calculate_revenue_per_bag()
        
        print("${:<11.2f} ${:<14,.0f} ${:<14,.0f} ${:<15.2f}".format(
            price,
            summary['total_revenue'],
            summary['total_margin'],
            per_bag['user_net_cost']
        ))
    
    print("\n💡 Insight: Higher bag prices increase margin but raise user net cost.")
    print("   Optimal price balances margin with keeping bags 'free' for users.")

def demo_tier_analysis():
    """Demonstrate tier-based economics."""
    print("\n" + "="*80)
    print("DEMO 5: TIER-BASED ECONOMICS")
    print("="*80)
    
    calc = CommerceRewardCalculator()
    tier_data = calc.calculate_tier_economics()
    
    print("\n🏆 User Tier Breakdown:")
    print("{:<12} {:<18} {:<15} {:<15}".format("Tier", "Cash Limit/Month", "Max Views", "User %"))
    print("-" * 65)
    
    for tier_name, tier_info in tier_data['tiers'].items():
        print("{:<12} ${:<17.2f} {:<15} {:<15}%".format(
            tier_name,
            tier_info['monthly_cash_limit'],
            tier_info['max_monthly_views'],
            tier_info['user_percentage']
        ))
    
    print(f"\n📊 Weighted Average Cash Limit: ${tier_data['weighted_avg_cash_limit']:.2f}/month")
    print(f"   (Based on user distribution across tiers)")

def main():
    """Run all demonstrations."""
    print("\n" + "="*80)
    print("KSHIPRA COMMERCE-REWARD MODEL - COMPREHENSIVE DEMO")
    print("="*80)
    
    demo_basic_model()
    demo_scenarios()
    demo_investor_metrics()
    demo_sensitivity_analysis()
    demo_tier_analysis()
    
    print("\n" + "="*80)
    print("📋 SUMMARY")
    print("="*80)
    print("\n✅ Model Features Demonstrated:")
    print("   1. Basic revenue & margin calculations")
    print("   2. Best/Base/Worst case scenarios")
    print("   3. Investor metrics (LTV, CAC, payback)")
    print("   4. Sensitivity analysis (ad views, pricing)")
    print("   5. Tier-based user economics")
    
    print("\n📁 Generated Files:")
    print("   • commerce_reward_calculator.py - Python calculator")
    print("   • export_commerce_reward_excel.py - Excel generator")
    print("   • kshipra_commerce_reward_model.xlsx - Interactive Excel workbook")
    
    print("\n💡 Next Steps:")
    print("   1. Open kshipra_commerce_reward_model.xlsx")
    print("   2. Go to Assumptions sheet")
    print("   3. Edit YELLOW cells to test different scenarios")
    print("   4. Watch all metrics update automatically!")
    
    print("\n" + "="*80 + "\n")

if __name__ == "__main__":
    main()
