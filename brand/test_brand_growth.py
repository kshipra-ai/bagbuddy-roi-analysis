from commerce_reward_calculator import CommerceRewardCalculator

c = CommerceRewardCalculator()
bg = c.calculate_brand_growth_metrics()

print("BRAND PARTNERSHIP GROWTH:")
print(f"Brands: Q1={bg['brands_q1']:.0f}, Q2={bg['brands_q2']:.0f}, Q3={bg['brands_q3']:.0f}, Q4={bg['brands_q4']:.0f}")
print(f"Avg Brands Year 1: {bg['avg_brands_year1']:.1f}")
print(f"Fill Rate: {bg['base_fill_rate']:.1f}% -> {bg['year_end_fill_rate']:.1f}% (avg {bg['avg_fill_rate']:.1f}%)")
print(f"Revenue Lift: ${bg['revenue_lift_from_growth']:,.0f} ({bg['revenue_lift_pct']:.1f}%)")
print(f"\nAnnual Ad Revenue:")
print(f"  Base (no growth): ${bg['base_annual_ad_revenue']:,.0f}")
print(f"  With growth: ${bg['total_annual_ad_revenue']:,.0f}")
