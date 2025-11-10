from commerce_reward_calculator import CommerceRewardCalculator

c = CommerceRewardCalculator()
growth = c.calculate_business_growth_metrics()

print("BUSINESS GROWTH PROJECTIONS:")
print(f"\nBag Sales:")
print(f"  Initial: {growth['initial_bags']:,}")
print(f"  Q1: {growth['bags_q1']:,.0f} | Q2: {growth['bags_q2']:,.0f} | Q3: {growth['bags_q3']:,.0f} | Q4: {growth['bags_q4']:,.0f}")
print(f"  Growth Rate: {growth['bag_growth_rate_quarterly']}% per quarter")
print(f"  Revenue Lift: ${growth['bag_revenue_lift']:,.0f} (+{growth['bag_revenue_lift_pct']:.1f}%)")

print(f"\nBrand Partnerships:")
print(f"  Initial: {growth['initial_brands']}")
print(f"  Q1: {growth['brands_q1']:.0f} | Q2: {growth['brands_q2']:.0f} | Q3: {growth['brands_q3']:.0f} | Q4: {growth['brands_q4']:.0f}")
print(f"  Growth Rate: {growth['brand_growth_rate_quarterly']}% per quarter")
print(f"  Fill Rate: {growth['base_fill_rate']:.1f}% -> {growth['year_end_fill_rate']:.1f}%")
print(f"  Revenue Lift: ${growth['ad_revenue_lift']:,.0f} (+{growth['ad_revenue_lift_pct']:.1f}%)")

print(f"\nCombined Impact:")
print(f"  Base Annual Revenue: ${growth['base_annual_revenue']:,.0f}")
print(f"  Year 1 Actual Revenue: ${growth['total_annual_revenue']:,.0f}")
print(f"  Total Lift: ${growth['total_revenue_lift']:,.0f} (+{growth['total_revenue_lift_pct']:.1f}%)")
