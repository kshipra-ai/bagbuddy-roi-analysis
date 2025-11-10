"""
Kshipra AI Commerce-Reward Business Model ROI Calculator

New business model where users buy bags and earn rewards by watching ads.
Includes tier-based cash credit limits and comprehensive financial modeling.

Author: Kshipra AI
Version: 1.0
Date: November 2025
"""

class CommerceRewardCalculator:
    """
    Calculate ROI for Kshipra's commerce-reward business model.
    
    Business Model:
    - Users purchase bags at retail price
    - Users watch ads to earn cash credits + reward points
    - Cash credits unlock based on tier (Bronze/Silver/Gold/Platinum)
    - Rewards redeemable at local stores
    - Brands pay per verified ad view (CPV model)
    """
    
    def __init__(self, assumptions=None):
        """Initialize calculator with default or custom assumptions."""
        self.assumptions = assumptions or self.get_default_assumptions()
    
    @staticmethod
    def get_default_assumptions():
        """Return default business model assumptions."""
        return {
            # Pricing
            'bag_retail_price': 0.40,           # Price user pays per bag
            'bags_sold_per_month': 10000,       # Monthly bag sales volume (starting)
            
            # Business Growth
            'bag_sales_growth_quarterly': 15,   # % increase in bag sales each quarter
            'initial_brands_enrolled': 10,      # Starting number of brand partners
            'brand_growth_rate_quarterly': 25,  # % increase in brands each quarter
            
            # Ad Economics
            'cpv_brand_pays': 0.20,             # Cost per verified ad view (brand pays)
            'cash_credit_per_view': 0.06,       # Cash credit given to user per ad
            'reward_points_per_view': 0.06,     # Reward points value per ad
            'kshipra_margin_per_view': 0.08,    # Kshipra margin per ad view
            
            # User Engagement
            'active_user_rate': 60,             # % of buyers who actually watch ads
            
            # User Behavior
            'avg_ads_to_recover_bag': 3,        # Avg ads watched to recover bag cost
            'avg_monthly_ad_views': 12,         # Average ad views per active user/month
            
            # Tier Limits (monthly cash credit unlock caps)
            'bronze_cash_limit_bags': 1,        # Bronze: 1 bag value/month
            'silver_cash_limit_bags': 3,        # Silver: 3 bag values/month
            'gold_cash_limit_bags': 7,          # Gold: 7 bag values/month
            'platinum_daily_cap': 1.00,         # Platinum: $1/day max
            
            # User Distribution by Tier (%)
            'pct_bronze': 60,                   # 60% of users in Bronze
            'pct_silver': 25,                   # 25% in Silver
            'pct_gold': 12,                     # 12% in Gold
            'pct_platinum': 3,                  # 3% in Platinum
            
            # Store & Redemption
            'reward_redemption_rate': 75,       # % of rewards redeemed at stores
            'repeat_visit_increase': 15,        # % increase in repeat visits
            'basket_size_uplift': 8,            # % basket size increase
            'avg_basket_value': 25.00,          # Average store basket value
            
            # Brand Benchmarks
            'meta_cpm': 10.00,                  # Meta CPM benchmark
            'tiktok_cpm': 8.50,                 # TikTok CPM benchmark
            'industry_avg_ctr': 1.0,            # Industry average CTR %
        }
    
    def calculate_revenue_per_user(self):
        """Calculate monthly revenue metrics per active user."""
        a = self.assumptions
        
        # Monthly ad views per user
        monthly_views = a['avg_monthly_ad_views']
        
        # Brand spend (revenue to Kshipra)
        brand_spend = monthly_views * a['cpv_brand_pays']
        
        # User rewards given
        cash_credit_given = monthly_views * a['cash_credit_per_view']
        reward_points_given = monthly_views * a['reward_points_per_view']
        total_user_rewards = cash_credit_given + reward_points_given
        
        # Kshipra margin
        kshipra_margin = monthly_views * a['kshipra_margin_per_view']
        
        # Verify economics balance
        total_per_view = a['cash_credit_per_view'] + a['reward_points_per_view'] + a['kshipra_margin_per_view']
        economics_check = abs(total_per_view - a['cpv_brand_pays']) < 0.01
        
        return {
            'monthly_ad_views': monthly_views,
            'brand_spend_per_user': brand_spend,
            'cash_credit_given': cash_credit_given,
            'reward_points_given': reward_points_given,
            'total_user_rewards': total_user_rewards,
            'kshipra_margin_per_user': kshipra_margin,
            'economics_balanced': economics_check,
            'margin_percentage': (kshipra_margin / brand_spend * 100) if brand_spend > 0 else 0
        }
    
    def calculate_revenue_per_bag(self):
        """Calculate revenue metrics attributed to each bag sold."""
        a = self.assumptions
        
        # Ad views attributed per bag (based on recovery behavior)
        avg_views_per_bag = a['avg_ads_to_recover_bag']
        
        # Revenue per bag
        brand_revenue_per_bag = avg_views_per_bag * a['cpv_brand_pays']
        kshipra_margin_per_bag = avg_views_per_bag * a['kshipra_margin_per_view']
        
        # User value per bag
        user_earnings_per_bag = avg_views_per_bag * (a['cash_credit_per_view'] + a['reward_points_per_view'])
        
        # Bag economics
        bag_retail_revenue = a['bag_retail_price']
        total_revenue_per_bag = bag_retail_revenue + brand_revenue_per_bag
        net_margin_per_bag = bag_retail_revenue + kshipra_margin_per_bag
        
        return {
            'avg_ad_views_per_bag': avg_views_per_bag,
            'bag_retail_revenue': bag_retail_revenue,
            'brand_revenue_per_bag': brand_revenue_per_bag,
            'total_revenue_per_bag': total_revenue_per_bag,
            'kshipra_margin_per_bag': kshipra_margin_per_bag,
            'net_margin_per_bag': net_margin_per_bag,
            'user_earnings_per_bag': user_earnings_per_bag,
            'user_net_cost': bag_retail_revenue - user_earnings_per_bag
        }
    
    def calculate_tier_economics(self):
        """Calculate cash credit limits and economics for each user tier."""
        a = self.assumptions
        
        tiers = {
            'Bronze': {
                'monthly_cash_limit': a['bronze_cash_limit_bags'] * a['bag_retail_price'],
                'user_percentage': a['pct_bronze'],
                'max_monthly_views': int((a['bronze_cash_limit_bags'] * a['bag_retail_price']) / a['cash_credit_per_view']) if a['cash_credit_per_view'] > 0 else 0
            },
            'Silver': {
                'monthly_cash_limit': a['silver_cash_limit_bags'] * a['bag_retail_price'],
                'user_percentage': a['pct_silver'],
                'max_monthly_views': int((a['silver_cash_limit_bags'] * a['bag_retail_price']) / a['cash_credit_per_view']) if a['cash_credit_per_view'] > 0 else 0
            },
            'Gold': {
                'monthly_cash_limit': a['gold_cash_limit_bags'] * a['bag_retail_price'],
                'user_percentage': a['pct_gold'],
                'max_monthly_views': int((a['gold_cash_limit_bags'] * a['bag_retail_price']) / a['cash_credit_per_view']) if a['cash_credit_per_view'] > 0 else 0
            },
            'Platinum': {
                'monthly_cash_limit': a['platinum_daily_cap'] * 30,  # 30 days
                'user_percentage': a['pct_platinum'],
                'max_monthly_views': int((a['platinum_daily_cap'] * 30) / a['cash_credit_per_view']) if a['cash_credit_per_view'] > 0 else 0
            }
        }
        
        # Calculate weighted average limits
        total_pct = a['pct_bronze'] + a['pct_silver'] + a['pct_gold'] + a['pct_platinum']
        weighted_cash_limit = sum(
            tier['monthly_cash_limit'] * tier['user_percentage'] / 100
            for tier in tiers.values()
        )
        
        return {
            'tiers': tiers,
            'weighted_avg_cash_limit': weighted_cash_limit,
            'total_percentage': total_pct
        }
    
    def calculate_store_value(self):
        """Calculate value delivered to local stores through reward redemptions."""
        a = self.assumptions
        
        # Total bags sold
        total_bags_sold = a['bags_sold_per_month']
        
        # Active users who actually watch ads
        active_users = total_bags_sold * (a['active_user_rate'] / 100)
        
        # Total rewards generated (only active users generate rewards)
        per_user = self.calculate_revenue_per_user()
        total_rewards_generated = active_users * per_user['total_user_rewards']
        
        # Rewards redeemed at stores
        reward_redemption_value = total_rewards_generated * (a['reward_redemption_rate'] / 100)
        
        # Store traffic impact
        estimated_transactions = active_users * (a['reward_redemption_rate'] / 100)
        repeat_visits_increase = estimated_transactions * (a['repeat_visit_increase'] / 100)
        
        # Basket value impact
        baseline_basket_value = estimated_transactions * a['avg_basket_value']
        uplifted_basket_value = baseline_basket_value * (1 + a['basket_size_uplift'] / 100)
        incremental_basket_value = uplifted_basket_value - baseline_basket_value
        
        # Total store value
        total_store_value = reward_redemption_value + incremental_basket_value
        
        return {
            'total_rewards_generated': total_rewards_generated,
            'reward_redemption_value': reward_redemption_value,
            'estimated_transactions': estimated_transactions,
            'repeat_visits_increase': repeat_visits_increase,
            'baseline_basket_value': baseline_basket_value,
            'uplifted_basket_value': uplifted_basket_value,
            'incremental_basket_value': incremental_basket_value,
            'total_monthly_store_value': total_store_value
        }
    
    def calculate_brand_roi(self):
        """Calculate ROI metrics for brand advertisers."""
        a = self.assumptions
        
        # Cost per engaged customer (CPV model)
        cost_per_engaged_customer = a['cpv_brand_pays']
        
        # Compare to Meta/TikTok CPM
        # CPM = Cost per 1000 impressions
        # Effective CPM for BagBuddy (all views are verified/engaged)
        bagbuddy_cpm = a['cpv_brand_pays'] * 1000
        
        # Traditional platform effective cost (accounting for CTR)
        meta_effective_cost = (a['meta_cpm'] / 1000) / (a['industry_avg_ctr'] / 100) if a['industry_avg_ctr'] > 0 else 0
        tiktok_effective_cost = (a['tiktok_cpm'] / 1000) / (a['industry_avg_ctr'] / 100) if a['industry_avg_ctr'] > 0 else 0
        
        # Cost per basket influenced (using redemption data)
        store_metrics = self.calculate_store_value()
        total_influenced_transactions = store_metrics['estimated_transactions']
        
        # Calculate total brand spend (only active users watch ads)
        active_users = a['bags_sold_per_month'] * (a['active_user_rate'] / 100)
        total_brand_spend = active_users * a['avg_monthly_ad_views'] * a['cpv_brand_pays']
        
        cost_per_basket_influenced = total_brand_spend / total_influenced_transactions if total_influenced_transactions > 0 else 0
        
        # Value comparison
        bagbuddy_vs_meta = ((meta_effective_cost - cost_per_engaged_customer) / meta_effective_cost * 100) if meta_effective_cost > 0 else 0
        bagbuddy_vs_tiktok = ((tiktok_effective_cost - cost_per_engaged_customer) / tiktok_effective_cost * 100) if tiktok_effective_cost > 0 else 0
        
        return {
            'cost_per_engaged_customer': cost_per_engaged_customer,
            'bagbuddy_cpm': bagbuddy_cpm,
            'meta_cpm': a['meta_cpm'],
            'tiktok_cpm': a['tiktok_cpm'],
            'meta_effective_cost_per_engagement': meta_effective_cost,
            'tiktok_effective_cost_per_engagement': tiktok_effective_cost,
            'cost_per_basket_influenced': cost_per_basket_influenced,
            'bagbuddy_savings_vs_meta_pct': bagbuddy_vs_meta,
            'bagbuddy_savings_vs_tiktok_pct': bagbuddy_vs_tiktok,
            'total_monthly_brand_spend': total_brand_spend
        }
    
    def calculate_monthly_summary(self):
        """Calculate comprehensive monthly business metrics."""
        a = self.assumptions
        
        # Revenue streams
        bags_sold = a['bags_sold_per_month']
        bag_revenue = bags_sold * a['bag_retail_price']
        
        # Active users (only they watch ads)
        active_users = bags_sold * (a['active_user_rate'] / 100)
        
        # Ad revenue (only from active users)
        total_ad_views = active_users * a['avg_monthly_ad_views']
        ad_revenue = total_ad_views * a['cpv_brand_pays']
        
        # Total revenue
        total_revenue = bag_revenue + ad_revenue
        
        # Costs (only active users get rewards)
        total_cash_credits = total_ad_views * a['cash_credit_per_view']
        total_reward_points = total_ad_views * a['reward_points_per_view']
        total_user_rewards = total_cash_credits + total_reward_points
        
        # Margin
        gross_margin = total_ad_views * a['kshipra_margin_per_view']
        total_margin = bag_revenue + gross_margin
        margin_percentage = (total_margin / total_revenue * 100) if total_revenue > 0 else 0
        
        return {
            'bags_sold': bags_sold,
            'bag_revenue': bag_revenue,
            'total_ad_views': total_ad_views,
            'ad_revenue': ad_revenue,
            'total_revenue': total_revenue,
            'total_cash_credits_paid': total_cash_credits,
            'total_reward_points_issued': total_reward_points,
            'total_user_rewards': total_user_rewards,
            'gross_margin': gross_margin,
            'total_margin': total_margin,
            'margin_percentage': margin_percentage,
            'avg_revenue_per_bag': total_revenue / bags_sold if bags_sold > 0 else 0,
            'avg_margin_per_bag': total_margin / bags_sold if bags_sold > 0 else 0
        }
    
    def generate_scenario(self, scenario_type='base'):
        """
        Generate best, base, or worst-case scenarios.
        
        Args:
            scenario_type: 'best', 'base', or 'worst'
        """
        base_assumptions = self.get_default_assumptions()
        
        if scenario_type == 'best':
            # Optimistic scenario
            base_assumptions['avg_monthly_ad_views'] = 20  # High engagement
            base_assumptions['pct_platinum'] = 8           # More premium users
            base_assumptions['pct_gold'] = 20
            base_assumptions['reward_redemption_rate'] = 85
            base_assumptions['repeat_visit_increase'] = 25
            base_assumptions['basket_size_uplift'] = 15
            
        elif scenario_type == 'worst':
            # Conservative scenario
            base_assumptions['avg_monthly_ad_views'] = 6   # Low engagement
            base_assumptions['pct_platinum'] = 1           # Fewer premium users
            base_assumptions['pct_gold'] = 5
            base_assumptions['reward_redemption_rate'] = 50
            base_assumptions['repeat_visit_increase'] = 8
            base_assumptions['basket_size_uplift'] = 3
        
        # Base scenario uses defaults
        return CommerceRewardCalculator(base_assumptions)
    
    def calculate_investor_metrics(self):
        """Calculate investor-specific metrics: LTV, CAC, Payback Period."""
        a = self.assumptions
        
        # Customer Acquisition Cost (CAC)
        # Assume marketing/sales cost to acquire one bag buyer
        marketing_cost_per_user = 2.00  # Assumed $2 to acquire user
        
        # Customer Lifetime Value (LTV)
        # Assume user stays active for 12 months, buying 1 bag/month
        avg_user_lifetime_months = 12
        per_user = self.calculate_revenue_per_user()
        per_bag = self.calculate_revenue_per_bag()
        
        # LTV from user purchases
        bags_purchased_lifetime = avg_user_lifetime_months
        bag_ltv = bags_purchased_lifetime * per_bag['net_margin_per_bag']
        
        # LTV from ad revenue
        total_ad_views_lifetime = avg_user_lifetime_months * a['avg_monthly_ad_views']
        ad_ltv = total_ad_views_lifetime * a['kshipra_margin_per_view']
        
        # Total LTV
        total_ltv = bag_ltv + ad_ltv
        
        # LTV:CAC Ratio
        ltv_cac_ratio = total_ltv / marketing_cost_per_user if marketing_cost_per_user > 0 else 0
        
        # Payback Period (months to recover CAC)
        monthly_margin_per_user = per_user['kshipra_margin_per_user'] + (a['bag_retail_price'] * 1)  # 1 bag/month
        payback_months = marketing_cost_per_user / monthly_margin_per_user if monthly_margin_per_user > 0 else 0
        
        # Unit Economics
        monthly_summary = self.calculate_monthly_summary()
        contribution_margin_per_user = monthly_margin_per_user
        contribution_margin_pct = (contribution_margin_per_user / (a['bag_retail_price'] + per_user['brand_spend_per_user'])) * 100
        
        return {
            'cac': marketing_cost_per_user,
            'ltv': total_ltv,
            'ltv_cac_ratio': ltv_cac_ratio,
            'payback_months': payback_months,
            'avg_user_lifetime_months': avg_user_lifetime_months,
            'contribution_margin_per_user': contribution_margin_per_user,
                        'contribution_margin_pct': contribution_margin_pct
        }
    
    def calculate_business_growth_metrics(self):
        """Calculate bag sales and brand partnership growth projections."""
        a = self.assumptions
        
        # Starting values
        initial_bags = a['bags_sold_per_month']
        initial_brands = a['initial_brands_enrolled']
        bag_growth_rate = a['bag_sales_growth_quarterly'] / 100
        brand_growth_rate = a['brand_growth_rate_quarterly'] / 100
        
        # Calculate quarterly growth for bags
        bags_q1 = initial_bags
        bags_q2 = bags_q1 * (1 + bag_growth_rate)
        bags_q3 = bags_q2 * (1 + bag_growth_rate)
        bags_q4 = bags_q3 * (1 + bag_growth_rate)
        
        avg_bags_year1 = (bags_q1 * 3 + bags_q2 * 3 + bags_q3 * 3 + bags_q4 * 3) / 12
        
        # Calculate quarterly growth for brands
        brands_q1 = initial_brands
        brands_q2 = brands_q1 * (1 + brand_growth_rate)
        brands_q3 = brands_q2 * (1 + brand_growth_rate)
        brands_q4 = brands_q3 * (1 + brand_growth_rate)
        
        avg_brands_year1 = (brands_q1 + brands_q2 + brands_q3 + brands_q4) / 4
        
        # Calculate quarterly revenue with BOTH bag growth AND brand growth
        # Base assumptions with initial brands
        base_fill_rate = 0.85  # 85% fill rate with 10 brands
        
        # Fill rate improves with more brands (diminishing returns)
        def calculate_fill_rate(num_brands):
            # Logarithmic growth: more brands = higher fill rate, capped at 95%
            improvement = (num_brands / initial_brands - 1) * 0.08
            return min(0.95, base_fill_rate + improvement)
        
        fill_rate_q1 = calculate_fill_rate(brands_q1)
        fill_rate_q2 = calculate_fill_rate(brands_q2)
        fill_rate_q3 = calculate_fill_rate(brands_q3)
        fill_rate_q4 = calculate_fill_rate(brands_q4)
        
        # Get base monthly metrics (using initial bags_sold_per_month)
        base_monthly_summary = self.calculate_monthly_summary()
        base_bag_revenue = initial_bags * a['bag_retail_price']
        base_ad_revenue = base_monthly_summary['ad_revenue']
        
        # Calculate quarterly revenue accounting for BOTH growth factors
        # Q1: Month 1-3
        bag_revenue_q1 = bags_q1 * a['bag_retail_price'] * 3
        active_users_q1 = bags_q1 * (a['active_user_rate'] / 100)
        ad_views_q1 = active_users_q1 * a['avg_monthly_ad_views']
        ad_revenue_q1 = ad_views_q1 * a['cpv_brand_pays'] * (fill_rate_q1 / base_fill_rate) * 3
        total_revenue_q1 = bag_revenue_q1 + ad_revenue_q1
        
        # Q2: Month 4-6
        bag_revenue_q2 = bags_q2 * a['bag_retail_price'] * 3
        active_users_q2 = bags_q2 * (a['active_user_rate'] / 100)
        ad_views_q2 = active_users_q2 * a['avg_monthly_ad_views']
        ad_revenue_q2 = ad_views_q2 * a['cpv_brand_pays'] * (fill_rate_q2 / base_fill_rate) * 3
        total_revenue_q2 = bag_revenue_q2 + ad_revenue_q2
        
        # Q3: Month 7-9
        bag_revenue_q3 = bags_q3 * a['bag_retail_price'] * 3
        active_users_q3 = bags_q3 * (a['active_user_rate'] / 100)
        ad_views_q3 = active_users_q3 * a['avg_monthly_ad_views']
        ad_revenue_q3 = ad_views_q3 * a['cpv_brand_pays'] * (fill_rate_q3 / base_fill_rate) * 3
        total_revenue_q3 = bag_revenue_q3 + ad_revenue_q3
        
        # Q4: Month 10-12
        bag_revenue_q4 = bags_q4 * a['bag_retail_price'] * 3
        active_users_q4 = bags_q4 * (a['active_user_rate'] / 100)
        ad_views_q4 = active_users_q4 * a['avg_monthly_ad_views']
        ad_revenue_q4 = ad_views_q4 * a['cpv_brand_pays'] * (fill_rate_q4 / base_fill_rate) * 3
        total_revenue_q4 = bag_revenue_q4 + ad_revenue_q4
        
        # Annual totals
        total_annual_bag_revenue = bag_revenue_q1 + bag_revenue_q2 + bag_revenue_q3 + bag_revenue_q4
        total_annual_ad_revenue = ad_revenue_q1 + ad_revenue_q2 + ad_revenue_q3 + ad_revenue_q4
        total_annual_revenue = total_annual_bag_revenue + total_annual_ad_revenue
        
        # Compare to base (no growth scenario)
        base_annual_bag_revenue = base_bag_revenue * 12
        base_annual_ad_revenue = base_ad_revenue * 12
        base_annual_revenue = base_annual_bag_revenue + base_annual_ad_revenue
        
        # Revenue lifts
        bag_revenue_lift = total_annual_bag_revenue - base_annual_bag_revenue
        bag_revenue_lift_pct = (bag_revenue_lift / base_annual_bag_revenue * 100) if base_annual_bag_revenue > 0 else 0
        
        ad_revenue_lift = total_annual_ad_revenue - base_annual_ad_revenue
        ad_revenue_lift_pct = (ad_revenue_lift / base_annual_ad_revenue * 100) if base_annual_ad_revenue > 0 else 0
        
        total_revenue_lift = total_annual_revenue - base_annual_revenue
        total_revenue_lift_pct = (total_revenue_lift / base_annual_revenue * 100) if base_annual_revenue > 0 else 0
        
        # Calculate average fill rate across year
        avg_fill_rate = (fill_rate_q1 + fill_rate_q2 + fill_rate_q3 + fill_rate_q4) / 4
        
        return {
            # Bag growth
            'initial_bags': initial_bags,
            'bag_growth_rate_quarterly': a['bag_sales_growth_quarterly'],
            'bags_q1': bags_q1,
            'bags_q2': bags_q2,
            'bags_q3': bags_q3,
            'bags_q4': bags_q4,
            'avg_bags_year1': avg_bags_year1,
            'bag_revenue_q1': bag_revenue_q1,
            'bag_revenue_q2': bag_revenue_q2,
            'bag_revenue_q3': bag_revenue_q3,
            'bag_revenue_q4': bag_revenue_q4,
            
            # Brand growth
            'initial_brands': initial_brands,
            'brand_growth_rate_quarterly': a['brand_growth_rate_quarterly'],
            'brands_q1': brands_q1,
            'brands_q2': brands_q2,
            'brands_q3': brands_q3,
            'brands_q4': brands_q4,
            'avg_brands_year1': avg_brands_year1,
            'base_fill_rate': base_fill_rate * 100,
            'year_end_fill_rate': fill_rate_q4 * 100,
            'avg_fill_rate': avg_fill_rate * 100,
            
            # Revenue breakdown
            'ad_revenue_q1': ad_revenue_q1,
            'ad_revenue_q2': ad_revenue_q2,
            'ad_revenue_q3': ad_revenue_q3,
            'ad_revenue_q4': ad_revenue_q4,
            'total_revenue_q1': total_revenue_q1,
            'total_revenue_q2': total_revenue_q2,
            'total_revenue_q3': total_revenue_q3,
            'total_revenue_q4': total_revenue_q4,
            
            # Annual totals
            'base_annual_bag_revenue': base_annual_bag_revenue,
            'total_annual_bag_revenue': total_annual_bag_revenue,
            'bag_revenue_lift': bag_revenue_lift,
            'bag_revenue_lift_pct': bag_revenue_lift_pct,
            
            'base_annual_ad_revenue': base_annual_ad_revenue,
            'total_annual_ad_revenue': total_annual_ad_revenue,
            'ad_revenue_lift': ad_revenue_lift,
            'ad_revenue_lift_pct': ad_revenue_lift_pct,
            
            'base_annual_revenue': base_annual_revenue,
            'total_annual_revenue': total_annual_revenue,
            'total_revenue_lift': total_revenue_lift,
            'total_revenue_lift_pct': total_revenue_lift_pct
        }
    
    def print_summary(self):
        """Print comprehensive business model summary."""
        print("\n" + "="*70)
        print("KSHIPRA AI COMMERCE-REWARD BUSINESS MODEL - SUMMARY")
        print("="*70)
        
        # Key Assumptions
        print("\n📊 KEY ASSUMPTIONS:")
        print(f"Bag Retail Price: ${self.assumptions['bag_retail_price']:.2f}")
        print(f"Monthly Bags Sold (Initial): {self.assumptions['bags_sold_per_month']:,}")
        print(f"Quarterly Bag Growth: {self.assumptions['bag_sales_growth_quarterly']}%")
        print(f"Initial Brands Enrolled: {self.assumptions['initial_brands_enrolled']}")
        print(f"Quarterly Brand Growth: {self.assumptions['brand_growth_rate_quarterly']}%")
        print(f"Brand CPV: ${self.assumptions['cpv_brand_pays']:.2f}")
        print(f"Cash Credit per View: ${self.assumptions['cash_credit_per_view']:.2f}")
        print(f"Reward Points per View: ${self.assumptions['reward_points_per_view']:.2f}")
        print(f"Kshipra Margin per View: ${self.assumptions['kshipra_margin_per_view']:.2f}")
        print(f"Avg Ads to Recover Bag Cost: {self.assumptions['avg_ads_to_recover_bag']}")
        
        # Business Growth Impact
        growth = self.calculate_business_growth_metrics()
        print("\n📈 BUSINESS GROWTH PROJECTIONS (Year 1):")
        print(f"\nBag Sales Growth:")
        print(f"  Q1: {growth['bags_q1']:,.0f} bags → Q4: {growth['bags_q4']:,.0f} bags")
        print(f"  Year 1 Average: {growth['avg_bags_year1']:,.0f} bags/month")
        print(f"  Revenue Lift: ${growth['bag_revenue_lift']:,.0f} (+{growth['bag_revenue_lift_pct']:.1f}%)")
        
        print(f"\nBrand Partnership Growth:")
        print(f"  Q1: {growth['brands_q1']:.0f} brands → Q4: {growth['brands_q4']:.0f} brands")
        print(f"  Year 1 Average: {growth['avg_brands_year1']:.1f} brands")
        print(f"  Ad Fill Rate: {growth['base_fill_rate']:.0f}% → {growth['year_end_fill_rate']:.0f}%")
        print(f"  Revenue Lift: ${growth['ad_revenue_lift']:,.0f} (+{growth['ad_revenue_lift_pct']:.1f}%)")
        
        print(f"\nTotal Combined Growth Impact:")
        print(f"  Base Annual Revenue: ${growth['base_annual_revenue']:,.0f}")
        print(f"  Year 1 Actual Revenue: ${growth['total_annual_revenue']:,.0f}")
        print(f"  Total Revenue Lift: ${growth['total_revenue_lift']:,.0f} (+{growth['total_revenue_lift_pct']:.1f}%)")
        
        # Call helper to print rest of summary
        self._print_remaining_summary()
    
    def calculate_sensitivity_analysis(self, variable, values):
        """
        Run sensitivity analysis on a single variable.
        
        Args:
            variable: Name of assumption to vary (e.g., 'avg_monthly_ad_views')
            values: List of values to test
        
        Returns:
            List of results for each value
        """
        results = []
        original_value = self.assumptions[variable]
        
        for value in values:
            # Update assumption
            self.assumptions[variable] = value
            
            # Calculate metrics
            summary = self.calculate_monthly_summary()
            per_user = self.calculate_revenue_per_user()
            
            results.append({
                'value': value,
                'total_revenue': summary['total_revenue'],
                'total_margin': summary['total_margin'],
                'margin_pct': summary['margin_percentage'],
                'user_rewards': per_user['total_user_rewards']
            })
        
        # Restore original value
        self.assumptions[variable] = original_value
        
        return results
    
    def _print_remaining_summary(self):
        """Print remaining summary sections (called by print_summary)."""
        # Monthly Summary
        summary = self.calculate_monthly_summary()
        print("\n💰 MONTHLY REVENUE:")
        print(f"Bag Sales Revenue: ${summary['bag_revenue']:,.2f}")
        print(f"Ad Revenue: ${summary['ad_revenue']:,.2f}")
        print(f"Total Revenue: ${summary['total_revenue']:,.2f}")
        print(f"Total Margin: ${summary['total_margin']:,.2f} ({summary['margin_percentage']:.1f}%)")
        
        # Per User
        per_user = self.calculate_revenue_per_user()
        print("\n👤 PER USER METRICS:")
        print(f"Monthly Ad Views: {per_user['monthly_ad_views']}")
        print(f"Brand Spend: ${per_user['brand_spend_per_user']:.2f}")
        print(f"User Rewards: ${per_user['total_user_rewards']:.2f}")
        print(f"Kshipra Margin: ${per_user['kshipra_margin_per_user']:.2f}")
        
        # Per Bag
        per_bag = self.calculate_revenue_per_bag()
        print("\n🛍️  PER BAG METRICS:")
        print(f"Total Revenue: ${per_bag['total_revenue_per_bag']:.2f}")
        print(f"Net Margin: ${per_bag['net_margin_per_bag']:.2f}")
        print(f"User Net Cost: ${per_bag['user_net_cost']:.2f}")
        
        # Store Value
        store = self.calculate_store_value()
        print("\n🏪 STORE VALUE:")
        print(f"Monthly Reward Redemptions: ${store['reward_redemption_value']:,.2f}")
        print(f"Incremental Basket Value: ${store['incremental_basket_value']:,.2f}")
        print(f"Total Store Value: ${store['total_monthly_store_value']:,.2f}")
        
        # Brand ROI
        brand = self.calculate_brand_roi()
        print("\n📈 BRAND ROI:")
        print(f"Cost per Engaged Customer: ${brand['cost_per_engaged_customer']:.2f}")
        print(f"BagBuddy CPM: ${brand['bagbuddy_cpm']:.2f}")
        print(f"Meta CPM: ${brand['meta_cpm']:.2f}")
        print(f"Savings vs Meta: {brand['bagbuddy_savings_vs_meta_pct']:.1f}%")
        print(f"Cost per Basket Influenced: ${brand['cost_per_basket_influenced']:.2f}")
        
        # Investor Metrics
        investor = self.calculate_investor_metrics()
        print("\n💼 INVESTOR METRICS:")
        print(f"Customer LTV: ${investor['ltv']:.2f}")
        print(f"CAC: ${investor['cac']:.2f}")
        print(f"LTV:CAC Ratio: {investor['ltv_cac_ratio']:.1f}x")
        print(f"Payback Period: {investor['payback_months']:.1f} months")
        print(f"Contribution Margin: ${investor['contribution_margin_per_user']:.2f}/user ({investor['contribution_margin_pct']:.1f}%)")
        
        print("\n" + "="*70)


if __name__ == "__main__":
    # Demo: Base scenario
    print("\n🎯 BASE SCENARIO")
    calc_base = CommerceRewardCalculator()
    calc_base.print_summary()
    
    # Demo: Best case
    print("\n\n🚀 BEST CASE SCENARIO")
    calc_best = calc_base.generate_scenario('best')
    calc_best.print_summary()
    
    # Demo: Worst case
    print("\n\n⚠️  WORST CASE SCENARIO")
    calc_worst = calc_base.generate_scenario('worst')
    calc_worst.print_summary()
