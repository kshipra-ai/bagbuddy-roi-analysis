"""
Export Kshipra Commerce-Reward Model to Interactive Excel with Formulas

Creates a comprehensive Excel workbook with:
- Editable assumptions (yellow cells)
- Formula-based calculations
- Three scenarios (Base, Best, Worst)
- Dashboard summary
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from commerce_reward_calculator import CommerceRewardCalculator

def create_commerce_reward_excel():
    """Create interactive Excel calculator for commerce-reward model."""
    
    wb = Workbook()
    
    # Remove default sheet
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    
    # Create sheets
    ws_dashboard = wb.create_sheet("Dashboard")
    ws_assumptions = wb.create_sheet("Assumptions")
    ws_calculations = wb.create_sheet("Calculations")
    ws_scenarios = wb.create_sheet("Scenarios")
    ws_investor = wb.create_sheet("Investor Metrics")
    ws_sensitivity = wb.create_sheet("Sensitivity Analysis")
    
    # Styling helpers
    def style_header(cell, color="2E75B6"):
        cell.font = Font(bold=True, size=11, color="FFFFFF")
        cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        cell.alignment = Alignment(horizontal="left", vertical="center")
    
    def style_editable(cell):
        cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        cell.font = Font(bold=True)
    
    def style_section(cell):
        cell.font = Font(bold=True, size=10)
        cell.fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    
    # ===== ASSUMPTIONS SHEET =====
    row = 1
    ws = ws_assumptions
    
    ws[f'A{row}'] = "KSHIPRA COMMERCE-REWARD MODEL - ASSUMPTIONS"
    style_header(ws[f'A{row}'], "FF6600")
    ws.merge_cells(f'A{row}:D{row}')
    row += 2
    
    # Pricing Section
    ws[f'A{row}'] = "PRICING & VOLUME"
    style_section(ws[f'A{row}'])
    ws.merge_cells(f'A{row}:D{row}')
    row += 1
    
    ws[f'A{row}'] = "Bag Retail Price"
    ws[f'B{row}'] = 0.40
    style_editable(ws[f'B{row}'])
    ws[f'B{row}'].number_format = '$0.00'
    ws[f'C{row}'] = "$"
    bag_price_cell = f'B{row}'
    row += 1
    
    ws[f'A{row}'] = "Bags Sold per Month"
    ws[f'B{row}'] = 10000
    style_editable(ws[f'B{row}'])
    ws[f'B{row}'].number_format = '#,##0'
    ws[f'C{row}'] = "bags"
    bags_sold_cell = f'B{row}'
    row += 2
    
    # Ad Economics
    ws[f'A{row}'] = "AD ECONOMICS"
    style_section(ws[f'A{row}'])
    ws.merge_cells(f'A{row}:D{row}')
    row += 1
    
    ws[f'A{row}'] = "CPV - Brand Pays per View"
    ws[f'B{row}'] = 0.25
    style_editable(ws[f'B{row}'])
    ws[f'B{row}'].number_format = '$0.00'
    ws[f'C{row}'] = "$"
    cpv_cell = f'B{row}'
    row += 1
    
    ws[f'A{row}'] = "Cash Credit per View"
    ws[f'B{row}'] = 0.08
    style_editable(ws[f'B{row}'])
    ws[f'B{row}'].number_format = '$0.00'
    ws[f'C{row}'] = "$"
    cash_credit_cell = f'B{row}'
    row += 1
    
    ws[f'A{row}'] = "Reward Points per View"
    ws[f'B{row}'] = 0.07
    style_editable(ws[f'B{row}'])
    ws[f'B{row}'].number_format = '$0.00'
    ws[f'C{row}'] = "$"
    reward_points_cell = f'B{row}'
    row += 1
    
    ws[f'A{row}'] = "Kshipra Margin per View"
    ws[f'B{row}'] = 0.10
    style_editable(ws[f'B{row}'])
    ws[f'B{row}'].number_format = '$0.00'
    ws[f'C{row}'] = "$"
    margin_per_view_cell = f'B{row}'
    row += 2
    
    # User Behavior
    ws[f'A{row}'] = "USER BEHAVIOR"
    style_section(ws[f'A{row}'])
    ws.merge_cells(f'A{row}:D{row}')
    row += 1
    
    ws[f'A{row}'] = "Avg Ads to Recover Bag Cost"
    ws[f'B{row}'] = 3
    style_editable(ws[f'B{row}'])
    ws[f'B{row}'].number_format = '0'
    ws[f'C{row}'] = "ads"
    ads_to_recover_cell = f'B{row}'
    row += 1
    
    ws[f'A{row}'] = "Avg Monthly Ad Views per User"
    ws[f'B{row}'] = 12
    style_editable(ws[f'B{row}'])
    ws[f'B{row}'].number_format = '0'
    ws[f'C{row}'] = "views"
    monthly_views_cell = f'B{row}'
    row += 2
    
    # Tier Limits
    ws[f'A{row}'] = "TIER CASH CREDIT LIMITS"
    style_section(ws[f'A{row}'])
    ws.merge_cells(f'A{row}:D{row}')
    row += 1
    
    ws[f'A{row}'] = "Bronze - Bag Values/Month"
    ws[f'B{row}'] = 1
    style_editable(ws[f'B{row}'])
    bronze_limit_cell = f'B{row}'
    row += 1
    
    ws[f'A{row}'] = "Silver - Bag Values/Month"
    ws[f'B{row}'] = 3
    style_editable(ws[f'B{row}'])
    silver_limit_cell = f'B{row}'
    row += 1
    
    ws[f'A{row}'] = "Gold - Bag Values/Month"
    ws[f'B{row}'] = 7
    style_editable(ws[f'B{row}'])
    gold_limit_cell = f'B{row}'
    row += 1
    
    ws[f'A{row}'] = "Platinum - Daily Cap ($)"
    ws[f'B{row}'] = 1.00
    style_editable(ws[f'B{row}'])
    ws[f'B{row}'].number_format = '$0.00'
    platinum_cap_cell = f'B{row}'
    row += 2
    
    # User Distribution
    ws[f'A{row}'] = "USER DISTRIBUTION BY TIER (%)"
    style_section(ws[f'A{row}'])
    ws.merge_cells(f'A{row}:D{row}')
    row += 1
    
    ws[f'A{row}'] = "Bronze %"
    ws[f'B{row}'] = 60
    style_editable(ws[f'B{row}'])
    ws[f'C{row}'] = "%"
    pct_bronze_cell = f'B{row}'
    row += 1
    
    ws[f'A{row}'] = "Silver %"
    ws[f'B{row}'] = 25
    style_editable(ws[f'B{row}'])
    ws[f'C{row}'] = "%"
    pct_silver_cell = f'B{row}'
    row += 1
    
    ws[f'A{row}'] = "Gold %"
    ws[f'B{row}'] = 12
    style_editable(ws[f'B{row}'])
    ws[f'C{row}'] = "%"
    pct_gold_cell = f'B{row}'
    row += 1
    
    ws[f'A{row}'] = "Platinum %"
    ws[f'B{row}'] = 3
    style_editable(ws[f'B{row}'])
    ws[f'C{row}'] = "%"
    pct_platinum_cell = f'B{row}'
    row += 2
    
    # Store & Redemption
    ws[f'A{row}'] = "STORE & REDEMPTION METRICS"
    style_section(ws[f'A{row}'])
    ws.merge_cells(f'A{row}:D{row}')
    row += 1
    
    ws[f'A{row}'] = "Reward Redemption Rate"
    ws[f'B{row}'] = 75
    style_editable(ws[f'B{row}'])
    ws[f'C{row}'] = "%"
    redemption_rate_cell = f'B{row}'
    row += 1
    
    ws[f'A{row}'] = "Repeat Visit Increase"
    ws[f'B{row}'] = 15
    style_editable(ws[f'B{row}'])
    ws[f'C{row}'] = "%"
    repeat_visit_cell = f'B{row}'
    row += 1
    
    ws[f'A{row}'] = "Basket Size Uplift"
    ws[f'B{row}'] = 8
    style_editable(ws[f'B{row}'])
    ws[f'C{row}'] = "%"
    basket_uplift_cell = f'B{row}'
    row += 1
    
    ws[f'A{row}'] = "Average Basket Value"
    ws[f'B{row}'] = 25.00
    style_editable(ws[f'B{row}'])
    ws[f'B{row}'].number_format = '$0.00'
    ws[f'C{row}'] = "$"
    avg_basket_cell = f'B{row}'
    row += 2
    
    # Benchmarks
    ws[f'A{row}'] = "BRAND BENCHMARKS"
    style_section(ws[f'A{row}'])
    ws.merge_cells(f'A{row}:D{row}')
    row += 1
    
    ws[f'A{row}'] = "Meta CPM"
    ws[f'B{row}'] = 10.00
    style_editable(ws[f'B{row}'])
    ws[f'B{row}'].number_format = '$0.00'
    ws[f'C{row}'] = "$"
    meta_cpm_cell = f'B{row}'
    row += 1
    
    ws[f'A{row}'] = "TikTok CPM"
    ws[f'B{row}'] = 8.50
    style_editable(ws[f'B{row}'])
    ws[f'B{row}'].number_format = '$0.00'
    ws[f'C{row}'] = "$"
    tiktok_cpm_cell = f'B{row}'
    row += 1
    
    ws[f'A{row}'] = "Industry Avg CTR"
    ws[f'B{row}'] = 1.0
    style_editable(ws[f'B{row}'])
    ws[f'C{row}'] = "%"
    ctr_cell = f'B{row}'
    
    # Auto-fit columns
    for col in ['A', 'B', 'C', 'D']:
        ws.column_dimensions[col].width = 25
    
    # ===== CALCULATIONS SHEET =====
    ws = ws_calculations
    row = 1
    
    ws[f'A{row}'] = "DETAILED CALCULATIONS"
    style_header(ws[f'A{row}'], "2E75B6")
    ws.merge_cells(f'A{row}:C{row}')
    row += 2
    
    # Monthly Summary
    ws[f'A{row}'] = "MONTHLY BUSINESS SUMMARY"
    style_section(ws[f'A{row}'])
    ws.merge_cells(f'A{row}:C{row}')
    row += 1
    
    ws[f'A{row}'] = "Bags Sold"
    ws[f'B{row}'] = f"=Assumptions!{bags_sold_cell}"
    ws[f'B{row}'].number_format = '#,##0'
    row += 1
    
    ws[f'A{row}'] = "Bag Revenue"
    ws[f'B{row}'] = f"=Assumptions!{bags_sold_cell}*Assumptions!{bag_price_cell}"
    ws[f'B{row}'].number_format = '$#,##0.00'
    bag_revenue_row = row
    row += 1
    
    ws[f'A{row}'] = "Total Ad Views"
    ws[f'B{row}'] = f"=Assumptions!{bags_sold_cell}*Assumptions!{monthly_views_cell}"
    ws[f'B{row}'].number_format = '#,##0'
    total_views_row = row
    row += 1
    
    ws[f'A{row}'] = "Ad Revenue (Brand Spend)"
    ws[f'B{row}'] = f"=B{total_views_row}*Assumptions!{cpv_cell}"
    ws[f'B{row}'].number_format = '$#,##0.00'
    ad_revenue_row = row
    row += 1
    
    ws[f'A{row}'] = "Total Revenue"
    ws[f'B{row}'] = f"=B{bag_revenue_row}+B{ad_revenue_row}"
    ws[f'B{row}'].number_format = '$#,##0.00'
    ws[f'B{row}'].font = Font(bold=True)
    total_revenue_row = row
    row += 2
    
    ws[f'A{row}'] = "Total Cash Credits Paid"
    ws[f'B{row}'] = f"=B{total_views_row}*Assumptions!{cash_credit_cell}"
    ws[f'B{row}'].number_format = '$#,##0.00'
    row += 1
    
    ws[f'A{row}'] = "Total Reward Points Issued"
    ws[f'B{row}'] = f"=B{total_views_row}*Assumptions!{reward_points_cell}"
    ws[f'B{row}'].number_format = '$#,##0.00'
    row += 1
    
    ws[f'A{row}'] = "Gross Margin (from Ads)"
    ws[f'B{row}'] = f"=B{total_views_row}*Assumptions!{margin_per_view_cell}"
    ws[f'B{row}'].number_format = '$#,##0.00'
    gross_margin_row = row
    row += 1
    
    ws[f'A{row}'] = "Total Margin"
    ws[f'B{row}'] = f"=B{bag_revenue_row}+B{gross_margin_row}"
    ws[f'B{row}'].number_format = '$#,##0.00'
    ws[f'B{row}'].font = Font(bold=True, color="00AA00")
    total_margin_row = row
    row += 1
    
    ws[f'A{row}'] = "Margin %"
    ws[f'B{row}'] = f"=B{total_margin_row}/B{total_revenue_row}"
    ws[f'B{row}'].number_format = '0.0%'
    ws[f'B{row}'].font = Font(bold=True)
    row += 2
    
    # Per User Metrics
    ws[f'A{row}'] = "PER USER METRICS"
    style_section(ws[f'A{row}'])
    ws.merge_cells(f'A{row}:C{row}')
    row += 1
    
    ws[f'A{row}'] = "Monthly Ad Views"
    ws[f'B{row}'] = f"=Assumptions!{monthly_views_cell}"
    row += 1
    
    ws[f'A{row}'] = "Brand Spend per User"
    ws[f'B{row}'] = f"=Assumptions!{monthly_views_cell}*Assumptions!{cpv_cell}"
    ws[f'B{row}'].number_format = '$0.00'
    row += 1
    
    ws[f'A{row}'] = "User Rewards per Month"
    ws[f'B{row}'] = f"=Assumptions!{monthly_views_cell}*(Assumptions!{cash_credit_cell}+Assumptions!{reward_points_cell})"
    ws[f'B{row}'].number_format = '$0.00'
    row += 1
    
    ws[f'A{row}'] = "Kshipra Margin per User"
    ws[f'B{row}'] = f"=Assumptions!{monthly_views_cell}*Assumptions!{margin_per_view_cell}"
    ws[f'B{row}'].number_format = '$0.00'
    row += 2
    
    # Per Bag Metrics
    ws[f'A{row}'] = "PER BAG METRICS"
    style_section(ws[f'A{row}'])
    ws.merge_cells(f'A{row}:C{row}')
    row += 1
    
    ws[f'A{row}'] = "Bag Retail Revenue"
    ws[f'B{row}'] = f"=Assumptions!{bag_price_cell}"
    ws[f'B{row}'].number_format = '$0.00'
    row += 1
    
    ws[f'A{row}'] = "Ad Views per Bag"
    ws[f'B{row}'] = f"=Assumptions!{ads_to_recover_cell}"
    row += 1
    
    ws[f'A{row}'] = "Brand Revenue per Bag"
    ws[f'B{row}'] = f"=Assumptions!{ads_to_recover_cell}*Assumptions!{cpv_cell}"
    ws[f'B{row}'].number_format = '$0.00'
    row += 1
    
    ws[f'A{row}'] = "User Earnings per Bag"
    ws[f'B{row}'] = f"=Assumptions!{ads_to_recover_cell}*(Assumptions!{cash_credit_cell}+Assumptions!{reward_points_cell})"
    ws[f'B{row}'].number_format = '$0.00'
    user_earnings_row = row
    row += 1
    
    ws[f'A{row}'] = "User Net Cost"
    ws[f'B{row}'] = f"=Assumptions!{bag_price_cell}-B{user_earnings_row}"
    ws[f'B{row}'].number_format = '$0.00'
    ws[f'B{row}'].font = Font(bold=True)
    row += 2
    
    # Store Value
    ws[f'A{row}'] = "STORE VALUE METRICS"
    style_section(ws[f'A{row}'])
    ws.merge_cells(f'A{row}:C{row}')
    row += 1
    
    ws[f'A{row}'] = "Total Rewards Generated"
    ws[f'B{row}'] = f"=B{total_views_row}*(Assumptions!{cash_credit_cell}+Assumptions!{reward_points_cell})"
    ws[f'B{row}'].number_format = '$#,##0.00'
    total_rewards_row = row
    row += 1
    
    ws[f'A{row}'] = "Reward Redemption Value"
    ws[f'B{row}'] = f"=B{total_rewards_row}*Assumptions!{redemption_rate_cell}/100"
    ws[f'B{row}'].number_format = '$#,##0.00'
    row += 1
    
    ws[f'A{row}'] = "Estimated Transactions"
    ws[f'B{row}'] = f"=Assumptions!{bags_sold_cell}*Assumptions!{redemption_rate_cell}/100"
    ws[f'B{row}'].number_format = '#,##0'
    transactions_row = row
    row += 1
    
    ws[f'A{row}'] = "Baseline Basket Value"
    ws[f'B{row}'] = f"=B{transactions_row}*Assumptions!{avg_basket_cell}"
    ws[f'B{row}'].number_format = '$#,##0.00'
    baseline_basket_row = row
    row += 1
    
    ws[f'A{row}'] = "Incremental Basket Value"
    ws[f'B{row}'] = f"=B{baseline_basket_row}*Assumptions!{basket_uplift_cell}/100"
    ws[f'B{row}'].number_format = '$#,##0.00'
    row += 2
    
    # Brand ROI
    ws[f'A{row}'] = "BRAND ROI METRICS"
    style_section(ws[f'A{row}'])
    ws.merge_cells(f'A{row}:C{row}')
    row += 1
    
    ws[f'A{row}'] = "BagBuddy CPM"
    ws[f'B{row}'] = f"=Assumptions!{cpv_cell}*1000"
    ws[f'B{row}'].number_format = '$0.00'
    row += 1
    
    ws[f'A{row}'] = "Cost per Basket Influenced"
    ws[f'B{row}'] = f"=B{ad_revenue_row}/B{transactions_row}"
    ws[f'B{row}'].number_format = '$0.00'
    row += 1
    
    ws[f'A{row}'] = "BagBuddy vs Meta Savings"
    ws[f'B{row}'] = f"=(Assumptions!{meta_cpm_cell}-Assumptions!{cpv_cell}*1000)/Assumptions!{meta_cpm_cell}"
    ws[f'B{row}'].number_format = '0.0%'
    
    # Auto-fit
    for col in ['A', 'B', 'C']:
        ws.column_dimensions[col].width = 30
    
    # ===== DASHBOARD =====
    ws = ws_dashboard
    row = 1
    
    ws[f'A{row}'] = "KSHIPRA COMMERCE-REWARD MODEL DASHBOARD"
    ws[f'A{row}'].font = Font(bold=True, size=14, color="FFFFFF")
    ws[f'A{row}'].fill = PatternFill(start_color="FF6600", end_color="FF6600", fill_type="solid")
    ws.merge_cells(f'A{row}:D{row}')
    row += 2
    
    ws[f'A{row}'] = "KEY METRICS"
    style_section(ws[f'A{row}'])
    ws.merge_cells(f'A{row}:D{row}')
    row += 1
    
    ws[f'A{row}'] = "Monthly Revenue:"
    ws[f'B{row}'] = f"=Calculations!B{total_revenue_row}"
    ws[f'B{row}'].number_format = '$#,##0'
    ws[f'B{row}'].font = Font(bold=True, size=12)
    row += 1
    
    ws[f'A{row}'] = "Monthly Margin:"
    ws[f'B{row}'] = f"=Calculations!B{total_margin_row}"
    ws[f'B{row}'].number_format = '$#,##0'
    ws[f'B{row}'].font = Font(bold=True, size=12, color="00AA00")
    row += 1
    
    ws[f'A{row}'] = "Margin %:"
    ws[f'B{row}'] = f"=Calculations!B{total_margin_row}/Calculations!B{total_revenue_row}"
    ws[f'B{row}'].number_format = '0.0%'
    ws[f'B{row}'].font = Font(bold=True, size=12)
    row += 2
    
    ws[f'A{row}'] = "User Net Bag Cost:"
    ws[f'B{row}'] = f"=Calculations!B{user_earnings_row+1}"
    ws[f'B{row}'].number_format = '$0.00'
    ws[f'B{row}'].font = Font(bold=True, size=12)
    row += 2
    
    ws[f'A{row}'] = "📝 Edit assumptions in the 'Assumptions' sheet (yellow cells)"
    ws[f'A{row}'].font = Font(italic=True, color="666666")
    ws.merge_cells(f'A{row}:D{row}')
    
    for col in ['A', 'B', 'C', 'D']:
        ws.column_dimensions[col].width = 25
    
    # ===== SCENARIOS SHEET =====
    ws = ws_scenarios
    row = 1
    
    ws[f'A{row}'] = "SCENARIO COMPARISON"
    style_header(ws[f'A{row}'], "2E75B6")
    ws.merge_cells(f'A{row}:E{row}')
    row += 2
    
    # Create scenarios
    calc_base = CommerceRewardCalculator()
    calc_best = calc_base.generate_scenario('best')
    calc_worst = calc_base.generate_scenario('worst')
    
    # Headers
    ws[f'A{row}'] = "Metric"
    ws[f'B{row}'] = "Worst Case"
    ws[f'C{row}'] = "Base Case"
    ws[f'D{row}'] = "Best Case"
    ws[f'E{row}'] = "Unit"
    for col in ['A', 'B', 'C', 'D', 'E']:
        style_section(ws[f'{col}{row}'])
    row += 1
    
    # Get summaries
    worst_summary = calc_worst.calculate_monthly_summary()
    base_summary = calc_base.calculate_monthly_summary()
    best_summary = calc_best.calculate_monthly_summary()
    
    # Monthly Revenue
    ws[f'A{row}'] = "Monthly Revenue"
    ws[f'B{row}'] = worst_summary['total_revenue']
    ws[f'C{row}'] = base_summary['total_revenue']
    ws[f'D{row}'] = best_summary['total_revenue']
    ws[f'E{row}'] = "$"
    for col in ['B', 'C', 'D']:
        ws[f'{col}{row}'].number_format = '$#,##0'
    row += 1
    
    ws[f'A{row}'] = "Monthly Margin"
    ws[f'B{row}'] = worst_summary['total_margin']
    ws[f'C{row}'] = base_summary['total_margin']
    ws[f'D{row}'] = best_summary['total_margin']
    ws[f'E{row}'] = "$"
    for col in ['B', 'C', 'D']:
        ws[f'{col}{row}'].number_format = '$#,##0'
    row += 1
    
    ws[f'A{row}'] = "Margin %"
    ws[f'B{row}'] = worst_summary['margin_percentage'] / 100
    ws[f'C{row}'] = base_summary['margin_percentage'] / 100
    ws[f'D{row}'] = best_summary['margin_percentage'] / 100
    ws[f'E{row}'] = "%"
    for col in ['B', 'C', 'D']:
        ws[f'{col}{row}'].number_format = '0.0%'
    row += 1
    
    ws[f'A{row}'] = "Total Ad Views"
    ws[f'B{row}'] = worst_summary['total_ad_views']
    ws[f'C{row}'] = base_summary['total_ad_views']
    ws[f'D{row}'] = best_summary['total_ad_views']
    ws[f'E{row}'] = "views"
    for col in ['B', 'C', 'D']:
        ws[f'{col}{row}'].number_format = '#,##0'
    row += 2
    
    # Store Value
    worst_store = calc_worst.calculate_store_value()
    base_store = calc_base.calculate_store_value()
    best_store = calc_best.calculate_store_value()
    
    ws[f'A{row}'] = "Store Value Created"
    ws[f'B{row}'] = worst_store['total_monthly_store_value']
    ws[f'C{row}'] = base_store['total_monthly_store_value']
    ws[f'D{row}'] = best_store['total_monthly_store_value']
    ws[f'E{row}'] = "$"
    for col in ['B', 'C', 'D']:
        ws[f'{col}{row}'].number_format = '$#,##0'
    row += 1
    
    ws[f'A{row}'] = "Reward Redemptions"
    ws[f'B{row}'] = worst_store['reward_redemption_value']
    ws[f'C{row}'] = base_store['reward_redemption_value']
    ws[f'D{row}'] = best_store['reward_redemption_value']
    ws[f'E{row}'] = "$"
    for col in ['B', 'C', 'D']:
        ws[f'{col}{row}'].number_format = '$#,##0'
    row += 2
    
    # Key Assumptions Differences
    ws[f'A{row}'] = "KEY ASSUMPTION DIFFERENCES"
    style_section(ws[f'A{row}'])
    ws.merge_cells(f'A{row}:E{row}')
    row += 1
    
    ws[f'A{row}'] = "Avg Monthly Ad Views/User"
    ws[f'B{row}'] = calc_worst.assumptions['avg_monthly_ad_views']
    ws[f'C{row}'] = calc_base.assumptions['avg_monthly_ad_views']
    ws[f'D{row}'] = calc_best.assumptions['avg_monthly_ad_views']
    row += 1
    
    ws[f'A{row}'] = "Platinum Users %"
    ws[f'B{row}'] = calc_worst.assumptions['pct_platinum']
    ws[f'C{row}'] = calc_base.assumptions['pct_platinum']
    ws[f'D{row}'] = calc_best.assumptions['pct_platinum']
    ws[f'E{row}'] = "%"
    row += 1
    
    ws[f'A{row}'] = "Redemption Rate %"
    ws[f'B{row}'] = calc_worst.assumptions['reward_redemption_rate']
    ws[f'C{row}'] = calc_base.assumptions['reward_redemption_rate']
    ws[f'D{row}'] = calc_best.assumptions['reward_redemption_rate']
    ws[f'E{row}'] = "%"
    
    for col in ['A', 'B', 'C', 'D', 'E']:
        ws.column_dimensions[col].width = 22
    
    # ===== INVESTOR METRICS SHEET =====
    ws = ws_investor
    row = 1
    
    ws[f'A{row}'] = "INVESTOR METRICS"
    style_header(ws[f'A{row}'], "FF6600")
    ws.merge_cells(f'A{row}:C{row}')
    row += 2
    
    ws[f'A{row}'] = "UNIT ECONOMICS"
    style_section(ws[f'A{row}'])
    ws.merge_cells(f'A{row}:C{row}')
    row += 1
    
    ws[f'A{row}'] = "Customer Acquisition Cost (CAC)"
    ws[f'B{row}'] = 2.00
    style_editable(ws[f'B{row}'])
    ws[f'B{row}'].number_format = '$0.00'
    ws[f'C{row}'] = "Assumed marketing cost per user"
    cac_cell = f'B{row}'
    row += 1
    
    ws[f'A{row}'] = "Avg User Lifetime (months)"
    ws[f'B{row}'] = 12
    style_editable(ws[f'B{row}'])
    ws[f'C{row}'] = "months"
    lifetime_cell = f'B{row}'
    row += 2
    
    ws[f'A{row}'] = "LIFETIME VALUE (LTV) CALCULATION"
    style_section(ws[f'A{row}'])
    ws.merge_cells(f'A{row}:C{row}')
    row += 1
    
    ws[f'A{row}'] = "LTV from Bag Sales"
    ws[f'B{row}'] = f"={lifetime_cell}*(Assumptions!{bag_price_cell})"
    ws[f'B{row}'].number_format = '$0.00'
    bag_ltv_row = row
    row += 1
    
    ws[f'A{row}'] = "LTV from Ad Revenue"
    ws[f'B{row}'] = f"={lifetime_cell}*Assumptions!{monthly_views_cell}*Assumptions!{margin_per_view_cell}"
    ws[f'B{row}'].number_format = '$0.00'
    ad_ltv_row = row
    row += 1
    
    ws[f'A{row}'] = "Total LTV"
    ws[f'B{row}'] = f"=B{bag_ltv_row}+B{ad_ltv_row}"
    ws[f'B{row}'].number_format = '$0.00'
    ws[f'B{row}'].font = Font(bold=True, size=12)
    ltv_row = row
    row += 2
    
    ws[f'A{row}'] = "KEY RATIOS"
    style_section(ws[f'A{row}'])
    ws.merge_cells(f'A{row}:C{row}')
    row += 1
    
    ws[f'A{row}'] = "LTV:CAC Ratio"
    ws[f'B{row}'] = f"=B{ltv_row}/{cac_cell}"
    ws[f'B{row}'].number_format = '0.0"x"'
    ws[f'B{row}'].font = Font(bold=True, size=11)
    ws[f'C{row}'] = "Target: >3x"
    row += 1
    
    ws[f'A{row}'] = "Monthly Contribution Margin"
    ws[f'B{row}'] = f"=(Assumptions!{bag_price_cell})+(Assumptions!{monthly_views_cell}*Assumptions!{margin_per_view_cell})"
    ws[f'B{row}'].number_format = '$0.00'
    monthly_contrib_row = row
    row += 1
    
    ws[f'A{row}'] = "Payback Period (months)"
    ws[f'B{row}'] = f"={cac_cell}/B{monthly_contrib_row}"
    ws[f'B{row}'].number_format = '0.0" months"'
    ws[f'B{row}'].font = Font(bold=True, size=11)
    ws[f'C{row}'] = "Target: <6 months"
    row += 2
    
    ws[f'A{row}'] = "ANNUAL PROJECTIONS"
    style_section(ws[f'A{row}'])
    ws.merge_cells(f'A{row}:C{row}')
    row += 1
    
    ws[f'A{row}'] = "Annual Revenue (12 months)"
    ws[f'B{row}'] = f"=Calculations!B{total_revenue_row}*12"
    ws[f'B{row}'].number_format = '$#,##0'
    row += 1
    
    ws[f'A{row}'] = "Annual Margin (12 months)"
    ws[f'B{row}'] = f"=Calculations!B{total_margin_row}*12"
    ws[f'B{row}'].number_format = '$#,##0'
    ws[f'B{row}'].font = Font(bold=True, color="00AA00")
    row += 1
    
    ws[f'A{row}'] = "Annual Store Value Created"
    ws[f'B{row}'] = f"=Calculations!B{transactions_row}*12*Assumptions!{avg_basket_cell}*1.15"
    ws[f'B{row}'].number_format = '$#,##0'
    
    for col in ['A', 'B', 'C']:
        ws.column_dimensions[col].width = 30
    
    # ===== SENSITIVITY ANALYSIS SHEET =====
    ws = ws_sensitivity
    row = 1
    
    ws[f'A{row}'] = "SENSITIVITY ANALYSIS"
    style_header(ws[f'A{row}'], "2E75B6")
    ws.merge_cells(f'A{row}:F{row}')
    row += 2
    
    # Sensitivity: Monthly Ad Views
    ws[f'A{row}'] = "SENSITIVITY: Monthly Ad Views per User"
    style_section(ws[f'A{row}'])
    ws.merge_cells(f'A{row}:F{row}')
    row += 1
    
    ws[f'A{row}'] = "Ad Views"
    ws[f'B{row}'] = "Revenue"
    ws[f'C{row}'] = "Margin"
    ws[f'D{row}'] = "Margin %"
    ws[f'E{row}'] = "User Rewards"
    for col in ['A', 'B', 'C', 'D', 'E']:
        style_section(ws[f'{col}{row}'])
    row += 1
    
    # Run sensitivity
    calc = CommerceRewardCalculator()
    ad_view_values = [3, 6, 9, 12, 15, 18, 20, 25]
    sensitivity_results = calc.calculate_sensitivity_analysis('avg_monthly_ad_views', ad_view_values)
    
    for result in sensitivity_results:
        ws[f'A{row}'] = result['value']
        ws[f'B{row}'] = result['total_revenue']
        ws[f'B{row}'].number_format = '$#,##0'
        ws[f'C{row}'] = result['total_margin']
        ws[f'C{row}'].number_format = '$#,##0'
        ws[f'D{row}'] = result['margin_pct'] / 100
        ws[f'D{row}'].number_format = '0.0%'
        ws[f'E{row}'] = result['user_rewards']
        ws[f'E{row}'].number_format = '$0.00'
        row += 1
    
    row += 2
    
    # Sensitivity: Bag Price
    ws[f'A{row}'] = "SENSITIVITY: Bag Retail Price"
    style_section(ws[f'A{row}'])
    ws.merge_cells(f'A{row}:F{row}')
    row += 1
    
    ws[f'A{row}'] = "Bag Price"
    ws[f'B{row}'] = "Revenue"
    ws[f'C{row}'] = "Margin"
    ws[f'D{row}'] = "Margin %"
    ws[f'E{row}'] = "User Net Cost"
    for col in ['A', 'B', 'C', 'D', 'E']:
        style_section(ws[f'{col}{row}'])
    row += 1
    
    bag_price_values = [0.25, 0.30, 0.35, 0.40, 0.45, 0.50, 0.60, 0.75]
    for price in bag_price_values:
        calc.assumptions['bag_retail_price'] = price
        summary = calc.calculate_monthly_summary()
        per_bag = calc.calculate_revenue_per_bag()
        
        ws[f'A{row}'] = price
        ws[f'A{row}'].number_format = '$0.00'
        ws[f'B{row}'] = summary['total_revenue']
        ws[f'B{row}'].number_format = '$#,##0'
        ws[f'C{row}'] = summary['total_margin']
        ws[f'C{row}'].number_format = '$#,##0'
        ws[f'D{row}'] = summary['margin_percentage'] / 100
        ws[f'D{row}'].number_format = '0.0%'
        ws[f'E{row}'] = per_bag['user_net_cost']
        ws[f'E{row}'].number_format = '$0.00'
        row += 1
    
    row += 2
    
    # Sensitivity: Bags Sold
    ws[f'A{row}'] = "SENSITIVITY: Monthly Volume (Bags Sold)"
    style_section(ws[f'A{row}'])
    ws.merge_cells(f'A{row}:F{row}')
    row += 1
    
    ws[f'A{row}'] = "Bags Sold"
    ws[f'B{row}'] = "Revenue"
    ws[f'C{row}'] = "Margin"
    ws[f'D{row}'] = "Margin %"
    ws[f'E{row}'] = "Store Value"
    for col in ['A', 'B', 'C', 'D', 'E']:
        style_section(ws[f'{col}{row}'])
    row += 1
    
    calc = CommerceRewardCalculator()  # Reset
    volume_values = [5000, 7500, 10000, 15000, 20000, 25000, 50000]
    for volume in volume_values:
        calc.assumptions['bags_sold_per_month'] = volume
        summary = calc.calculate_monthly_summary()
        store = calc.calculate_store_value()
        
        ws[f'A{row}'] = volume
        ws[f'A{row}'].number_format = '#,##0'
        ws[f'B{row}'] = summary['total_revenue']
        ws[f'B{row}'].number_format = '$#,##0'
        ws[f'C{row}'] = summary['total_margin']
        ws[f'C{row}'].number_format = '$#,##0'
        ws[f'D{row}'] = summary['margin_percentage'] / 100
        ws[f'D{row}'].number_format = '0.0%'
        ws[f'E{row}'] = store['total_monthly_store_value']
        ws[f'E{row}'].number_format = '$#,##0'
        row += 1
    
    for col in ['A', 'B', 'C', 'D', 'E', 'F']:
        ws.column_dimensions[col].width = 18
    
    # Save
    output_file = 'kshipra_commerce_reward_model.xlsx'
    wb.save(output_file)
    
    print(f"\n✅ Excel calculator created: {output_file}")
    print(f"\n📊 SHEETS:")
    print(f"   • Dashboard: Key metrics summary")
    print(f"   • Assumptions: Editable inputs (YELLOW cells)")
    print(f"   • Calculations: Detailed formulas")
    print(f"   • Scenarios: Best/Base/Worst case comparison")
    print(f"   • Investor Metrics: LTV, CAC, payback period")
    print(f"   • Sensitivity Analysis: Impact of varying key inputs")
    print(f"\n✨ All metrics use formulas - change yellow cells to recalculate!")

if __name__ == "__main__":
    create_commerce_reward_excel()
