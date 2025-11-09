"""
BagBuddy Campaign Comparison Tool

Compare multiple BagBuddy brand campaigns and identify optimization opportunities.
"""

import csv
from typing import Dict, List, Any
from pathlib import Path
from roi_calculator import BrandROICalculator


class BagBuddyCampaignComparison:
    """Compare and analyze multiple BagBuddy brand campaigns."""
    
    def __init__(self):
        """Initialize the comparison tool."""
        self.campaigns = []
        self.reports = []
    
    def load_from_csv(self, csv_file: str) -> None:
        """
        Load campaigns from CSV file.
        
        Args:
            csv_file: Path to CSV file with BagBuddy campaign data
        """
        self.campaigns = []
        self.reports = []
        
        try:
            with open(csv_file, 'r') as file:
                reader = csv.DictReader(file)
                for row in reader:
                    campaign = {
                        'campaign_id': row['campaign_id'],
                        'campaign_name': row['campaign_name'],
                        'num_quarters': int(row['num_quarters']),
                        'scan_rate': float(row['scan_rate']),
                        'conversion_rate': float(row['conversion_rate']),
                        'avg_revenue_per_conversion': float(row['avg_revenue_per_conversion']),
                        'impressions_per_bag': int(row['impressions_per_bag']),
                        'trees_planted': int(row['trees_planted'])
                    }
                    self.campaigns.append(campaign)
                    
                    # Calculate metrics for this campaign
                    calculator = BrandROICalculator(campaign)
                    report = calculator.get_full_report()
                    report['campaign_name'] = campaign['campaign_name']
                    report['campaign_id'] = campaign['campaign_id']
                    self.reports.append(report)
                    
        except Exception as e:
            print(f"Error loading data: {e}")
    
    def generate_summary_table(self) -> str:
        """Generate a summary table of all campaigns."""
        if not self.reports:
            return "No data loaded"
        
        table = "\n" + "=" * 140 + "\n"
        table += "BAGBUDDY CAMPAIGN PERFORMANCE SUMMARY\n"
        table += "=" * 140 + "\n"
        table += f"{'Campaign':<25} {'Cost':>10} {'Bags':>8} {'Scans':>8} {'Conv':>6} {'Sales':>12} {'ROI':>10} {'CPI':>8} {'CPA':>8}\n"
        table += "-" * 140 + "\n"
        
        for report in self.reports:
            table += f"{report['campaign_name'][:25]:<25} "
            table += f"${report['campaign_cost']:>9,.0f} "
            table += f"{report['num_bags_distributed']:>8,} "
            table += f"{report['engagements_scans']:>8,} "
            table += f"{report['conversions_redemptions']:>6,} "
            table += f"${report['total_sales_generated']:>11,.0f} "
            table += f"{report['roi_percent']:>9.1f}% "
            table += f"${report['cost_per_impression']:>7.4f} "
            table += f"${report['cost_per_conversion']:>7.2f}\n"
        
        # Add totals
        total_cost = sum(r['campaign_cost'] for r in self.reports)
        total_bags = sum(r['num_bags_distributed'] for r in self.reports)
        total_scans = sum(r['engagements_scans'] for r in self.reports)
        total_conversions = sum(r['conversions_redemptions'] for r in self.reports)
        total_sales = sum(r['total_sales_generated'] for r in self.reports)
        overall_roi = ((total_sales - total_cost) / total_cost * 100) if total_cost > 0 else 0
        
        table += "-" * 140 + "\n"
        table += f"{'TOTAL':<25} "
        table += f"${total_cost:>9,.0f} "
        table += f"{total_bags:>8,} "
        table += f"{total_scans:>8,} "
        table += f"{total_conversions:>6,} "
        table += f"${total_sales:>11,.0f} "
        table += f"{overall_roi:>9.1f}% "
        table += f"{'N/A':>8} "
        table += f"{'N/A':>8}\n"
        table += "=" * 140 + "\n"
        
        return table
    
    def generate_environmental_summary(self) -> str:
        """Generate environmental impact summary."""
        if not self.reports:
            return "No data loaded"
        
        total_trees = sum(r['environmental_impact']['trees_planted'] for r in self.reports)
        total_plastic_saved = sum(r['environmental_impact']['plastic_bags_saved'] for r in self.reports)
        total_carbon = sum(r['environmental_impact']['carbon_offset_kg'] for r in self.reports)
        total_eco_points = sum(r['environmental_impact']['eco_points'] for r in self.reports)
        
        summary = "\n" + "=" * 80 + "\n"
        summary += "ENVIRONMENTAL IMPACT SUMMARY\n"
        summary += "=" * 80 + "\n"
        summary += f"🌳 Total Trees Planted: {total_trees:,}\n"
        summary += f"♻️  Plastic Bags Saved: {total_plastic_saved:,}\n"
        summary += f"🌍 Carbon Offset: {total_carbon:,} kg CO2\n"
        summary += f"⭐ Total Eco-Points: {total_eco_points:,}\n"
        summary += "=" * 80 + "\n"
        
        return summary
    
    def find_best_performers(self) -> Dict[str, Any]:
        """Identify the best performing campaigns."""
        if not self.reports:
            return {}
        
        best = {
            'highest_roi': max(self.reports, key=lambda x: x['roi_percent']),
            'most_scans': max(self.reports, key=lambda x: x['engagements_scans']),
            'highest_conversion_rate': max(self.reports, key=lambda x: x['conversion_rate_percent']),
            'most_sales': max(self.reports, key=lambda x: x['total_sales_generated']),
            'lowest_cpa': min(self.reports, key=lambda x: x['cost_per_conversion'] if x['cost_per_conversion'] > 0 else float('inf')),
            'most_eco_friendly': max(self.reports, key=lambda x: x['environmental_impact']['trees_planted'])
        }
        
        return best
    
    def generate_insights(self) -> str:
        """Generate insights and recommendations."""
        if not self.reports:
            return "No data loaded"
        
        best = self.find_best_performers()
        
        insights = "\n" + "=" * 80 + "\n"
        insights += "PERFORMANCE INSIGHTS & RECOMMENDATIONS\n"
        insights += "=" * 80 + "\n\n"
        
        insights += "🏆 TOP PERFORMERS:\n"
        insights += "-" * 80 + "\n"
        insights += f"Best ROI: {best['highest_roi']['campaign_name']}\n"
        insights += f"  → {best['highest_roi']['roi_percent']:.1f}% ROI\n\n"
        
        insights += f"Most Scans: {best['most_scans']['campaign_name']}\n"
        insights += f"  → {best['most_scans']['engagements_scans']:,} scans\n\n"
        
        insights += f"Highest Conversion Rate: {best['highest_conversion_rate']['campaign_name']}\n"
        insights += f"  → {best['highest_conversion_rate']['conversion_rate_percent']}%\n\n"
        
        insights += f"Most Sales: {best['most_sales']['campaign_name']}\n"
        insights += f"  → ${best['most_sales']['total_sales_generated']:,.2f}\n\n"
        
        insights += f"Lowest Cost per Conversion: {best['lowest_cpa']['campaign_name']}\n"
        insights += f"  → ${best['lowest_cpa']['cost_per_conversion']:.2f}\n\n"
        
        insights += f"Most Eco-Friendly: {best['most_eco_friendly']['campaign_name']}\n"
        insights += f"  → {best['most_eco_friendly']['environmental_impact']['trees_planted']:,} trees planted\n\n"
        
        # Calculate averages
        avg_scan_rate = sum(r['scan_rate_percent'] for r in self.reports) / len(self.reports)
        avg_conversion_rate = sum(r['conversion_rate_percent'] for r in self.reports) / len(self.reports)
        avg_roi = sum(r['roi_percent'] for r in self.reports) / len(self.reports)
        
        insights += "\n📊 AVERAGES:\n"
        insights += "-" * 80 + "\n"
        insights += f"Average Scan Rate: {avg_scan_rate:.1f}%\n"
        insights += f"Average Conversion Rate: {avg_conversion_rate:.1f}%\n"
        insights += f"Average ROI: {avg_roi:.1f}%\n\n"
        
        insights += "\n💡 RECOMMENDATIONS:\n"
        insights += "-" * 80 + "\n"
        insights += f"1. Study {best['highest_roi']['campaign_name']}'s strategy to replicate\n"
        insights += f"   their {best['highest_roi']['roi_percent']:.1f}% ROI success.\n\n"
        insights += f"2. Improve scan rates for campaigns below {avg_scan_rate:.1f}% average.\n"
        insights += f"   Consider QR code placement and visibility.\n\n"
        insights += f"3. Optimize conversion rates for campaigns below {avg_conversion_rate:.1f}% average.\n"
        insights += f"   Review offer attractiveness and redemption process.\n\n"
        insights += "=" * 80 + "\n"
        
        return insights


def main():
    """Example usage of the BagBuddy comparison tool."""
    current_dir = Path(__file__).parent
    data_file = current_dir / "sample_data.csv"
    
    if not data_file.exists():
        print(f"Sample data file not found at {data_file}")
        return
    
    # Create comparison tool
    comparison = BagBuddyCampaignComparison()
    comparison.load_from_csv(str(data_file))
    
    # Generate reports
    print(comparison.generate_summary_table())
    print(comparison.generate_environmental_summary())
    print(comparison.generate_insights())


if __name__ == "__main__":
    main()
