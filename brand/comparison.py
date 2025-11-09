"""
Campaign Comparison Tool

Compare multiple brand campaigns and identify optimization opportunities.
"""

import csv
from typing import Dict, List, Any
from pathlib import Path


class CampaignComparison:
    """Compare and analyze multiple brand campaigns."""
    
    def __init__(self):
        """Initialize the comparison tool."""
        self.campaigns = []
    
    def add_campaign(self, campaign_data: Dict[str, Any]) -> None:
        """
        Add a campaign to the comparison.
        
        Args:
            campaign_data: Dictionary containing campaign metrics
        """
        self.campaigns.append(campaign_data)
    
    def load_from_csv(self, csv_file: str) -> None:
        """
        Load campaigns from CSV file.
        
        Args:
            csv_file: Path to CSV file
        """
        self.campaigns = []
        try:
            with open(csv_file, 'r') as file:
                reader = csv.DictReader(file)
                for row in reader:
                    campaign = {
                        'campaign_id': row['campaign_id'],
                        'campaign_name': row['campaign_name'],
                        'start_date': row['start_date'],
                        'end_date': row['end_date'],
                        'total_investment': float(row['total_investment']),
                        'total_revenue': float(row['total_revenue']),
                        'new_customers': int(row['new_customers']),
                        'impressions': int(row['impressions']),
                        'engagements': int(row['engagements']),
                        'clicks': int(row['clicks']),
                        'conversions': int(row['conversions']),
                        'ad_spend': float(row['ad_spend']),
                        'creative_cost': float(row['creative_cost']),
                        'platform_fee': float(row['platform_fee'])
                    }
                    self.campaigns.append(campaign)
        except Exception as e:
            print(f"Error loading data: {e}")
    
    def calculate_metrics(self, campaign: Dict) -> Dict[str, float]:
        """
        Calculate all relevant metrics for a campaign.
        
        Args:
            campaign: Campaign data dictionary
            
        Returns:
            Dictionary with calculated metrics
        """
        investment = campaign['total_investment']
        revenue = campaign['total_revenue']
        customers = campaign['new_customers']
        impressions = campaign['impressions']
        engagements = campaign['engagements']
        clicks = campaign['clicks']
        conversions = campaign['conversions']
        
        metrics = {
            'roi': ((revenue - investment) / investment * 100) if investment > 0 else 0,
            'profit': revenue - investment,
            'cac': investment / customers if customers > 0 else 0,
            'engagement_rate': (engagements / impressions * 100) if impressions > 0 else 0,
            'ctr': (clicks / impressions * 100) if impressions > 0 else 0,
            'cvr': (conversions / clicks * 100) if clicks > 0 else 0,
            'cpc': investment / clicks if clicks > 0 else 0,
            'cpa': investment / conversions if conversions > 0 else 0,
            'revenue_per_customer': revenue / customers if customers > 0 else 0,
            'roas': revenue / investment if investment > 0 else 0
        }
        
        return metrics
    
    def compare_all(self) -> List[Dict[str, Any]]:
        """
        Compare all campaigns with calculated metrics.
        
        Returns:
            List of campaigns with their metrics
        """
        results = []
        for campaign in self.campaigns:
            metrics = self.calculate_metrics(campaign)
            result = {
                'campaign_name': campaign['campaign_name'],
                'campaign_id': campaign['campaign_id'],
                **metrics
            }
            results.append(result)
        
        return results
    
    def find_best_performers(self) -> Dict[str, Any]:
        """
        Identify the best performing campaigns across different metrics.
        
        Returns:
            Dictionary with best performers for each metric
        """
        if not self.campaigns:
            return {}
        
        comparisons = self.compare_all()
        
        best = {
            'highest_roi': max(comparisons, key=lambda x: x['roi']),
            'highest_profit': max(comparisons, key=lambda x: x['profit']),
            'lowest_cac': min(comparisons, key=lambda x: x['cac']),
            'highest_engagement': max(comparisons, key=lambda x: x['engagement_rate']),
            'highest_ctr': max(comparisons, key=lambda x: x['ctr']),
            'highest_cvr': max(comparisons, key=lambda x: x['cvr']),
            'highest_roas': max(comparisons, key=lambda x: x['roas'])
        }
        
        return best
    
    def find_worst_performers(self) -> Dict[str, Any]:
        """
        Identify the worst performing campaigns that need improvement.
        
        Returns:
            Dictionary with worst performers for each metric
        """
        if not self.campaigns:
            return {}
        
        comparisons = self.compare_all()
        
        worst = {
            'lowest_roi': min(comparisons, key=lambda x: x['roi']),
            'lowest_profit': min(comparisons, key=lambda x: x['profit']),
            'highest_cac': max(comparisons, key=lambda x: x['cac']),
            'lowest_engagement': min(comparisons, key=lambda x: x['engagement_rate']),
            'lowest_ctr': min(comparisons, key=lambda x: x['ctr']),
            'lowest_cvr': min(comparisons, key=lambda x: x['cvr']),
            'lowest_roas': min(comparisons, key=lambda x: x['roas'])
        }
        
        return worst
    
    def calculate_averages(self) -> Dict[str, float]:
        """
        Calculate average metrics across all campaigns.
        
        Returns:
            Dictionary with average values
        """
        if not self.campaigns:
            return {}
        
        comparisons = self.compare_all()
        n = len(comparisons)
        
        averages = {
            'avg_roi': sum(c['roi'] for c in comparisons) / n,
            'avg_profit': sum(c['profit'] for c in comparisons) / n,
            'avg_cac': sum(c['cac'] for c in comparisons) / n,
            'avg_engagement_rate': sum(c['engagement_rate'] for c in comparisons) / n,
            'avg_ctr': sum(c['ctr'] for c in comparisons) / n,
            'avg_cvr': sum(c['cvr'] for c in comparisons) / n,
            'avg_roas': sum(c['roas'] for c in comparisons) / n
        }
        
        return averages
    
    def generate_comparison_report(self) -> str:
        """
        Generate a comprehensive comparison report.
        
        Returns:
            Formatted report string
        """
        if not self.campaigns:
            return "No campaigns to compare"
        
        report = "\n" + "=" * 80 + "\n"
        report += "CAMPAIGN COMPARISON REPORT\n"
        report += "=" * 80 + "\n\n"
        
        # Overall Statistics
        report += "OVERALL STATISTICS\n"
        report += "-" * 80 + "\n"
        total_investment = sum(c['total_investment'] for c in self.campaigns)
        total_revenue = sum(c['total_revenue'] for c in self.campaigns)
        total_profit = total_revenue - total_investment
        total_customers = sum(c['new_customers'] for c in self.campaigns)
        
        report += f"Total Campaigns: {len(self.campaigns)}\n"
        report += f"Total Investment: ${total_investment:,.2f}\n"
        report += f"Total Revenue: ${total_revenue:,.2f}\n"
        report += f"Total Profit: ${total_profit:,.2f}\n"
        report += f"Total New Customers: {total_customers:,}\n"
        report += f"Overall ROI: {(total_profit / total_investment * 100):.2f}%\n\n"
        
        # Averages
        report += "AVERAGE METRICS\n"
        report += "-" * 80 + "\n"
        averages = self.calculate_averages()
        for key, value in averages.items():
            if 'rate' in key or 'roi' in key or 'ctr' in key or 'cvr' in key:
                report += f"{key.replace('_', ' ').title()}: {value:.2f}%\n"
            elif 'roas' in key:
                report += f"{key.replace('_', ' ').title()}: {value:.2f}x\n"
            else:
                report += f"{key.replace('_', ' ').title()}: ${value:.2f}\n"
        report += "\n"
        
        # Best Performers
        report += "TOP PERFORMERS\n"
        report += "-" * 80 + "\n"
        best = self.find_best_performers()
        report += f"🏆 Highest ROI: {best['highest_roi']['campaign_name']} ({best['highest_roi']['roi']:.2f}%)\n"
        report += f"💰 Highest Profit: {best['highest_profit']['campaign_name']} (${best['highest_profit']['profit']:,.2f})\n"
        report += f"💵 Lowest CAC: {best['lowest_cac']['campaign_name']} (${best['lowest_cac']['cac']:.2f})\n"
        report += f"📈 Highest Engagement: {best['highest_engagement']['campaign_name']} ({best['highest_engagement']['engagement_rate']:.2f}%)\n"
        report += f"🎯 Highest CVR: {best['highest_cvr']['campaign_name']} ({best['highest_cvr']['cvr']:.2f}%)\n"
        report += f"📊 Highest ROAS: {best['highest_roas']['campaign_name']} ({best['highest_roas']['roas']:.2f}x)\n\n"
        
        # Needs Improvement
        report += "NEEDS IMPROVEMENT\n"
        report += "-" * 80 + "\n"
        worst = self.find_worst_performers()
        report += f"⚠️  Lowest ROI: {worst['lowest_roi']['campaign_name']} ({worst['lowest_roi']['roi']:.2f}%)\n"
        report += f"⚠️  Lowest Profit: {worst['lowest_profit']['campaign_name']} (${worst['lowest_profit']['profit']:,.2f})\n"
        report += f"⚠️  Highest CAC: {worst['highest_cac']['campaign_name']} (${worst['highest_cac']['cac']:.2f})\n"
        report += f"⚠️  Lowest Engagement: {worst['lowest_engagement']['campaign_name']} ({worst['lowest_engagement']['engagement_rate']:.2f}%)\n\n"
        
        # Recommendations
        report += "RECOMMENDATIONS\n"
        report += "-" * 80 + "\n"
        avg_roi = averages['avg_roi']
        best_roi = best['highest_roi']['roi']
        
        report += f"1. Focus on replicating strategies from '{best['highest_roi']['campaign_name']}'\n"
        report += f"   which achieved {best_roi:.1f}% ROI.\n\n"
        report += f"2. Optimize '{worst['lowest_roi']['campaign_name']}' which has the lowest ROI\n"
        report += f"   of {worst['lowest_roi']['roi']:.1f}%.\n\n"
        report += f"3. Improve CAC for '{worst['highest_cac']['campaign_name']}' (${worst['highest_cac']['cac']:.2f})\n"
        report += f"   by learning from '{best['lowest_cac']['campaign_name']}' (${best['lowest_cac']['cac']:.2f}).\n\n"
        
        report += "=" * 80 + "\n"
        
        return report


def main():
    """Example usage of the comparison tool."""
    current_dir = Path(__file__).parent
    data_file = current_dir / "sample_data.csv"
    
    if not data_file.exists():
        print(f"Sample data file not found at {data_file}")
        return
    
    # Create comparison tool
    comparison = CampaignComparison()
    comparison.load_from_csv(str(data_file))
    
    # Generate report
    print(comparison.generate_comparison_report())
    
    # Show detailed comparison
    print("\nDETAILED METRICS COMPARISON")
    print("=" * 120)
    comparisons = comparison.compare_all()
    print(f"{'Campaign':<20} {'ROI':>8} {'Profit':>12} {'CAC':>8} {'Eng%':>7} {'CTR':>7} {'CVR':>7} {'ROAS':>7}")
    print("-" * 120)
    for c in comparisons:
        print(f"{c['campaign_name'][:20]:<20} "
              f"{c['roi']:>7.1f}% "
              f"${c['profit']:>11,.0f} "
              f"${c['cac']:>7.2f} "
              f"{c['engagement_rate']:>6.1f}% "
              f"{c['ctr']:>6.2f}% "
              f"{c['cvr']:>6.1f}% "
              f"{c['roas']:>6.2f}x")


if __name__ == "__main__":
    main()
