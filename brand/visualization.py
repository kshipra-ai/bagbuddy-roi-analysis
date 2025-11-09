"""
Brand ROI Visualization Module

Generate charts and graphs for brand campaign ROI analysis.
"""

import csv
from typing import Dict, List, Any
from pathlib import Path


class BrandROIVisualizer:
    """Create visualizations for brand ROI metrics."""
    
    def __init__(self, data_file: str = None):
        """
        Initialize the visualizer.
        
        Args:
            data_file: Path to CSV file with campaign data
        """
        self.data_file = data_file
        self.campaigns = []
        
        if data_file:
            self.load_data(data_file)
    
    def load_data(self, csv_file: str) -> None:
        """
        Load campaign data from CSV file.
        
        Args:
            csv_file: Path to CSV file
        """
        self.campaigns = []
        try:
            with open(csv_file, 'r') as file:
                reader = csv.DictReader(file)
                for row in reader:
                    # Convert numeric fields
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
        except FileNotFoundError:
            print(f"Error: File {csv_file} not found")
        except Exception as e:
            print(f"Error loading data: {e}")
    
    def calculate_campaign_roi(self, campaign: Dict) -> float:
        """Calculate ROI for a single campaign."""
        investment = campaign['total_investment']
        revenue = campaign['total_revenue']
        if investment == 0:
            return 0.0
        return ((revenue - investment) / investment) * 100
    
    def generate_ascii_bar_chart(self, metric: str = 'roi') -> str:
        """
        Generate ASCII bar chart for comparison.
        
        Args:
            metric: Metric to visualize ('roi', 'revenue', 'conversions')
            
        Returns:
            ASCII art bar chart as string
        """
        if not self.campaigns:
            return "No data loaded"
        
        chart = f"\n{metric.upper()} Comparison\n"
        chart += "=" * 70 + "\n\n"
        
        # Calculate values
        values = []
        for campaign in self.campaigns:
            if metric == 'roi':
                value = self.calculate_campaign_roi(campaign)
            elif metric == 'revenue':
                value = campaign['total_revenue']
            elif metric == 'conversions':
                value = campaign['conversions']
            else:
                value = campaign.get(metric, 0)
            values.append((campaign['campaign_name'], value))
        
        # Find max value for scaling
        max_value = max([v[1] for v in values]) if values else 1
        max_bar_length = 50
        
        # Generate bars
        for name, value in values:
            bar_length = int((value / max_value) * max_bar_length) if max_value > 0 else 0
            bar = "█" * bar_length
            
            if metric == 'roi':
                chart += f"{name[:20]:20} | {bar} {value:.1f}%\n"
            else:
                chart += f"{name[:20]:20} | {bar} {value:,.0f}\n"
        
        return chart
    
    def generate_summary_table(self) -> str:
        """
        Generate a summary table of all campaigns.
        
        Returns:
            ASCII table with campaign summaries
        """
        if not self.campaigns:
            return "No data loaded"
        
        table = "\nCampaign Performance Summary\n"
        table += "=" * 120 + "\n"
        table += f"{'Campaign':<20} {'Investment':>12} {'Revenue':>12} {'ROI':>10} {'Customers':>10} {'CVR':>8} {'CAC':>10}\n"
        table += "-" * 120 + "\n"
        
        for campaign in self.campaigns:
            roi = self.calculate_campaign_roi(campaign)
            cac = campaign['total_investment'] / campaign['new_customers'] if campaign['new_customers'] > 0 else 0
            cvr = (campaign['conversions'] / campaign['clicks'] * 100) if campaign['clicks'] > 0 else 0
            
            table += f"{campaign['campaign_name'][:20]:<20} "
            table += f"${campaign['total_investment']:>11,.0f} "
            table += f"${campaign['total_revenue']:>11,.0f} "
            table += f"{roi:>9.1f}% "
            table += f"{campaign['new_customers']:>10,} "
            table += f"{cvr:>7.1f}% "
            table += f"${cac:>9.2f}\n"
        
        # Add totals
        total_investment = sum(c['total_investment'] for c in self.campaigns)
        total_revenue = sum(c['total_revenue'] for c in self.campaigns)
        total_customers = sum(c['new_customers'] for c in self.campaigns)
        overall_roi = ((total_revenue - total_investment) / total_investment * 100) if total_investment > 0 else 0
        overall_cac = total_investment / total_customers if total_customers > 0 else 0
        
        table += "-" * 120 + "\n"
        table += f"{'TOTAL':<20} "
        table += f"${total_investment:>11,.0f} "
        table += f"${total_revenue:>11,.0f} "
        table += f"{overall_roi:>9.1f}% "
        table += f"{total_customers:>10,} "
        table += f"{'N/A':>8} "
        table += f"${overall_cac:>9.2f}\n"
        
        return table
    
    def generate_performance_insights(self) -> str:
        """
        Generate insights based on campaign performance.
        
        Returns:
            String with key insights
        """
        if not self.campaigns:
            return "No data loaded"
        
        insights = "\nPerformance Insights\n"
        insights += "=" * 70 + "\n\n"
        
        # Best ROI
        best_roi_campaign = max(self.campaigns, 
                               key=lambda c: self.calculate_campaign_roi(c))
        best_roi = self.calculate_campaign_roi(best_roi_campaign)
        insights += f"🏆 Best ROI: {best_roi_campaign['campaign_name']} ({best_roi:.1f}%)\n\n"
        
        # Highest Revenue
        best_revenue_campaign = max(self.campaigns, 
                                   key=lambda c: c['total_revenue'])
        insights += f"💰 Highest Revenue: {best_revenue_campaign['campaign_name']} "
        insights += f"(${best_revenue_campaign['total_revenue']:,.0f})\n\n"
        
        # Most Customers
        best_customer_campaign = max(self.campaigns, 
                                    key=lambda c: c['new_customers'])
        insights += f"👥 Most New Customers: {best_customer_campaign['campaign_name']} "
        insights += f"({best_customer_campaign['new_customers']:,})\n\n"
        
        # Best Engagement Rate
        engagement_rates = [(c, (c['engagements'] / c['impressions'] * 100) if c['impressions'] > 0 else 0) 
                           for c in self.campaigns]
        best_engagement = max(engagement_rates, key=lambda x: x[1])
        insights += f"📈 Best Engagement Rate: {best_engagement[0]['campaign_name']} "
        insights += f"({best_engagement[1]:.2f}%)\n\n"
        
        # Lowest CAC
        cac_values = [(c, c['total_investment'] / c['new_customers'] if c['new_customers'] > 0 else float('inf')) 
                     for c in self.campaigns]
        best_cac = min(cac_values, key=lambda x: x[1])
        if best_cac[1] != float('inf'):
            insights += f"💵 Lowest CAC: {best_cac[0]['campaign_name']} "
            insights += f"(${best_cac[1]:.2f})\n\n"
        
        return insights


def main():
    """Example usage of the visualizer."""
    # Get the path to sample_data.csv
    current_dir = Path(__file__).parent
    data_file = current_dir / "sample_data.csv"
    
    if not data_file.exists():
        print(f"Sample data file not found at {data_file}")
        return
    
    # Create visualizer
    viz = BrandROIVisualizer(str(data_file))
    
    # Generate visualizations
    print(viz.generate_summary_table())
    print(viz.generate_performance_insights())
    print(viz.generate_ascii_bar_chart('roi'))
    print(viz.generate_ascii_bar_chart('revenue'))
    print(viz.generate_ascii_bar_chart('conversions'))


if __name__ == "__main__":
    main()
