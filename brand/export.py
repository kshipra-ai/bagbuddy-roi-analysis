"""
Data Export Module

Export ROI analysis results to various formats.
"""

import csv
import json
from typing import Dict, List, Any
from pathlib import Path
from datetime import datetime


class ROIDataExporter:
    """Export ROI analysis data to different formats."""
    
    def __init__(self, output_dir: str = None):
        """
        Initialize the exporter.
        
        Args:
            output_dir: Directory to save exported files (defaults to current directory)
        """
        self.output_dir = Path(output_dir) if output_dir else Path.cwd()
        self.output_dir.mkdir(exist_ok=True)
    
    def export_to_csv(self, data: List[Dict[str, Any]], filename: str) -> str:
        """
        Export data to CSV file.
        
        Args:
            data: List of dictionaries containing campaign data
            filename: Name of the output file (without extension)
            
        Returns:
            Path to the created file
        """
        if not data:
            raise ValueError("No data to export")
        
        output_file = self.output_dir / f"{filename}.csv"
        
        # Get all unique keys from all dictionaries
        fieldnames = set()
        for item in data:
            fieldnames.update(item.keys())
        fieldnames = sorted(list(fieldnames))
        
        with open(output_file, 'w', newline='') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(data)
        
        return str(output_file)
    
    def export_to_json(self, data: Any, filename: str, pretty: bool = True) -> str:
        """
        Export data to JSON file.
        
        Args:
            data: Data to export (can be dict, list, etc.)
            filename: Name of the output file (without extension)
            pretty: Whether to format JSON with indentation
            
        Returns:
            Path to the created file
        """
        output_file = self.output_dir / f"{filename}.json"
        
        with open(output_file, 'w') as jsonfile:
            if pretty:
                json.dump(data, jsonfile, indent=2)
            else:
                json.dump(data, jsonfile)
        
        return str(output_file)
    
    def export_comparison_report(self, comparison_data: List[Dict], 
                                 best_performers: Dict, 
                                 averages: Dict,
                                 filename: str = None) -> str:
        """
        Export a complete comparison report.
        
        Args:
            comparison_data: List of campaign comparisons
            best_performers: Dictionary of best performers
            averages: Dictionary of average metrics
            filename: Optional custom filename
            
        Returns:
            Path to the created file
        """
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"roi_comparison_report_{timestamp}"
        
        report = {
            'generated_at': datetime.now().isoformat(),
            'summary': {
                'total_campaigns': len(comparison_data),
                'averages': averages
            },
            'best_performers': best_performers,
            'campaigns': comparison_data
        }
        
        return self.export_to_json(report, filename)
    
    def export_summary_text(self, content: str, filename: str) -> str:
        """
        Export text summary to a file.
        
        Args:
            content: Text content to export
            filename: Name of the output file (without extension)
            
        Returns:
            Path to the created file
        """
        output_file = self.output_dir / f"{filename}.txt"
        
        with open(output_file, 'w') as textfile:
            textfile.write(content)
        
        return str(output_file)
    
    def export_campaign_report(self, campaign_data: Dict[str, Any], 
                              metrics: Dict[str, Any],
                              filename: str = None) -> str:
        """
        Export a single campaign report.
        
        Args:
            campaign_data: Original campaign data
            metrics: Calculated metrics
            filename: Optional custom filename
            
        Returns:
            Path to the created file
        """
        if not filename:
            campaign_name = campaign_data.get('campaign_name', 'unknown').replace(' ', '_')
            timestamp = datetime.now().strftime("%Y%m%d")
            filename = f"campaign_report_{campaign_name}_{timestamp}"
        
        report = {
            'generated_at': datetime.now().isoformat(),
            'campaign_info': campaign_data,
            'metrics': metrics
        }
        
        return self.export_to_json(report, filename)
    
    def create_excel_friendly_csv(self, data: List[Dict[str, Any]], 
                                  filename: str) -> str:
        """
        Export data to Excel-friendly CSV with formatted numbers.
        
        Args:
            data: List of dictionaries containing campaign data
            filename: Name of the output file (without extension)
            
        Returns:
            Path to the created file
        """
        if not data:
            raise ValueError("No data to export")
        
        output_file = self.output_dir / f"{filename}.csv"
        
        # Format numbers for Excel
        formatted_data = []
        for item in data:
            formatted_item = {}
            for key, value in item.items():
                # Format percentages
                if any(term in key.lower() for term in ['rate', 'roi', 'margin', 'ctr', 'cvr']):
                    formatted_item[key] = f"{value}%" if isinstance(value, (int, float)) else value
                # Format currency
                elif any(term in key.lower() for term in ['cost', 'revenue', 'investment', 'profit', 'cac', 'cpc', 'cpa']):
                    formatted_item[key] = f"${value:,.2f}" if isinstance(value, (int, float)) else value
                # Format numbers with commas
                elif isinstance(value, int) and value > 999:
                    formatted_item[key] = f"{value:,}"
                else:
                    formatted_item[key] = value
            formatted_data.append(formatted_item)
        
        return self.export_to_csv(formatted_data, filename)


def main():
    """Example usage of the data exporter."""
    from roi_calculator import BrandROICalculator
    from comparison import CampaignComparison
    
    # Example: Export single campaign report
    campaign_data = {
        'campaign_name': 'Summer Sale 2025',
        'total_investment': 10000,
        'total_revenue': 25000,
        'new_customers': 150,
        'impressions': 250000,
        'engagements': 12500,
        'clicks': 5000,
        'conversions': 300
    }
    
    calculator = BrandROICalculator(campaign_data)
    metrics = calculator.get_full_report()
    
    exporter = ROIDataExporter()
    
    # Export campaign report
    json_file = exporter.export_campaign_report(campaign_data, metrics)
    print(f"✅ Campaign report exported to: {json_file}")
    
    # Load and export comparison data
    current_dir = Path(__file__).parent
    data_file = current_dir / "sample_data.csv"
    
    if data_file.exists():
        comparison = CampaignComparison()
        comparison.load_from_csv(str(data_file))
        
        comparisons = comparison.compare_all()
        best = comparison.find_best_performers()
        averages = comparison.calculate_averages()
        
        # Export comparison report
        report_file = exporter.export_comparison_report(comparisons, best, averages)
        print(f"✅ Comparison report exported to: {report_file}")
        
        # Export Excel-friendly CSV
        excel_file = exporter.create_excel_friendly_csv(comparisons, "campaigns_for_excel")
        print(f"✅ Excel-friendly CSV exported to: {excel_file}")
        
        # Export text summary
        text_report = comparison.generate_comparison_report()
        text_file = exporter.export_summary_text(text_report, "comparison_summary")
        print(f"✅ Text summary exported to: {text_file}")
    else:
        print("⚠️  sample_data.csv not found. Skipping comparison export.")


if __name__ == "__main__":
    main()
