"""
Brand ROI Dashboard

Interactive command-line dashboard for brand ROI analysis.
"""

import os
import sys
from pathlib import Path
from roi_calculator import BrandROICalculator
from comparison import CampaignComparison
from visualization import BrandROIVisualizer
from export import ROIDataExporter


def clear_screen():
    """Clear the terminal screen."""
    os.system('cls' if os.name == 'nt' else 'clear')


def print_header():
    """Print dashboard header."""
    print("\n" + "=" * 80)
    print(" " * 25 + "BRAND ROI ANALYSIS DASHBOARD")
    print("=" * 80 + "\n")


def print_menu():
    """Print main menu options."""
    print("Please select an option:\n")
    print("  1. Analyze Single Campaign")
    print("  2. Compare All Campaigns")
    print("  3. View Performance Visualizations")
    print("  4. Generate Reports")
    print("  5. Export Data")
    print("  6. View Top Performers")
    print("  7. Load Custom Data File")
    print("  8. Exit")
    print("\n" + "-" * 80)


def analyze_single_campaign():
    """Analyze a single campaign with user input."""
    clear_screen()
    print_header()
    print("SINGLE CAMPAIGN ANALYSIS\n")
    
    try:
        campaign_name = input("Campaign Name: ")
        investment = float(input("Total Investment ($): "))
        revenue = float(input("Total Revenue ($): "))
        customers = int(input("New Customers: "))
        impressions = int(input("Impressions: "))
        engagements = int(input("Engagements: "))
        clicks = int(input("Clicks: "))
        conversions = int(input("Conversions: "))
        
        campaign = {
            'campaign_name': campaign_name,
            'total_investment': investment,
            'total_revenue': revenue,
            'new_customers': customers,
            'impressions': impressions,
            'engagements': engagements,
            'clicks': clicks,
            'conversions': conversions
        }
        
        calculator = BrandROICalculator(campaign)
        report = calculator.get_full_report()
        
        print("\n" + "=" * 80)
        print(f"RESULTS: {campaign_name}")
        print("=" * 80 + "\n")
        
        print("FINANCIAL METRICS")
        print("-" * 80)
        print(f"ROI: {report['roi_percentage']}%")
        print(f"ROAS: {report['return_on_ad_spend']}x")
        print(f"Profit: ${report['total_profit']:,.2f}")
        print(f"Profit Margin: {report['profit_margin']}%")
        print(f"Grade: {report['performance_grade']}\n")
        
        print("CUSTOMER METRICS")
        print("-" * 80)
        print(f"CAC: ${report['customer_acquisition_cost']}")
        print(f"Revenue per Customer: ${report['revenue_per_customer']}")
        print(f"Break-even Customers: {report['break_even_customers']}\n")
        
        print("ENGAGEMENT METRICS")
        print("-" * 80)
        print(f"Engagement Rate: {report['engagement_rate']}%")
        print(f"Click-Through Rate: {report['click_through_rate']}%")
        print(f"Conversion Rate: {report['conversion_rate']}%")
        print(f"CPC: ${report['cost_per_click']}")
        print(f"CPA: ${report['cost_per_acquisition']}\n")
        
    except ValueError:
        print("\n❌ Invalid input. Please enter numeric values.")
    except Exception as e:
        print(f"\n❌ Error: {e}")


def compare_campaigns(data_file):
    """Compare all campaigns from data file."""
    clear_screen()
    print_header()
    
    if not data_file.exists():
        print("❌ No data file loaded. Please load a CSV file first (Option 7).\n")
        return
    
    comparison = CampaignComparison()
    comparison.load_from_csv(str(data_file))
    
    print(comparison.generate_comparison_report())


def view_visualizations(data_file):
    """Display performance visualizations."""
    clear_screen()
    print_header()
    
    if not data_file.exists():
        print("❌ No data file loaded. Please load a CSV file first (Option 7).\n")
        return
    
    viz = BrandROIVisualizer(str(data_file))
    
    print(viz.generate_summary_table())
    print(viz.generate_performance_insights())
    print(viz.generate_ascii_bar_chart('roi'))


def generate_reports(data_file):
    """Generate and save reports."""
    clear_screen()
    print_header()
    print("GENERATE REPORTS\n")
    
    if not data_file.exists():
        print("❌ No data file loaded. Please load a CSV file first (Option 7).\n")
        return
    
    comparison = CampaignComparison()
    comparison.load_from_csv(str(data_file))
    
    exporter = ROIDataExporter()
    
    try:
        # Generate comparison report
        comparisons = comparison.compare_all()
        best = comparison.find_best_performers()
        averages = comparison.calculate_averages()
        
        report_file = exporter.export_comparison_report(comparisons, best, averages)
        print(f"✅ Comparison report saved to: {report_file}")
        
        # Generate text summary
        text_report = comparison.generate_comparison_report()
        text_file = exporter.export_summary_text(text_report, "comparison_summary")
        print(f"✅ Text summary saved to: {text_file}")
        
        # Generate Excel-friendly CSV
        excel_file = exporter.create_excel_friendly_csv(comparisons, "campaigns_for_excel")
        print(f"✅ Excel-friendly CSV saved to: {excel_file}\n")
        
    except Exception as e:
        print(f"❌ Error generating reports: {e}\n")


def export_data_menu(data_file):
    """Export data in various formats."""
    clear_screen()
    print_header()
    print("EXPORT DATA\n")
    
    if not data_file.exists():
        print("❌ No data file loaded. Please load a CSV file first (Option 7).\n")
        return
    
    print("Select export format:")
    print("  1. JSON (detailed)")
    print("  2. CSV (Excel-friendly)")
    print("  3. Text Report")
    print("  4. All formats")
    
    choice = input("\nChoice: ")
    
    comparison = CampaignComparison()
    comparison.load_from_csv(str(data_file))
    exporter = ROIDataExporter()
    
    try:
        comparisons = comparison.compare_all()
        
        if choice == '1' or choice == '4':
            best = comparison.find_best_performers()
            averages = comparison.calculate_averages()
            json_file = exporter.export_comparison_report(comparisons, best, averages)
            print(f"✅ JSON export: {json_file}")
        
        if choice == '2' or choice == '4':
            csv_file = exporter.create_excel_friendly_csv(comparisons, "campaigns_export")
            print(f"✅ CSV export: {csv_file}")
        
        if choice == '3' or choice == '4':
            text_report = comparison.generate_comparison_report()
            text_file = exporter.export_summary_text(text_report, "report_export")
            print(f"✅ Text export: {text_file}")
        
        print()
    except Exception as e:
        print(f"❌ Error exporting data: {e}\n")


def view_top_performers(data_file):
    """Display top performing campaigns."""
    clear_screen()
    print_header()
    print("TOP PERFORMERS\n")
    
    if not data_file.exists():
        print("❌ No data file loaded. Please load a CSV file first (Option 7).\n")
        return
    
    comparison = CampaignComparison()
    comparison.load_from_csv(str(data_file))
    
    best = comparison.find_best_performers()
    
    print("🏆 BEST PERFORMERS")
    print("-" * 80)
    print(f"\nHighest ROI:")
    print(f"  Campaign: {best['highest_roi']['campaign_name']}")
    print(f"  ROI: {best['highest_roi']['roi']:.2f}%\n")
    
    print(f"Highest Profit:")
    print(f"  Campaign: {best['highest_profit']['campaign_name']}")
    print(f"  Profit: ${best['highest_profit']['profit']:,.2f}\n")
    
    print(f"Lowest CAC:")
    print(f"  Campaign: {best['lowest_cac']['campaign_name']}")
    print(f"  CAC: ${best['lowest_cac']['cac']:.2f}\n")
    
    print(f"Highest Engagement Rate:")
    print(f"  Campaign: {best['highest_engagement']['campaign_name']}")
    print(f"  Engagement Rate: {best['highest_engagement']['engagement_rate']:.2f}%\n")
    
    print(f"Highest Conversion Rate:")
    print(f"  Campaign: {best['highest_cvr']['campaign_name']}")
    print(f"  CVR: {best['highest_cvr']['cvr']:.2f}%\n")
    
    print(f"Highest ROAS:")
    print(f"  Campaign: {best['highest_roas']['campaign_name']}")
    print(f"  ROAS: {best['highest_roas']['roas']:.2f}x\n")


def load_custom_file():
    """Load a custom data file."""
    clear_screen()
    print_header()
    print("LOAD CUSTOM DATA FILE\n")
    
    file_path = input("Enter the path to your CSV file: ").strip()
    
    if not file_path:
        return None
    
    file_path = Path(file_path)
    
    if not file_path.exists():
        print(f"\n❌ File not found: {file_path}\n")
        return None
    
    if file_path.suffix.lower() != '.csv':
        print(f"\n❌ File must be a CSV file\n")
        return None
    
    print(f"\n✅ File loaded successfully: {file_path}\n")
    return file_path


def main():
    """Main dashboard loop."""
    # Default to sample data
    current_dir = Path(__file__).parent
    data_file = current_dir / "sample_data.csv"
    
    while True:
        clear_screen()
        print_header()
        
        if data_file.exists():
            print(f"📊 Current Data File: {data_file.name}\n")
        else:
            print("⚠️  No data file loaded\n")
        
        print_menu()
        
        choice = input("\nEnter your choice (1-8): ").strip()
        
        if choice == '1':
            analyze_single_campaign()
        elif choice == '2':
            compare_campaigns(data_file)
        elif choice == '3':
            view_visualizations(data_file)
        elif choice == '4':
            generate_reports(data_file)
        elif choice == '5':
            export_data_menu(data_file)
        elif choice == '6':
            view_top_performers(data_file)
        elif choice == '7':
            custom_file = load_custom_file()
            if custom_file:
                data_file = custom_file
        elif choice == '8':
            clear_screen()
            print("\n👋 Thank you for using Brand ROI Analysis Dashboard!\n")
            sys.exit(0)
        else:
            print("\n❌ Invalid choice. Please select 1-8.")
        
        input("\nPress Enter to continue...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        clear_screen()
        print("\n\n👋 Dashboard closed. Goodbye!\n")
        sys.exit(0)
