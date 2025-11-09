"""
Investor ROI Calculator

This module calculates ROI metrics for investors in the platform.
"""

from typing import Dict, Any, List
from datetime import datetime


class InvestorROICalculator:
    """Calculate ROI metrics for investor portfolios."""
    
    def __init__(self, investment_data: Dict[str, Any]):
        """
        Initialize the ROI calculator with investment data.
        
        Args:
            investment_data: Dictionary containing investment metrics
        """
        self.investment_data = investment_data
    
    def calculate_roi(self) -> float:
        """
        Calculate basic ROI percentage.
        
        Returns:
            ROI as a percentage
        """
        initial_investment = self.investment_data.get('initial_investment', 0)
        current_value = self.investment_data.get('current_value', 0)
        
        if initial_investment == 0:
            return 0.0
        
        roi = ((current_value - initial_investment) / initial_investment) * 100
        return round(roi, 2)
    
    def calculate_annualized_return(self) -> float:
        """
        Calculate annualized return rate.
        
        Returns:
            Annualized return as a percentage
        """
        initial_investment = self.investment_data.get('initial_investment', 0)
        current_value = self.investment_data.get('current_value', 0)
        years = self.investment_data.get('investment_period_years', 1)
        
        if initial_investment == 0 or years == 0:
            return 0.0
        
        annualized = (((current_value / initial_investment) ** (1 / years)) - 1) * 100
        return round(annualized, 2)
    
    def calculate_irr(self, cash_flows: List[float]) -> float:
        """
        Calculate Internal Rate of Return (IRR).
        
        Args:
            cash_flows: List of cash flows over time
            
        Returns:
            IRR as a percentage (simplified calculation)
        """
        # Simplified IRR calculation - in production, use numpy.irr or similar
        # This is a placeholder for the concept
        if not cash_flows or len(cash_flows) < 2:
            return 0.0
        
        return round(10.0, 2)  # Placeholder value
    
    def calculate_multiple_on_invested_capital(self) -> float:
        """
        Calculate MOIC (Multiple on Invested Capital).
        
        Returns:
            MOIC value
        """
        initial_investment = self.investment_data.get('initial_investment', 0)
        current_value = self.investment_data.get('current_value', 0)
        
        if initial_investment == 0:
            return 0.0
        
        moic = current_value / initial_investment
        return round(moic, 2)
    
    def calculate_profit_margin(self) -> float:
        """
        Calculate profit margin.
        
        Returns:
            Profit margin as a percentage
        """
        initial_investment = self.investment_data.get('initial_investment', 0)
        current_value = self.investment_data.get('current_value', 0)
        
        if current_value == 0:
            return 0.0
        
        profit_margin = ((current_value - initial_investment) / current_value) * 100
        return round(profit_margin, 2)
    
    def get_full_report(self) -> Dict[str, Any]:
        """
        Generate a comprehensive ROI report for investors.
        
        Returns:
            Dictionary with all calculated metrics
        """
        return {
            'roi_percentage': self.calculate_roi(),
            'annualized_return': self.calculate_annualized_return(),
            'moic': self.calculate_multiple_on_invested_capital(),
            'profit_margin': self.calculate_profit_margin(),
            'initial_investment': self.investment_data.get('initial_investment', 0),
            'current_value': self.investment_data.get('current_value', 0),
            'unrealized_gain': self.investment_data.get('current_value', 0) - 
                             self.investment_data.get('initial_investment', 0)
        }


if __name__ == "__main__":
    # Example usage
    sample_investment = {
        'initial_investment': 100000,
        'current_value': 150000,
        'investment_period_years': 2,
        'investment_date': '2023-01-01'
    }
    
    calculator = InvestorROICalculator(sample_investment)
    report = calculator.get_full_report()
    
    print("Investor ROI Report")
    print("=" * 50)
    for key, value in report.items():
        print(f"{key}: {value}")
