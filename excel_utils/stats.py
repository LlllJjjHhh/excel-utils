import pandas as pd
import numpy as np
from typing import Dict, Optional


class DataStats:
    """Data statistics and analysis for Excel data"""
    
    def __init__(self, df: pd.DataFrame):
        self.df = df.copy()
    
    def basic_summary(self) -> Dict:
        """Get basic summary statistics"""
        summary = {
            'total_rows': len(self.df),
            'total_columns': len(self.df.columns),
            'memory_usage': f"{self.df.memory_usage(deep=True).sum() / (1024*1024):.2f} MB",
            'numeric_columns': list(self.df.select_dtypes(include=[np.number]).columns),
            'categorical_columns': list(self.df.select_dtypes(include=['object', 'category']).columns),
            'missing_values': self.df.isna().sum().to_dict(),
            'missing_percentage': (self.df.isna().sum() / len(self.df) * 100).round(2).to_dict()
        }
        return summary
    
    def numeric_stats(self) -> Dict:
        """Get detailed statistics for numeric columns"""
        numeric_cols = self.df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) == 0:
            return {}
        
        stats = self.df[numeric_cols].agg(['mean', 'median', 'min', 'max', 'std', 'count']).to_dict()
        return stats
    
    def frequency_analysis(self, column: str, top_n: int = 10) -> Dict:
        """Get frequency analysis for categorical columns"""
        if column not in self.df.columns:
            raise ValueError(f"Column {column} not found")
        
        freq = self.df[column].value_counts().head(top_n)
        return freq.to_dict()
    
    def correlation_matrix(self) -> Optional[pd.DataFrame]:
        """Calculate correlation matrix for numeric columns"""
        numeric_cols = self.df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) < 2:
            return None
        
        return self.df[numeric_cols].corr()
    
    def find_outliers(self, column: str, method: str = 'iqr') -> pd.DataFrame:
        """Find outliers in a numeric column using IQR or Z-score method"""
        if column not in self.df.columns:
            raise ValueError(f"Column {column} not found")
        
        if not pd.api.types.is_numeric_dtype(self.df[column]):
            raise ValueError(f"Column {column} is not numeric")
        
        if method == 'iqr':
            q1 = self.df[column].quantile(0.25)
            q3 = self.df[column].quantile(0.75)
            iqr = q3 - q1
            lower_bound = q1 - 1.5 * iqr
            upper_bound = q3 + 1.5 * iqr
            outliers = self.df[(self.df[column] < lower_bound) | (self.df[column] > upper_bound)]
        elif method == 'zscore':
            z_scores = np.abs((self.df[column] - self.df[column].mean()) / self.df[column].std())
            outliers = self.df[z_scores > 3]
        else:
            raise ValueError("Method must be 'iqr' or 'zscore'")
        
        return outliers
    
    def group_by_analysis(self, group_column: str, value_column: str, agg_func: str = 'mean') -> Dict:
        """Group by analysis for a specific column"""
        if group_column not in self.df.columns or value_column not in self.df.columns:
            raise ValueError("One or more columns not found")
        
        result = self.df.groupby(group_column)[value_column].agg(agg_func).sort_values(ascending=False)
        return result.to_dict()
    
    def export_report(self, output_file: str) -> None:
        """Export a comprehensive statistics report to HTML"""
        summary = self.basic_summary()
        numeric_stats = self.numeric_stats()
        
        # Create HTML report
        html = "<html><head><title>Excel Data Statistics Report</title></head><body>"
        html += "<h1>Excel Data Statistics Report</h1>"
        
        html += "<h2>Basic Information</h2>"
        html += "<ul>"
        html += f"<li>Total Rows: {summary['total_rows']}</li>"
        html += f"<li>Total Columns: {summary['total_columns']}</li>"
        html += f"<li>Memory Usage: {summary['memory_usage']}</li>"
        html += (
            f"<li>Numeric Columns: {', '.join(summary['numeric_columns'])}</li>"
        )
        html += (
            f"<li>Categorical Columns: {', '.join(summary['categorical_columns'])}</li>"
        )
        html += "</ul>"
        
        html += "<h2>Missing Values</h2>"
        html += (
            "<table border='1'>"
            "<tr><th>Column</th><th>Missing Values</th>"
            "<th>Missing Percentage</th></tr>"
        )
        for col in summary['missing_values']:
            pct = summary['missing_percentage'][col]
            html += f"<tr><td>{col}</td><td>{summary['missing_values'][col]}</td>"
            html += f"<td>{pct}%</td></tr>"
        html += "</table>"
        
        if numeric_stats:
            html += "<h2>Numeric Column Statistics</h2>"
            html += (
                "<table border='1'>"
                "<tr><th>Column</th><th>Mean</th><th>Median</th>"
                "<th>Min</th><th>Max</th><th>Std Dev</th></tr>"
            )
            for col, stats in numeric_stats.items():
                html += (
                    f"<tr><td>{col}</td>"
                    f"<td>{stats['mean']:.4f}</td>"
                    f"<td>{stats['median']:.4f}</td>"
                    f"<td>{stats['min']:.4f}</td>"
                    f"<td>{stats['max']:.4f}</td>"
                    f"<td>{stats['std']:.4f}</td></tr>"
                )
            html += "</table>"
        
        html += "</body></html>"
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html)
