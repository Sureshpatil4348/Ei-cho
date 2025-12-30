"""
Comprehensive Portfolio Analysis Sheets Generator
Creates 3 Excel sheets:
1. Strategy Statistics (all pairs data per strategy)
2. Pair Capital Distribution (within each strategy with 5 allocation methods + Sharpe/XIRR)
3. Strategy Capital Distribution (between strategies with 5 allocation methods + Sharpe/XIRR)

IMPORTANT: Initial Balance = Max Drawdown × 2 (for proper capital requirement calculation)
"""

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
import os

# ============================================================================
# DATA LOADING AND RECALCULATION
# ============================================================================

# Scaling factors for strategies tested at higher capital
SCALING_FACTORS = {
    'Gold_Dip': 10,        # Tested at 10x ($10,000 instead of $1,000)
    'RSI_Correlation': 100  # Tested at 100x ($100,000 instead of $1,000)
}

def recalculate_metrics(df, strategy_name=None):
    """
    Recalculate all metrics based on correct Initial Balance = Max Drawdown × 2
    Apply scaling factor if strategy was tested at higher capital
    """
    df = df.copy()
    
    # Apply scaling factor for strategies tested at higher capital
    scale = SCALING_FACTORS.get(strategy_name, 1)
    if scale > 1:
        df['Total_Profit'] = df['Total_Profit'] / scale
        df['Max_Drawdown'] = df['Max_Drawdown'] / scale
        # Also scale other dollar-based columns if they exist
        if 'Initial_Balance' in df.columns:
            df['Initial_Balance'] = df['Initial_Balance'] / scale
        if 'Final_Balance' in df.columns:
            df['Final_Balance'] = df['Final_Balance'] / scale
    
    # Calculate correct Initial Balance = Max Drawdown × 2
    df['Correct_Initial_Balance'] = df['Max_Drawdown'] * 2
    
    # Recalculate Final Balance
    df['Correct_Final_Balance'] = df['Correct_Initial_Balance'] + df['Total_Profit']
    
    # Recalculate Total Return %
    df['Correct_Return_Percent'] = (df['Total_Profit'] / df['Correct_Initial_Balance']) * 100
    
    # Recalculate XIRR (annualized return based on trading period)
    df['Trading_Years'] = df['Trading_Period_Days'] / 365
    
    # XIRR = ((Final/Initial)^(1/years) - 1) × 100
    # Handle edge cases
    df['Correct_XIRR'] = df.apply(
        lambda row: ((row['Correct_Final_Balance'] / row['Correct_Initial_Balance']) ** 
                    (1 / max(row['Trading_Years'], 0.1)) - 1) * 100 
        if row['Correct_Initial_Balance'] > 0 else 0, 
        axis=1
    )
    
    # Store the scaling factor applied
    df['Scale_Factor'] = scale
    
    return df

# Load and recalculate all pair statistics files
strategies_data = {}

# 7th Strategy
df_7th = pd.read_csv('/Users/sureshpatil/Desktop/Portfolio Creation/7th_Strategy_pair_statistics.csv')
df_7th = recalculate_metrics(df_7th, '7th_Strategy')
strategies_data['7th_Strategy'] = df_7th

# Falcon
df_falcon = pd.read_csv('/Users/sureshpatil/Desktop/Portfolio Creation/Falcon_pair_statistics.csv')
df_falcon = recalculate_metrics(df_falcon, 'Falcon')
strategies_data['Falcon'] = df_falcon

# Gold_Dip (Tested at 10x capital - values divided by 10)
df_gold_dip = pd.read_csv('/Users/sureshpatil/Desktop/Portfolio Creation/Gold_Dip_pair_statistics.csv')
df_gold_dip = recalculate_metrics(df_gold_dip, 'Gold_Dip')
strategies_data['Gold_Dip'] = df_gold_dip

# RSI_Correlation (Tested at 100x capital - values divided by 100)
df_rsi_corr = pd.read_csv('/Users/sureshpatil/Desktop/Portfolio Creation/RSI_Correlation_pair_statistics.csv')
df_rsi_corr = recalculate_metrics(df_rsi_corr, 'RSI_Correlation')
strategies_data['RSI_Correlation'] = df_rsi_corr

# RSI_6_Trades
df_rsi_6 = pd.read_csv('/Users/sureshpatil/Desktop/Portfolio Creation/RSI_6_Trades_pair_statistics.csv')
df_rsi_6 = recalculate_metrics(df_rsi_6, 'RSI_6_Trades')
strategies_data['RSI_6_Trades'] = df_rsi_6

# AURUM
df_aurum = pd.read_csv('/Users/sureshpatil/Desktop/Portfolio Creation/AURUM_pair_statistics.csv')
df_aurum = recalculate_metrics(df_aurum, 'AURUM')
strategies_data['AURUM'] = df_aurum

# PairTradingEA
df_pair_trading = pd.read_csv('/Users/sureshpatil/Desktop/Portfolio Creation/PairTradingEA_pair_statistics.csv')
df_pair_trading = recalculate_metrics(df_pair_trading, 'PairTradingEA')
strategies_data['PairTradingEA'] = df_pair_trading

# ============================================================================
# ALLOCATION METHODS
# ============================================================================

def calculate_equal_weight(n_assets):
    """Equal weight allocation"""
    return np.ones(n_assets) / n_assets * 100

def calculate_inverse_volatility_weight(max_drawdowns):
    """Inverse volatility (using max drawdown as volatility proxy)"""
    dd = np.array(max_drawdowns)
    dd = np.where(dd == 0, 0.001, dd)
    inv_vol = 1 / dd
    return inv_vol / inv_vol.sum() * 100

def calculate_sharpe_weight(sharpe_ratios):
    """Sharpe ratio weighted allocation"""
    sharpe = np.array(sharpe_ratios)
    sharpe = np.maximum(sharpe, 0.001)
    return sharpe / sharpe.sum() * 100

def calculate_risk_parity_weight(sharpe_ratios, max_drawdowns):
    """Risk parity - balance risk contribution"""
    sharpe = np.array(sharpe_ratios)
    dd = np.array(max_drawdowns)
    dd = np.where(dd == 0, 0.001, dd)
    
    risk_adj = sharpe / dd
    risk_adj = np.maximum(risk_adj, 0.001)
    return risk_adj / risk_adj.sum() * 100

def calculate_max_sharpe_weight(sharpe_ratios, returns, max_drawdowns):
    """Max Sharpe optimization approximation"""
    sharpe = np.array(sharpe_ratios)
    returns = np.array(returns)
    dd = np.array(max_drawdowns)
    dd = np.where(dd == 0, 0.001, dd)
    
    score = sharpe * returns / dd
    score = np.maximum(score, 0.001)
    return score / score.sum() * 100

def calculate_portfolio_sharpe(weights, sharpe_ratios):
    """Calculate portfolio Sharpe ratio as weighted average"""
    weights = np.array(weights) / 100
    sharpe = np.array(sharpe_ratios)
    portfolio_sharpe = np.sum(weights * sharpe)
    
    # Apply diversification benefit
    n_assets = len(weights[weights > 0.01])
    if n_assets > 1:
        diversification_factor = 1 + 0.1 * np.log(n_assets)
        portfolio_sharpe *= diversification_factor
    
    return portfolio_sharpe

def calculate_portfolio_xirr(weights, xirr_values):
    """Calculate portfolio XIRR as weighted average"""
    weights = np.array(weights) / 100
    xirr = np.array(xirr_values)
    return np.sum(weights * xirr)

def calculate_allocations_and_metrics(df, name_col='Currency_Pair'):
    """Calculate all 5 allocation methods and resulting Sharpe/XIRR for a set of pairs"""
    n_pairs = len(df)
    
    if n_pairs == 0:
        return None
    
    sharpe_ratios = df['Sharpe_Ratio'].values
    max_drawdowns = df['Max_Drawdown'].values
    returns = df['Correct_Return_Percent'].values  # Use corrected returns
    xirr_values = df['Correct_XIRR'].values  # Use corrected XIRR
    
    # Calculate weights for each method
    equal_w = calculate_equal_weight(n_pairs)
    inv_vol_w = calculate_inverse_volatility_weight(max_drawdowns)
    sharpe_w = calculate_sharpe_weight(sharpe_ratios)
    risk_parity_w = calculate_risk_parity_weight(sharpe_ratios, max_drawdowns)
    max_sharpe_w = calculate_max_sharpe_weight(sharpe_ratios, returns, max_drawdowns)
    
    methods = {
        'Equal_Weight': equal_w,
        'Inverse_Volatility': inv_vol_w,
        'Sharpe_Weighted': sharpe_w,
        'Risk_Parity': risk_parity_w,
        'Max_Sharpe': max_sharpe_w
    }
    
    results = {}
    for method_name, weights in methods.items():
        portfolio_sharpe = calculate_portfolio_sharpe(weights, sharpe_ratios)
        portfolio_xirr = calculate_portfolio_xirr(weights, xirr_values)
        
        results[method_name] = {
            'weights': weights,
            'portfolio_sharpe': portfolio_sharpe,
            'portfolio_xirr': portfolio_xirr
        }
    
    return results, df[name_col].values

# ============================================================================
# EXCEL STYLING
# ============================================================================

def style_header(ws, row, start_col, end_col, text=None):
    """Apply header styling"""
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')

def style_subheader(ws, row, start_col, end_col):
    """Apply subheader styling"""
    subheader_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    subheader_font = Font(bold=True, color="FFFFFF", size=10)
    
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = subheader_fill
        cell.font = subheader_font
        cell.alignment = Alignment(horizontal='center', vertical='center')

def style_strategy_header(ws, row, start_col, end_col, color="4472C4"):
    """Apply strategy section header styling"""
    strategy_fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    strategy_font = Font(bold=True, color="FFFFFF", size=12)
    
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = strategy_fill
        cell.font = strategy_font
        cell.alignment = Alignment(horizontal='center', vertical='center')

def style_result_row(ws, row, start_col, end_col, color="E2EFDA"):
    """Apply result row styling"""
    result_fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    result_font = Font(bold=True, size=10)
    
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = result_fill
        cell.font = result_font

def add_border(ws, start_row, end_row, start_col, end_col):
    """Add border to a range"""
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            ws.cell(row=row, column=col).border = thin_border

# ============================================================================
# SHEET 1: STRATEGY STATISTICS
# ============================================================================

def create_sheet1_statistics(wb):
    """Create Sheet 1: All strategy statistics with pairs data"""
    ws = wb.create_sheet("Strategy_Statistics")
    
    row = 1
    
    # Main title
    ws.cell(row=row, column=1, value="COMPREHENSIVE STRATEGY STATISTICS")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=17)
    ws.cell(row=row, column=1).font = Font(bold=True, size=16, color="1F4E79")
    ws.cell(row=row, column=1).alignment = Alignment(horizontal='center')
    row += 1
    
    ws.cell(row=row, column=1, value="Note: Initial Capital = Max Drawdown × 2 (for proper risk management)")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=17)
    ws.cell(row=row, column=1).font = Font(italic=True, size=10, color="C00000")
    row += 1
    
    ws.cell(row=row, column=1, value="IMPORTANT: Gold_Dip values divided by 10 (tested at 10x capital). RSI_Correlation values divided by 100 (tested at 100x capital).")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=17)
    ws.cell(row=row, column=1).font = Font(italic=True, size=10, color="0066CC")
    row += 2
    
    strategy_colors = ["4472C4", "ED7D31", "70AD47", "9E480E", "5B9BD5", "7030A0", "C00000"]
    
    for idx, (strategy_name, df) in enumerate(strategies_data.items()):
        # Strategy section header
        ws.cell(row=row, column=1, value=f"Strategy: {strategy_name}")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=17)
        style_strategy_header(ws, row, 1, 17, strategy_colors[idx % len(strategy_colors)])
        row += 1
        
        # Column headers
        headers = ['Currency_Pair', 'Sharpe_Ratio', 'Total_Trades', 'Win_Rate_%', 
                   'Total_Profit', 'Max_Drawdown', 'Initial_Capital', 'Final_Balance',
                   'Return_%', 'XIRR_%', 'Profit_Factor', 'Trading_Days', 
                   'Trading_Years', 'Start_Date', 'End_Date', 'Winning', 'Losing']
        
        for col_idx, header in enumerate(headers, 1):
            ws.cell(row=row, column=col_idx, value=header)
        style_subheader(ws, row, 1, len(headers))
        row += 1
        
        start_data_row = row
        
        # Data rows
        for _, pair_row in df.iterrows():
            ws.cell(row=row, column=1, value=pair_row['Currency_Pair'])
            ws.cell(row=row, column=2, value=round(pair_row['Sharpe_Ratio'], 4))
            ws.cell(row=row, column=3, value=int(pair_row['Total_Trades']))
            ws.cell(row=row, column=4, value=round(pair_row['Win_Rate_Percent'], 2))
            ws.cell(row=row, column=5, value=round(pair_row['Total_Profit'], 2))
            ws.cell(row=row, column=6, value=round(pair_row['Max_Drawdown'], 2))
            ws.cell(row=row, column=7, value=round(pair_row['Correct_Initial_Balance'], 2))
            ws.cell(row=row, column=8, value=round(pair_row['Correct_Final_Balance'], 2))
            ws.cell(row=row, column=9, value=round(pair_row['Correct_Return_Percent'], 2))
            ws.cell(row=row, column=10, value=round(pair_row['Correct_XIRR'], 2))
            ws.cell(row=row, column=11, value=round(pair_row['Profit_Factor'], 4))
            ws.cell(row=row, column=12, value=int(pair_row.get('Trading_Period_Days', 0)))
            ws.cell(row=row, column=13, value=round(pair_row['Trading_Years'], 2))
            ws.cell(row=row, column=14, value=str(pair_row.get('Start_Date', 'N/A')))
            ws.cell(row=row, column=15, value=str(pair_row.get('End_Date', 'N/A')))
            ws.cell(row=row, column=16, value=int(pair_row['Winning_Trades']))
            ws.cell(row=row, column=17, value=int(pair_row['Losing_Trades']))
            row += 1
        
        # Strategy Summary Row
        ws.cell(row=row, column=1, value="STRATEGY TOTAL")
        ws.cell(row=row, column=2, value=round(df['Sharpe_Ratio'].mean(), 4))
        ws.cell(row=row, column=3, value=int(df['Total_Trades'].sum()))
        total_trades = df['Total_Trades'].sum()
        total_wins = df['Winning_Trades'].sum()
        ws.cell(row=row, column=4, value=round(total_wins/total_trades*100 if total_trades > 0 else 0, 2))
        ws.cell(row=row, column=5, value=round(df['Total_Profit'].sum(), 2))
        ws.cell(row=row, column=6, value=round(df['Max_Drawdown'].sum(), 2))
        ws.cell(row=row, column=7, value=round(df['Correct_Initial_Balance'].sum(), 2))
        ws.cell(row=row, column=8, value=round(df['Correct_Final_Balance'].sum(), 2))
        total_init = df['Correct_Initial_Balance'].sum()
        total_profit = df['Total_Profit'].sum()
        ws.cell(row=row, column=9, value=round(total_profit/total_init*100 if total_init > 0 else 0, 2))
        ws.cell(row=row, column=10, value=round(df['Correct_XIRR'].mean(), 2))
        ws.cell(row=row, column=11, value=round(df['Profit_Factor'].mean(), 4))
        ws.cell(row=row, column=12, value="-")
        ws.cell(row=row, column=13, value=round(df['Trading_Years'].mean(), 2))
        ws.cell(row=row, column=14, value="-")
        ws.cell(row=row, column=15, value="-")
        ws.cell(row=row, column=16, value=int(df['Winning_Trades'].sum()))
        ws.cell(row=row, column=17, value=int(df['Losing_Trades'].sum()))
        style_result_row(ws, row, 1, 17, "FFF2CC")
        
        add_border(ws, start_data_row - 1, row, 1, 17)
        row += 3
    
    # Adjust column widths
    for col in range(1, 18):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.column_dimensions['A'].width = 25
    
    return ws

# ============================================================================
# SHEET 2: PAIR CAPITAL DISTRIBUTION (WITHIN STRATEGIES)
# ============================================================================

def create_sheet2_pair_allocation(wb):
    """Create Sheet 2: Capital distribution between pairs within each strategy"""
    ws = wb.create_sheet("Pair_Capital_Distribution")
    
    row = 1
    
    # Main title
    ws.cell(row=row, column=1, value="PAIR CAPITAL DISTRIBUTION WITHIN STRATEGIES")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
    ws.cell(row=row, column=1).font = Font(bold=True, size=16, color="1F4E79")
    ws.cell(row=row, column=1).alignment = Alignment(horizontal='center')
    row += 1
    
    ws.cell(row=row, column=1, value="(If 100% capital allocated to this strategy, how should it be distributed among pairs?)")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
    ws.cell(row=row, column=1).font = Font(italic=True, size=10, color="666666")
    row += 1
    
    ws.cell(row=row, column=1, value="Note: Initial Capital = Max Drawdown × 2 | Returns and XIRR calculated based on correct capital")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
    ws.cell(row=row, column=1).font = Font(italic=True, size=10, color="C00000")
    row += 1
    
    ws.cell(row=row, column=1, value="IMPORTANT: Gold_Dip values divided by 10 (tested at 10x capital). RSI_Correlation values divided by 100 (tested at 100x capital).")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
    ws.cell(row=row, column=1).font = Font(italic=True, size=10, color="0066CC")
    row += 2
    
    strategy_colors = ["4472C4", "ED7D31", "70AD47", "9E480E", "5B9BD5", "7030A0", "C00000"]
    
    for idx, (strategy_name, df) in enumerate(strategies_data.items()):
        if len(df) == 0:
            continue
            
        # Strategy section header
        ws.cell(row=row, column=1, value=f"STRATEGY: {strategy_name} ({len(df)} pairs)")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
        style_strategy_header(ws, row, 1, 12, strategy_colors[idx % len(strategy_colors)])
        row += 1
        
        # Calculate allocations
        results, pair_names = calculate_allocations_and_metrics(df, 'Currency_Pair')
        
        # Column headers for allocation table
        headers = ['Currency_Pair', 'Equal_%', 'Inv_Vol_%', 'Sharpe_%', 
                   'Risk_Parity_%', 'Max_Sharpe_%', 'Sharpe_Ratio', 'Return_%', 'XIRR_%', 
                   'Max_DD', 'Initial_Cap', 'Profit']
        
        for col_idx, header in enumerate(headers, 1):
            ws.cell(row=row, column=col_idx, value=header)
        style_subheader(ws, row, 1, len(headers))
        row += 1
        
        start_data_row = row
        
        # Data rows for each pair
        for i, pair_name in enumerate(pair_names):
            ws.cell(row=row, column=1, value=pair_name)
            ws.cell(row=row, column=2, value=round(results['Equal_Weight']['weights'][i], 2))
            ws.cell(row=row, column=3, value=round(results['Inverse_Volatility']['weights'][i], 2))
            ws.cell(row=row, column=4, value=round(results['Sharpe_Weighted']['weights'][i], 2))
            ws.cell(row=row, column=5, value=round(results['Risk_Parity']['weights'][i], 2))
            ws.cell(row=row, column=6, value=round(results['Max_Sharpe']['weights'][i], 2))
            ws.cell(row=row, column=7, value=round(df.iloc[i]['Sharpe_Ratio'], 4))
            ws.cell(row=row, column=8, value=round(df.iloc[i]['Correct_Return_Percent'], 2))
            ws.cell(row=row, column=9, value=round(df.iloc[i]['Correct_XIRR'], 2))
            ws.cell(row=row, column=10, value=round(df.iloc[i]['Max_Drawdown'], 2))
            ws.cell(row=row, column=11, value=round(df.iloc[i]['Correct_Initial_Balance'], 2))
            ws.cell(row=row, column=12, value=round(df.iloc[i]['Total_Profit'], 2))
            row += 1
        
        # Total row
        ws.cell(row=row, column=1, value="TOTAL")
        ws.cell(row=row, column=2, value=100.0)
        ws.cell(row=row, column=3, value=100.0)
        ws.cell(row=row, column=4, value=100.0)
        ws.cell(row=row, column=5, value=100.0)
        ws.cell(row=row, column=6, value=100.0)
        ws.cell(row=row, column=10, value=round(df['Max_Drawdown'].sum(), 2))
        ws.cell(row=row, column=11, value=round(df['Correct_Initial_Balance'].sum(), 2))
        ws.cell(row=row, column=12, value=round(df['Total_Profit'].sum(), 2))
        style_result_row(ws, row, 1, 12, "D9E1F2")
        row += 1
        
        add_border(ws, start_data_row - 1, row - 1, 1, 12)
        row += 1
        
        # Performance metrics section
        ws.cell(row=row, column=1, value="PORTFOLIO PERFORMANCE BY ALLOCATION METHOD:")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        ws.cell(row=row, column=1).font = Font(bold=True, size=11, color="1F4E79")
        row += 1
        
        # Method headers
        method_headers = ['Allocation_Method', 'Portfolio_Sharpe_Ratio', 'Portfolio_XIRR_%']
        for col_idx, header in enumerate(method_headers, 1):
            ws.cell(row=row, column=col_idx, value=header)
        style_subheader(ws, row, 1, 3)
        row += 1
        
        start_metric_row = row
        
        # Performance for each method
        for method_name in ['Equal_Weight', 'Inverse_Volatility', 'Sharpe_Weighted', 'Risk_Parity', 'Max_Sharpe']:
            ws.cell(row=row, column=1, value=method_name)
            ws.cell(row=row, column=2, value=round(results[method_name]['portfolio_sharpe'], 4))
            ws.cell(row=row, column=3, value=round(results[method_name]['portfolio_xirr'], 2))
            row += 1
        
        add_border(ws, start_metric_row - 1, row - 1, 1, 3)
        style_result_row(ws, row - 1, 1, 3, "E2EFDA")
        
        row += 3
    
    # Adjust column widths
    for col in range(1, 13):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.column_dimensions['A'].width = 28
    
    return ws

# ============================================================================
# SHEET 3: STRATEGY CAPITAL DISTRIBUTION
# ============================================================================

def create_sheet3_strategy_allocation(wb):
    """Create Sheet 3: Capital distribution between strategies"""
    ws = wb.create_sheet("Strategy_Capital_Distribution")
    
    row = 1
    
    # Main title
    ws.cell(row=row, column=1, value="STRATEGY CAPITAL DISTRIBUTION")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
    ws.cell(row=row, column=1).font = Font(bold=True, size=16, color="1F4E79")
    ws.cell(row=row, column=1).alignment = Alignment(horizontal='center')
    row += 1
    
    ws.cell(row=row, column=1, value="(If you have 100% capital, how should it be distributed among strategies?)")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
    ws.cell(row=row, column=1).font = Font(italic=True, size=10, color="666666")
    row += 1
    
    ws.cell(row=row, column=1, value="Note: Initial Capital = Max Drawdown × 2 | Returns and XIRR calculated based on correct capital")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
    ws.cell(row=row, column=1).font = Font(italic=True, size=10, color="C00000")
    row += 1
    
    ws.cell(row=row, column=1, value="IMPORTANT: Gold_Dip values divided by 10 (tested at 10x capital). RSI_Correlation values divided by 100 (tested at 100x capital).")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
    ws.cell(row=row, column=1).font = Font(italic=True, size=10, color="0066CC")
    row += 2
    
    # Create strategy-level metrics
    strategies = []
    for strategy_name, df in strategies_data.items():
        strategy_info = {
            'Strategy': strategy_name,
            'Num_Pairs': len(df),
            'Avg_Sharpe_Ratio': df['Sharpe_Ratio'].mean(),
            'Total_Profit': df['Total_Profit'].sum(),
            'Total_Initial_Capital': df['Correct_Initial_Balance'].sum(),
            'Correct_Return_Percent': (df['Total_Profit'].sum() / df['Correct_Initial_Balance'].sum() * 100) 
                                       if df['Correct_Initial_Balance'].sum() > 0 else 0,
            'Max_Drawdown': df['Max_Drawdown'].sum(),
            'Correct_XIRR': df['Correct_XIRR'].mean(),
            'Win_Rate': df['Win_Rate_Percent'].mean(),
            'Total_Trades': df['Total_Trades'].sum(),
            'Trading_Years': df['Trading_Years'].mean()
        }
        strategies.append(strategy_info)
    
    strategy_df = pd.DataFrame(strategies)
    
    # Section header
    ws.cell(row=row, column=1, value="STRATEGY WEIGHTS BY ALLOCATION METHOD")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
    style_strategy_header(ws, row, 1, 12, "1F4E79")
    row += 1
    
    # Calculate allocations for strategies
    n_strategies = len(strategy_df)
    sharpe_ratios = strategy_df['Avg_Sharpe_Ratio'].values
    max_drawdowns = strategy_df['Max_Drawdown'].values
    returns = strategy_df['Correct_Return_Percent'].values
    xirr_values = strategy_df['Correct_XIRR'].values
    
    equal_w = calculate_equal_weight(n_strategies)
    inv_vol_w = calculate_inverse_volatility_weight(max_drawdowns)
    sharpe_w = calculate_sharpe_weight(sharpe_ratios)
    risk_parity_w = calculate_risk_parity_weight(sharpe_ratios, max_drawdowns)
    max_sharpe_w = calculate_max_sharpe_weight(sharpe_ratios, returns, max_drawdowns)
    
    methods = {
        'Equal_Weight': equal_w,
        'Inverse_Volatility': inv_vol_w,
        'Sharpe_Weighted': sharpe_w,
        'Risk_Parity': risk_parity_w,
        'Max_Sharpe': max_sharpe_w
    }
    
    # Column headers
    headers = ['Strategy', 'Pairs', 'Sharpe', 'Return_%', 'XIRR_%', 'Capital_Req',
               'Equal_%', 'Inv_Vol_%', 'Sharpe_%', 'Risk_Parity_%', 'Max_Sharpe_%', 'Profit']
    
    for col_idx, header in enumerate(headers, 1):
        ws.cell(row=row, column=col_idx, value=header)
    style_subheader(ws, row, 1, len(headers))
    row += 1
    
    start_data_row = row
    
    # Data rows for each strategy
    for i, strat_row in strategy_df.iterrows():
        ws.cell(row=row, column=1, value=strat_row['Strategy'])
        ws.cell(row=row, column=2, value=int(strat_row['Num_Pairs']))
        ws.cell(row=row, column=3, value=round(strat_row['Avg_Sharpe_Ratio'], 4))
        ws.cell(row=row, column=4, value=round(strat_row['Correct_Return_Percent'], 2))
        ws.cell(row=row, column=5, value=round(strat_row['Correct_XIRR'], 2))
        ws.cell(row=row, column=6, value=round(strat_row['Total_Initial_Capital'], 2))
        ws.cell(row=row, column=7, value=round(equal_w[i], 2))
        ws.cell(row=row, column=8, value=round(inv_vol_w[i], 2))
        ws.cell(row=row, column=9, value=round(sharpe_w[i], 2))
        ws.cell(row=row, column=10, value=round(risk_parity_w[i], 2))
        ws.cell(row=row, column=11, value=round(max_sharpe_w[i], 2))
        ws.cell(row=row, column=12, value=round(strat_row['Total_Profit'], 2))
        row += 1
    
    # Total row
    ws.cell(row=row, column=1, value="TOTAL")
    ws.cell(row=row, column=2, value=int(strategy_df['Num_Pairs'].sum()))
    ws.cell(row=row, column=6, value=round(strategy_df['Total_Initial_Capital'].sum(), 2))
    ws.cell(row=row, column=7, value=100.0)
    ws.cell(row=row, column=8, value=100.0)
    ws.cell(row=row, column=9, value=100.0)
    ws.cell(row=row, column=10, value=100.0)
    ws.cell(row=row, column=11, value=100.0)
    ws.cell(row=row, column=12, value=round(strategy_df['Total_Profit'].sum(), 2))
    style_result_row(ws, row, 1, 12, "D9E1F2")
    
    add_border(ws, start_data_row - 1, row, 1, 12)
    row += 3
    
    # Performance metrics section
    ws.cell(row=row, column=1, value="FINAL PORTFOLIO PERFORMANCE BY ALLOCATION METHOD")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    style_strategy_header(ws, row, 1, 6, "70AD47")
    row += 1
    
    ws.cell(row=row, column=1, value="(If you allocate your capital according to this method, here's your expected performance)")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    ws.cell(row=row, column=1).font = Font(italic=True, size=10, color="666666")
    row += 1
    
    # Method headers
    method_headers = ['Allocation_Method', 'Portfolio_Sharpe_Ratio', 'Portfolio_XIRR_%', 
                      'Weighted_Return_%', 'Recommendation']
    for col_idx, header in enumerate(method_headers, 1):
        ws.cell(row=row, column=col_idx, value=header)
    style_subheader(ws, row, 1, 5)
    row += 1
    
    start_metric_row = row
    
    # Calculate and display performance for each method
    best_sharpe = 0
    best_method = ""
    
    for method_name, weights in methods.items():
        portfolio_sharpe = calculate_portfolio_sharpe(weights, sharpe_ratios)
        portfolio_xirr = calculate_portfolio_xirr(weights, xirr_values)
        weighted_return = np.sum((weights/100) * returns)
        
        ws.cell(row=row, column=1, value=method_name)
        ws.cell(row=row, column=2, value=round(portfolio_sharpe, 4))
        ws.cell(row=row, column=3, value=round(portfolio_xirr, 2))
        ws.cell(row=row, column=4, value=round(weighted_return, 2))
        
        if portfolio_sharpe > best_sharpe:
            best_sharpe = portfolio_sharpe
            best_method = method_name
        
        row += 1
    
    # Add recommendations
    rec_row = start_metric_row
    for method_name in methods.keys():
        if method_name == best_method:
            ws.cell(row=rec_row, column=5, value="★ BEST SHARPE")
            ws.cell(row=rec_row, column=5).font = Font(bold=True, color="006600")
        rec_row += 1
    
    add_border(ws, start_metric_row - 1, row - 1, 1, 5)
    row += 3
    
    # Summary box
    ws.cell(row=row, column=1, value="SUMMARY & RECOMMENDATION")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    style_strategy_header(ws, row, 1, 6, "C00000")
    row += 1
    
    ws.cell(row=row, column=1, value=f"Best Allocation Method: {best_method}")
    ws.cell(row=row, column=1).font = Font(bold=True, size=12)
    row += 1
    
    ws.cell(row=row, column=1, value=f"Expected Portfolio Sharpe Ratio: {round(best_sharpe, 4)}")
    row += 1
    
    # Total capital required
    total_capital = strategy_df['Total_Initial_Capital'].sum()
    ws.cell(row=row, column=1, value=f"Total Capital Required (MDD×2): ${total_capital:,.2f}")
    row += 2
    
    # Blended approach
    blended_weights = (equal_w + inv_vol_w + sharpe_w + risk_parity_w + max_sharpe_w) / 5
    blended_sharpe = calculate_portfolio_sharpe(blended_weights, sharpe_ratios)
    blended_xirr = calculate_portfolio_xirr(blended_weights, xirr_values)
    
    ws.cell(row=row, column=1, value=f"Blended Approach (Avg of 5 methods) - Sharpe: {round(blended_sharpe, 4)}, XIRR: {round(blended_xirr, 2)}%")
    row += 2
    
    # Blended weights table
    ws.cell(row=row, column=1, value="BLENDED WEIGHTS (Average of 5 Methods)")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    style_subheader(ws, row, 1, 4)
    row += 1
    
    ws.cell(row=row, column=1, value="Strategy")
    ws.cell(row=row, column=2, value="Blended_%")
    ws.cell(row=row, column=3, value=f"Capital (of ${total_capital:,.0f})")
    ws.cell(row=row, column=4, value="Min Capital Required")
    style_subheader(ws, row, 1, 4)
    row += 1
    
    for i, strat_row in strategy_df.iterrows():
        ws.cell(row=row, column=1, value=strat_row['Strategy'])
        ws.cell(row=row, column=2, value=round(blended_weights[i], 2))
        ws.cell(row=row, column=3, value=round(blended_weights[i]/100 * total_capital, 2))
        ws.cell(row=row, column=4, value=round(strat_row['Total_Initial_Capital'], 2))
        row += 1
    
    # Adjust column widths
    for col in range(1, 13):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['E'].width = 16
    
    return ws

# ============================================================================
# MAIN EXECUTION
# ============================================================================

def main():
    # Print verification
    print("=" * 80)
    print("VERIFICATION: Recalculated metrics using Initial Capital = Max Drawdown × 2")
    print("=" * 80)
    
    for strategy_name, df in strategies_data.items():
        print(f"\n{strategy_name}:")
        for _, row in df.iterrows():
            print(f"  {row['Currency_Pair']}:")
            print(f"    Max DD: ${row['Max_Drawdown']:.2f} → Initial Capital: ${row['Correct_Initial_Balance']:.2f}")
            print(f"    Total Profit: ${row['Total_Profit']:.2f}")
            print(f"    Return: {row['Correct_Return_Percent']:.2f}%")
            print(f"    XIRR: {row['Correct_XIRR']:.2f}%")
    
    # Create workbook
    wb = Workbook()
    wb.remove(wb.active)
    
    print("\n" + "=" * 80)
    print("Creating Excel sheets...")
    print("=" * 80)
    
    print("\nCreating Sheet 1: Strategy Statistics...")
    create_sheet1_statistics(wb)
    
    print("Creating Sheet 2: Pair Capital Distribution...")
    create_sheet2_pair_allocation(wb)
    
    print("Creating Sheet 3: Strategy Capital Distribution...")
    create_sheet3_strategy_allocation(wb)
    
    # Save workbook
    output_path = '/Users/sureshpatil/Desktop/Portfolio Creation/Portfolio_Analysis_3Sheets.xlsx'
    wb.save(output_path)
    print(f"\n✓ Workbook saved to: {output_path}")
    print("\nSheets created:")
    print("  1. Strategy_Statistics - All pairs data with CORRECT capital calculations")
    print("  2. Pair_Capital_Distribution - Capital allocation within strategies (5 methods + Sharpe/XIRR)")
    print("  3. Strategy_Capital_Distribution - Capital allocation between strategies (5 methods + Sharpe/XIRR)")

if __name__ == "__main__":
    main()
