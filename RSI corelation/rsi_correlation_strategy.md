# RSI Correlation Trading Strategy

## Overview
This trading strategy is based on the RSI (Relative Strength Index) indicator correlation between two currency pairs. The strategy utilizes both positive and negative correlation scenarios to identify trading opportunities and employs a martingale-style position sizing approach to manage losses.

## Strategy Components

### 1. Core Concept
The strategy monitors two currency pairs simultaneously and executes trades based on their RSI correlation:
- **Positive Correlation Mode**: Pairs move in opposite RSI zones (one oversold, one overbought)
- **Negative Correlation Mode**: Pairs move in the same RSI zone (both oversold or both overbought)

### 2. Configuration Parameters

#### Correlation Settings
- **Correlation Type**: Choose between Positive or Negative correlation
- **Pair 1**: First currency pair (e.g., EURUSD)
- **Pair 2**: Second currency pair (e.g., GBPUSD)

#### RSI Levels
- **Oversold Level**: Default 30 (RSI below this level indicates oversold)
- **Overbought Level**: Default 70 (RSI above this level indicates overbought)

#### Position Sizing
- **Lot Size Pair 1**: Independent lot size for the first pair
- **Lot Size Pair 2**: Independent lot size for the second pair

#### Risk Management
- **Take Profit**: Target profit in dollars (e.g., $300)
- **Initial Loss Threshold**: Loss level that triggers position averaging (e.g., $300)

## Trading Logic

### Positive Correlation Mode

#### Entry Conditions
- **Pair 1** is in the **oversold zone** (RSI ≤ 30) AND **Pair 2** is in the **overbought zone** (RSI ≥ 70), OR
- **Pair 1** is in the **overbought zone** (RSI ≥ 70) AND **Pair 2** is in the **oversold zone** (RSI ≤ 30)

#### Trade Execution
- **BUY** the pair that is in the oversold zone
- **SELL** the pair that is in the overbought zone

**Example:**
- GBPUSD RSI = 28 (oversold)
- EURUSD RSI = 72 (overbought)
- **Action**: BUY GBPUSD, SELL EURUSD

### Negative Correlation Mode

#### Entry Conditions
- **Both pairs** are in the **oversold zone** (RSI ≤ 30), OR
- **Both pairs** are in the **overbought zone** (RSI ≥ 70)

#### Trade Execution
- If both pairs are oversold: **BUY both pairs**
- If both pairs are overbought: **SELL both pairs**

**Example:**
- GBPUSD RSI = 27 (oversold)
- EURUSD RSI = 29 (oversold)
- **Action**: BUY GBPUSD, BUY EURUSD

## Position Management

### Averaging Down Strategy (Martingale Approach)

The strategy employs a position averaging technique based on floating loss thresholds:

1. **Initial Trade**: Execute trades based on entry conditions
2. **First Averaging**: When combined floating loss reaches $300, repeat the exact same trades (same direction, same lot sizes)
3. **Second Averaging**: When combined floating loss reaches $600, repeat the trades again
4. **Third Averaging**: When combined floating loss reaches $1,200, repeat the trades again
5. **Continue Pattern**: Keep doubling the loss threshold ($2,400, $4,800, etc.) and repeating trades

### Exit Strategy

#### Take Profit Exit
- Monitor the **combined P&L** of all trades in the current sequence
- When the combined P&L reaches the **Take Profit target** (e.g., $300 profit), close **ALL trades** in the sequence simultaneously
- Reset and wait for new entry signals

#### Trade Sequence Example

**Scenario**: Initial Take Profit = $300, Loss Threshold = $300

| Stage | Combined Floating Loss | Action | Trades Open |
|-------|------------------------|--------|-------------|
| 1 | $0 | Initial entry | 2 trades |
| 2 | -$300 | Add positions | 4 trades |
| 3 | -$600 | Add positions | 6 trades |
| 4 | -$1,200 | Add positions | 8 trades |
| 5 | +$300 | Close ALL trades | 0 trades |

## Entry Restrictions

### Active Sequence Protection
- If there is an **active trade sequence** running on the selected pair combination, **NO new trades** will be initiated
- New entry signals are **ignored** until the current sequence is closed
- This prevents multiple overlapping sequences and maintains risk control

### Signal Validation
- Entry signals are only valid when:
  - No active sequence exists for the pair combination
  - RSI levels meet the correlation criteria
  - Both pairs satisfy the entry conditions simultaneously

## Risk Considerations

### High-Risk Elements
1. **Martingale Approach**: Doubling down on losing positions can lead to significant drawdowns
2. **No Stop Loss**: Strategy relies on averaging and eventual mean reversion
3. **Capital Requirements**: Requires substantial capital to sustain multiple averaging levels
4. **Correlation Risk**: Pair correlations can change over time, invalidating the strategy premise

### Risk Mitigation
- Use appropriate position sizing relative to account size
- Set maximum number of averaging levels
- Monitor pair correlation regularly
- Ensure sufficient margin availability
- Consider implementing a maximum loss cut-off level

## Example Walkthrough

### Positive Correlation Example

**Configuration:**
- Pair 1: EURUSD, Lot Size: 0.10
- Pair 2: GBPUSD, Lot Size: 0.10
- Oversold: 30, Overbought: 70
- Take Profit: $300
- Loss Threshold: $300

**Trade Sequence:**

1. **Initial Signal**: EURUSD RSI = 72, GBPUSD RSI = 28
   - Action: SELL EURUSD 0.10 lots, BUY GBPUSD 0.10 lots

2. **Market moves against us**: Combined P&L = -$300
   - Action: SELL EURUSD 0.10 lots, BUY GBPUSD 0.10 lots (2nd set)

3. **Still losing**: Combined P&L = -$600
   - Action: SELL EURUSD 0.10 lots, BUY GBPUSD 0.10 lots (3rd set)

4. **Market reverses**: Combined P&L = +$300
   - Action: Close ALL 6 trades (3 EURUSD sells, 3 GBPUSD buys)
   - Net Profit: $300

### Negative Correlation Example

**Configuration:**
- Pair 1: EURUSD, Lot Size: 0.10
- Pair 2: GBPUSD, Lot Size: 0.10
- Oversold: 30, Overbought: 70
- Take Profit: $300
- Loss Threshold: $300

**Trade Sequence:**

1. **Initial Signal**: EURUSD RSI = 28, GBPUSD RSI = 27 (both oversold)
   - Action: BUY EURUSD 0.10 lots, BUY GBPUSD 0.10 lots

2. **Continues down**: Combined P&L = -$300
   - Action: BUY EURUSD 0.10 lots, BUY GBPUSD 0.10 lots (2nd set)

3. **Market recovers**: Combined P&L = +$300
   - Action: Close ALL 4 trades
   - Net Profit: $300

## Implementation Notes

### Technical Requirements
- Real-time RSI calculation for both pairs
- Combined P&L monitoring across all trades in sequence
- Trade sequencing and tracking system
- Automatic order execution capabilities
- Sufficient broker margin requirements

### Recommended Settings
- Start with small lot sizes to test the strategy
- Use demo account first before live trading
- Monitor correlation stability between pairs
- Keep detailed logs of all trade sequences
- Review and adjust RSI levels based on pair behavior

## Disclaimer

This strategy involves significant risk due to its martingale-style position averaging. It can result in substantial losses if the market moves persistently in one direction. Only trade with capital you can afford to lose, and consider implementing additional risk controls such as maximum drawdown limits.