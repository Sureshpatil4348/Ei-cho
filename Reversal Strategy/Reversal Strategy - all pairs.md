# RSI Daily Trading Strategy - Complete Documentation

## Strategy Overview

The RSI Daily Strategy is a **mean reversion trading system** that uses the Relative Strength Index (RSI) indicator on the **Daily (D1) timeframe** to identify oversold and overbought conditions. The strategy enters trades when RSI crosses key threshold levels and exits when RSI reaches predetermined exit levels.

### Key Features
- ✅ **Multi-directional trading**: Can trade both Long (Buy) and Short (Sell) positions
- ✅ **Multiple concurrent positions**: Allows multiple positions per symbol in the same direction
- ✅ **Multi-symbol support**: Can trade on 1 to 28 currency pairs simultaneously
- ✅ **Global profit/loss exit**: Closes all positions when total P/L reaches target
- ✅ **Daily timeframe**: All signals are based on daily candle closes

---

## RSI Indicator

### What is RSI?

The **Relative Strength Index (RSI)** is a momentum oscillator that measures the speed and magnitude of price movements. It oscillates between 0 and 100.

### RSI Calculation Formula

The RSI is calculated using the following steps:

**Step 1: Calculate Price Changes**
```
Gain = Current Close - Previous Close (if positive, otherwise 0)
Loss = Previous Close - Current Close (if positive, otherwise 0)
```

**Step 2: Calculate Average Gain and Average Loss**
```
Average Gain = SMA(Gains, Period)
Average Loss = SMA(Losses, Period)
```
Where:
- SMA = Simple Moving Average
- Period = RSI Period (default: 14)

**Step 3: Calculate Relative Strength (RS)**
```
RS = Average Gain / Average Loss
```

**Step 4: Calculate RSI**
```
RSI = 100 - (100 / (1 + RS))
```

### RSI Interpretation
- **RSI > 70**: Overbought condition (price may reverse downward)
- **RSI < 30**: Oversold condition (price may reverse upward)
- **RSI = 50**: Neutral (balanced buying and selling pressure)

---

## Trading Rules

### 1. LONG (Buy) Positions

#### Entry Condition
- **Signal**: RSI crosses **above** the Long Entry Level (default: 30)
- **Logic**: `Previous RSI < 30 AND Current RSI >= 30`
- **Meaning**: Market was oversold and is now recovering → Buy signal
- **Timing**: Entry occurs at the **open of the next bar** after the cross

#### Exit Condition
- **Signal**: RSI exceeds the Long Exit Level (default: 60)
- **Logic**: `Current RSI > 60`
- **Meaning**: Market has recovered sufficiently → Take profit
- **Action**: **Close ALL long positions** for that symbol

---

### 2. SHORT (Sell) Positions

#### Entry Condition
- **Signal**: RSI crosses **below** the Short Entry Level (default: 70)
- **Logic**: `Previous RSI > 70 AND Current RSI <= 70`
- **Meaning**: Market was overbought and is now declining → Sell signal
- **Timing**: Entry occurs at the **open of the next bar** after the cross

#### Exit Condition
- **Signal**: RSI drops below the Short Exit Level (default: 40)
- **Logic**: `Current RSI < 40`
- **Meaning**: Market has declined sufficiently → Take profit
- **Action**: **Close ALL short positions** for that symbol

---

## Entry and Exit Flow Diagram

```
LONG TRADE FLOW:
RSI < 30 (Oversold) → RSI crosses above 30 → BUY → Hold → RSI > 60 → EXIT

SHORT TRADE FLOW:
RSI > 70 (Overbought) → RSI crosses below 70 → SELL → Hold → RSI < 40 → EXIT
```

---

## Key Strategy Parameters

### RSI Settings
| Parameter | Default Value | Description |
|-----------|---------------|-------------|
| **RSI Period** | 14 | Number of bars used for RSI calculation |
| **Applied Price** | Close | Price type used (Close, Open, High, Low) |
| **Timeframe** | D1 (Daily) | Trading timeframe (only daily candles) |

### Entry Thresholds
| Parameter | Default Value | Description |
|-----------|---------------|-------------|
| **Long Entry Level** | 30 | RSI must cross above this to trigger BUY |
| **Short Entry Level** | 70 | RSI must cross below this to trigger SELL |

### Exit Thresholds
| Parameter | Default Value | Description |
|-----------|---------------|-------------|
| **Long Exit Level** | 60 | Close all LONG positions when RSI > 60 |
| **Short Exit Level** | 40 | Close all SHORT positions when RSI < 40 |

### Trade Direction Options
| Option | Description |
|--------|-------------|
| **Both** | Trade both Buy and Sell signals |
| **Buy Only** | Only take LONG trades |
| **Sell Only** | Only take SHORT trades |

---

## Position Management

### Multiple Positions Per Symbol
- The EA can open **multiple positions** for the same symbol in the same direction
- Each time RSI crosses the entry threshold on a new daily bar, a **new position** is opened
- Example: If RSI crosses above 30 on three different occasions, the EA will have **3 concurrent long positions**

### Position Limits
| Parameter | Default Value | Description |
|-----------|---------------|-------------|
| **Max Long Positions** | 10 | Maximum long positions per symbol |
| **Max Short Positions** | 10 | Maximum short positions per symbol |
| **Max Total Positions** | 100 | Maximum total positions across all symbols |

### Exit Behavior
- When exit conditions are met, **ALL positions** of that type for that symbol are closed
- Example: If there are 3 long positions open and RSI > 60, all 3 positions close

---

## Multi-Symbol Trading

### Symbol Mode Options

| Mode | Description | Number of Pairs |
|------|-------------|-----------------|
| **Current Only** | Trade only the chart symbol | 1 |
| **All Major & Minor** | Trade all 28 pairs | 28 |
| **Majors Only** | Trade 7 major pairs | 7 |
| **Minors Only** | Trade 21 minor/cross pairs | 21 |
| **Custom List** | Trade user-defined symbols | Variable |

### Supported Currency Pairs (28 Total)

#### Major Pairs (7)
- EURUSD, GBPUSD, USDJPY, USDCHF, AUDUSD, USDCAD, NZDUSD

#### EUR Crosses (6)
- EURGBP, EURJPY, EURCHF, EURAUD, EURCAD, EURNZD

#### GBP Crosses (5)
- GBPJPY, GBPCHF, GBPAUD, GBPCAD, GBPNZD

#### AUD Crosses (4)
- AUDJPY, AUDCHF, AUDCAD, AUDNZD

#### Other Crosses (6)
- NZDJPY, NZDCHF, NZDCAD, CADJPY, CADCHF, CHFJPY

---

## Lot Size Management

### Lot Size Mode Options

| Mode | Description |
|------|-------------|
| **Fixed Lot Size** | Use a constant lot size for all trades |
| **Risk Percentage** | Calculate lot size based on account balance (future enhancement) |

### Lot Size Validation
The EA automatically validates and adjusts lot sizes to comply with broker requirements:
```
Minimum Lot = Broker's minimum allowed lot size
Maximum Lot = Broker's maximum allowed lot size
Lot Step = Broker's lot size increment

Final Lot Size = FLOOR(Desired Lot / Lot Step) * Lot Step
```

---

## Global Profit Target Exit

### Purpose
Close **all open positions** across all symbols when total floating profit/loss reaches a specific threshold, then **restart** the strategy.

### Global Profit Target
- **Enabled**: Use Global Profit Target Exit = `true`
- **Target**: Global Profit Target = `$1000` (default)
- **Logic**: When total floating P/L >= $1000 → Close all positions
- **Action**: EA resets and starts looking for fresh signals

### Global Loss Limit
- **Enabled**: Use Global Loss Limit Exit = `false` (optional)
- **Limit**: Global Loss Limit = `$500` (default)
- **Logic**: When total floating P/L <= -$500 → Close all positions
- **Action**: EA resets and starts looking for fresh signals

### Calculation Formula
```
Total Floating P/L = SUM of (Position Profit + Position Swap) for all open positions
```

### Reset Behavior
After global exit is triggered:
1. All open positions are closed
2. Last bar times are reset to 0
3. EA begins monitoring for new entry signals on the next bar

---

## Important Trading Concepts

### 1. Bar-by-Bar Execution
- The EA only checks for signals when a **new daily bar opens**
- This prevents multiple signals within the same day
- Ensures each entry is based on a completed daily candle

### 2. Signal Detection Logic
**For Long Entry:**
```
IF (Previous Bar RSI < 30) AND (Current Bar RSI >= 30) THEN
    Open Long Position
```

**For Short Entry:**
```
IF (Previous Bar RSI > 70) AND (Current Bar RSI <= 70) THEN
    Open Short Position
```

### 3. Exit is Independent of Entry
- Exit conditions are checked on **every tick**
- Positions can be closed at any time when RSI reaches exit level
- No need to wait for a bar close to exit

---

## Order Execution Details

### Buy Order
```
Entry Price = ASK price
Order Type = Market Order (ORDER_TYPE_BUY)
Stop Loss = 0 (not used)
Take Profit = 0 (exit based on RSI only)
```

### Sell Order
```
Entry Price = BID price
Order Type = Market Order (ORDER_TYPE_SELL)
Stop Loss = 0 (not used)
Take Profit = 0 (exit based on RSI only)
```

### Magic Number
- Each trade is tagged with a **Magic Number** (default: 123456)
- This allows the EA to identify and manage only its own trades
- Multiple EAs can run on the same account with different magic numbers

---

## Risk Management Considerations

### No Stop Loss
- ⚠️ **Warning**: The strategy does NOT use hard stop losses
- Positions are managed purely by RSI exit levels
- Risk is controlled by:
  - Position limits
  - Global profit/loss targets
  - Exit thresholds

### Multiple Positions Risk
- With multiple concurrent positions, exposure can increase
- Example: 10 long positions on the same symbol = 10x the risk
- Monitor position limits carefully

### Global Exit as Safety Net
- The global profit/loss exit acts as a **portfolio-level stop**
- Prevents runaway losses across all symbols
- Locks in profits when target is reached

---

## Example Trade Scenarios

### Scenario 1: Single Long Trade
```
Day 1:  RSI = 28 (no signal yet, RSI not crossed)
Day 2:  RSI = 32 (CROSS ABOVE 30 → BUY signal)
Day 3:  Position opened at market
Day 4:  RSI = 45 (holding position)
Day 5:  RSI = 58 (holding position)
Day 6:  RSI = 62 (RSI > 60 → CLOSE all longs)
```

### Scenario 2: Multiple Long Positions
```
Week 1: RSI crosses above 30 → 1st long position opened
Week 2: RSI crosses above 30 again → 2nd long position opened
Week 3: RSI crosses above 30 again → 3rd long position opened
Week 4: RSI > 60 → ALL 3 positions closed simultaneously
```

### Scenario 3: Global Profit Exit
```
Symbol A: 2 long positions, +$300 floating
Symbol B: 3 short positions, +$450 floating
Symbol C: 1 long position, +$280 floating
Total P/L = $1030 (>= $1000 target)
→ ALL 6 positions across all symbols are closed
→ EA resets and starts fresh
```

---

## Dashboard Display

The EA displays real-time information on the chart:

- **Strategy Version**: RSI Daily Strategy v2.0
- **Symbol Mode**: Current trading mode (Current, All, Majors, etc.)
- **Active Symbols**: Number of symbols being monitored
- **Entry Thresholds**: Long < 30 | Short > 70
- **Exit Thresholds**: Long > 60 | Short < 40
- **Open Positions**: Count of long and short positions
- **Floating P/L**: Current total profit/loss across all positions
- **Trade Direction**: Buy Only, Sell Only, or Both
- **Active Symbol RSI Values**: Real-time RSI values for up to 10 symbols

---

## Configuration Checklist for Developers

When coding this strategy, ensure:

1. ✅ RSI is calculated on **Daily (D1) timeframe only**
2. ✅ Entry signals require **RSI crossing the threshold** (not just being above/below)
3. ✅ Exit signals close **ALL positions** of that direction for that symbol
4. ✅ New bar detection prevents duplicate signals within the same day
5. ✅ Global profit/loss check happens **before** processing any trades
6. ✅ Position limits are enforced before opening new trades
7. ✅ Lot sizes are validated and normalized to broker requirements
8. ✅ Magic number is applied to all trades for proper tracking
9. ✅ Multi-symbol mode loads and monitors correct symbols
10. ✅ Dashboard updates with real-time information

---

## Summary

This RSI Daily Strategy is a **mean reversion system** that:
- Uses daily RSI indicator to identify entry and exit points
- Enters on RSI threshold crosses (oversold/overbought)
- Exits when RSI reaches predetermined levels
- Supports multiple concurrent positions per symbol
- Can trade up to 28 currency pairs simultaneously
- Uses global profit/loss targets as portfolio-level risk management
- Does not use fixed stop losses or take profits (RSI-based exits only)

**Key Advantage**: Systematic approach to catching reversals from extreme RSI levels  
**Key Risk**: No hard stop losses, relying on RSI exit signals and global limits

---

*This documentation provides all necessary information for a developer to understand and code the RSI Daily Strategy EA.*
