# Pair Trading EA Strategy Documentation

## Overview
This is a **Pair Trading Strategy** that trades two currency pairs simultaneously. The EA opens trades on both symbols at the same time and manages them as a combined position, with profit/loss calculated based on the total performance of both trades together.

---

## Core Concept

### What is Pair Trading?
- Trade **two different currency pairs** at the same time
- Both positions are managed together as one "trade set"
- Take profit and stop loss are based on the **combined profit/loss** of both positions
- The strategy uses correlation analysis and relative strength to determine trade direction

---

## Correlation Calculation - Mathematical Formula

### What is Correlation?
Correlation measures how closely two currency pairs move together. The value ranges from **-1 to +1**:
- **+1**: Perfect positive correlation (pairs move exactly together)
- **0**: No correlation (pairs move independently)
- **-1**: Perfect negative correlation (pairs move exactly opposite)

### Formula Used: Pearson Correlation Coefficient

The EA uses the **Pearson correlation coefficient** to measure the linear relationship between two symbols.

#### Mathematical Formula:

```
Correlation (r) = Σ((X - X̄) × (Y - Ȳ)) / √(Σ(X - X̄)² × Σ(Y - Ȳ)²)
```

Where:
- **X** = Close prices of Symbol 1
- **Y** = Close prices of Symbol 2
- **X̄** = Mean (average) of Symbol 1 prices
- **Ȳ** = Mean (average) of Symbol 2 prices
- **Σ** = Sum of all values
- **n** = Number of bars (Correlation Period)

#### Step-by-Step Calculation Process:

**Step 1: Collect Price Data**
- Get the last **n bars** of close prices for both symbols
- Example: If Correlation Period = 30, get last 30 close prices for each symbol

**Step 2: Calculate Mean (Average) Prices**
```
Mean of Symbol 1 (X̄) = (Price₁ + Price₂ + ... + Priceₙ) / n
Mean of Symbol 2 (Ȳ) = (Price₁ + Price₂ + ... + Priceₙ) / n
```

**Step 3: Calculate Deviations from Mean**
For each bar i:
```
Deviation of Symbol 1 (dX) = Xᵢ - X̄
Deviation of Symbol 2 (dY) = Yᵢ - Ȳ
```

**Step 4: Calculate Three Components**

A. **Numerator (Covariance)**
```
Numerator = Σ(dX × dY)
          = (dX₁ × dY₁) + (dX₂ × dY₂) + ... + (dXₙ × dYₙ)
```

B. **Standard Deviation of Symbol 1**
```
SD₁ = √(Σ(dX²))
    = √((dX₁)² + (dX₂)² + ... + (dXₙ)²)
```

C. **Standard Deviation of Symbol 2**
```
SD₂ = √(Σ(dY²))
    = √((dY₁)² + (dY₂)² + ... + (dYₙ)²)
```

**Step 5: Calculate Final Correlation**
```
Correlation (r) = Numerator / (SD₁ × SD₂)
```

#### Practical Example:

**Settings:**
- Symbol 1: EURUSD
- Symbol 2: GBPUSD
- Correlation Period: 5 bars (simplified for example)
- Timeframe: D1 (Daily)

**Sample Data:**

| Bar | EURUSD Close | GBPUSD Close |
|-----|--------------|--------------|
| 1   | 1.1000       | 1.3000       |
| 2   | 1.1050       | 1.3100       |
| 3   | 1.1100       | 1.3150       |
| 4   | 1.1080       | 1.3120       |
| 5   | 1.1120       | 1.3180       |

**Calculation:**

1. **Mean Prices:**
   - EURUSD Mean (X̄) = (1.1000 + 1.1050 + 1.1100 + 1.1080 + 1.1120) / 5 = 1.1070
   - GBPUSD Mean (Ȳ) = (1.3000 + 1.3100 + 1.3150 + 1.3120 + 1.3180) / 5 = 1.3110

2. **Deviations:**
   - Bar 1: dX = -0.0070, dY = -0.0110
   - Bar 2: dX = -0.0020, dY = -0.0010
   - Bar 3: dX = +0.0030, dY = +0.0040
   - Bar 4: dX = +0.0010, dY = +0.0010
   - Bar 5: dX = +0.0050, dY = +0.0070

3. **Components:**
   - Numerator = Σ(dX × dY) = 0.000131
   - SD₁ = √(Σ(dX²)) = 0.00545
   - SD₂ = √(Σ(dY²)) = 0.00824

4. **Final Correlation:**
   - r = 0.000131 / (0.00545 × 0.00824) = **0.92**

This indicates a **strong positive correlation** (0.92 is close to +1).

---

## Relative Strength Calculation - Decision Logic

### What is Relative Strength?
Relative strength compares the **performance** of two symbols over a specified period to determine which one is stronger (outperforming) or weaker (underperforming).

### Formula Used: Percentage Return

The EA calculates the **percentage return** for each symbol over the lookback period:

```
Percentage Return = ((Current Price - Old Price) / Old Price) × 100
```

Where:
- **Current Price** = Close price of the most recent bar (bar 0)
- **Old Price** = Close price from n bars ago (bar n)
- **n** = Relative Strength Period (default: 48 bars)

### Calculation Process:

**Step 1: Get Price Data**
```
For Symbol 1:
  - Current Price (P₁_now) = Close price at bar 0
  - Old Price (P₁_old) = Close price at bar n
  
For Symbol 2:
  - Current Price (P₂_now) = Close price at bar 0
  - Old Price (P₂_old) = Close price at bar n
```

**Step 2: Calculate Returns**
```
Return₁ = ((P₁_now - P₁_old) / P₁_old) × 100
Return₂ = ((P₂_now - P₂_old) / P₂_old) × 100
```

**Step 3: Compare Returns**
- Compare Return₁ vs Return₂
- Determine which symbol outperformed (higher return)
- Determine which symbol underperformed (lower return)

**Step 4: Select Strategy Based on Correlation Type**

---

## Strategy Selection Decision Logic

The EA uses **different logic** depending on whether you're trading **Positive Correlation** or **Negative Correlation** pairs.

### Positive Correlation Strategy Logic (Mean Reversion)

**Concept:** When pairs normally move together but temporarily diverge, we expect them to converge back (mean reversion).

**Decision Rules:**

```
IF Return₁ > Return₂:
   Symbol 1 OUTPERFORMED (stronger)
   Symbol 2 UNDERPERFORMED (weaker)
   
   Strategy → SELL Symbol 1, BUY Symbol 2
   
   Reasoning: Symbol 1 has moved too high relative to Symbol 2.
              We expect Symbol 1 to fall back and Symbol 2 to catch up.

ELSE (Return₂ > Return₁):
   Symbol 2 OUTPERFORMED (stronger)
   Symbol 1 UNDERPERFORMED (weaker)
   
   Strategy → BUY Symbol 1, SELL Symbol 2
   
   Reasoning: Symbol 2 has moved too high relative to Symbol 1.
              We expect Symbol 2 to fall back and Symbol 1 to catch up.
```

#### Practical Example - Positive Correlation:

**Data:**
- Relative Strength Period: 48 bars
- Symbol 1 (EURUSD): Old Price = 1.1000, Current Price = 1.1132
- Symbol 2 (GBPUSD): Old Price = 1.3000, Current Price = 1.3065

**Calculation:**
```
Return₁ (EURUSD) = ((1.1132 - 1.1000) / 1.1000) × 100 = +1.20%
Return₂ (GBPUSD) = ((1.3065 - 1.3000) / 1.3000) × 100 = +0.50%
```

**Comparison:**
- Return₁ (+1.20%) > Return₂ (+0.50%)
- **EURUSD outperformed** by rising +1.20%
- **GBPUSD underperformed** by only rising +0.50%

**Decision:**
```
Strategy Selected: Strategy 2 (SELL EURUSD, BUY GBPUSD)

Logic: EURUSD has moved up too much compared to GBPUSD.
       We expect:
       - EURUSD to fall back (profit from SELL)
       - GBPUSD to catch up (profit from BUY)
```

---

### Negative Correlation Strategy Logic (Trend Following)

**Concept:** When pairs normally move opposite but start moving together (correlation weakens), we ride the trend in the direction both are moving.

**Decision Rules:**

```
IF (Return₁ > 0 AND Return₂ > 0):
   Both symbols are RISING
   Correlation is weakening (normally they move opposite)
   
   Strategy → BUY Symbol 1, BUY Symbol 2
   
   Reasoning: Both are in uptrend. Ride the bullish momentum.

ELSE IF (Return₁ < 0 AND Return₂ < 0):
   Both symbols are FALLING
   Correlation is weakening (normally they move opposite)
   
   Strategy → SELL Symbol 1, SELL Symbol 2
   
   Reasoning: Both are in downtrend. Ride the bearish momentum.

ELSE:
   Mixed signals (one up, one down - still negatively correlated)
   Follow the STRONGER mover's direction:
   
   IF |Return₁| > |Return₂|:
      Symbol 1 has stronger move
      
      IF Return₁ > 0:
         Strategy → BUY both (follow Symbol 1's uptrend)
      ELSE:
         Strategy → SELL both (follow Symbol 1's downtrend)
   
   ELSE:
      Symbol 2 has stronger move
      
      IF Return₂ > 0:
         Strategy → BUY both (follow Symbol 2's uptrend)
      ELSE:
         Strategy → SELL both (follow Symbol 2's downtrend)
```

#### Practical Example - Negative Correlation (Both Rising):

**Data:**
- Symbol 1 (USDCHF): Old Price = 0.9000, Current Price = 0.9135
- Symbol 2 (EURUSD): Old Price = 1.1000, Current Price = 1.1088

**Calculation:**
```
Return₁ (USDCHF) = ((0.9135 - 0.9000) / 0.9000) × 100 = +1.50%
Return₂ (EURUSD) = ((1.1088 - 1.1000) / 1.1000) × 100 = +0.80%
```

**Analysis:**
- Both returns are **positive** (both rising)
- Normally USDCHF and EURUSD move opposite (negative correlation)
- But now both are rising → Correlation is weakening

**Decision:**
```
Strategy Selected: Strategy 3 (BUY both)

Logic: Both USDCHF and EURUSD are in uptrend.
       The negative correlation is breaking down.
       We ride the uptrend momentum by buying both.
```

#### Practical Example - Negative Correlation (Mixed Signals):

**Data:**
- Symbol 1 (USDCHF): Old Price = 0.9000, Current Price = 0.9180
- Symbol 2 (EURUSD): Old Price = 1.1000, Current Price = 1.0956

**Calculation:**
```
Return₁ (USDCHF) = ((0.9180 - 0.9000) / 0.9000) × 100 = +2.00%
Return₂ (EURUSD) = ((1.0956 - 1.1000) / 1.1000) × 100 = -0.40%
```

**Analysis:**
- Return₁ = +2.00% (USDCHF rising)
- Return₂ = -0.40% (EURUSD falling)
- Mixed signals (pairs still negatively correlated)
- |Return₁| = 2.00% > |Return₂| = 0.40% (USDCHF is stronger mover)

**Decision:**
```
Strategy Selected: Strategy 3 (BUY both)

Logic: Symbol 1 (USDCHF) has the stronger move and is rising.
       We follow USDCHF's uptrend direction.
       Buy both symbols to ride the momentum.
```

---

## Complete Decision Flow Chart

```
CORRELATION TYPE?
       |
       ├─→ POSITIVE CORRELATION
       │       |
       │       Calculate Returns (Return₁, Return₂)
       │       |
       │       Compare Returns
       │       |
       │       ├─→ Return₁ > Return₂? 
       │       │      YES → Symbol 1 OUTPERFORMED
       │       │            Strategy: SELL Symbol 1, BUY Symbol 2
       │       │
       │       └─→ Return₂ > Return₁?
       │              YES → Symbol 2 OUTPERFORMED
       │                    Strategy: BUY Symbol 1, SELL Symbol 2
       │
       └─→ NEGATIVE CORRELATION
               |
               Calculate Returns (Return₁, Return₂)
               |
               Check Direction
               |
               ├─→ Both Positive (Return₁ > 0 AND Return₂ > 0)?
               │      YES → Strategy: BUY both
               │
               ├─→ Both Negative (Return₁ < 0 AND Return₂ < 0)?
               │      YES → Strategy: SELL both
               │
               └─→ Mixed Signals (one up, one down)?
                      YES → Compare |Return₁| vs |Return₂|
                            |
                            ├─→ |Return₁| > |Return₂|?
                            │      YES → Follow Symbol 1 direction
                            │            Return₁ > 0 → BUY both
                            │            Return₁ < 0 → SELL both
                            │
                            └─→ |Return₂| > |Return₁|?
                                   YES → Follow Symbol 2 direction
                                         Return₂ > 0 → BUY both
                                         Return₂ < 0 → SELL both
```

---

## Trading Pairs Configuration

### Symbol Selection
- **Symbol 1**: First currency pair (e.g., EURUSD)
- **Symbol 2**: Second currency pair (e.g., GBPUSD)
- **Lot Size 1**: Position size for Symbol 1
- **Lot Size 2**: Position size for Symbol 2

### Important Notes
- You can use any two symbols
- Lot sizes can be different for each symbol
- Both trades are opened simultaneously

---

## Trading Strategies (4 Options)

The EA supports **4 different trading strategies** that determine the direction of trades:

### Strategy 1: Buy Symbol 1, Sell Symbol 2
- **Symbol 1**: Open BUY position
- **Symbol 2**: Open SELL position
- **Use Case**: When Symbol 1 is expected to rise and Symbol 2 is expected to fall

### Strategy 2: Sell Symbol 1, Buy Symbol 2
- **Symbol 1**: Open SELL position
- **Symbol 2**: Open BUY position
- **Use Case**: When Symbol 1 is expected to fall and Symbol 2 is expected to rise

### Strategy 3: Buy Symbol 1, Buy Symbol 2
- **Symbol 1**: Open BUY position
- **Symbol 2**: Open BUY position
- **Use Case**: When both symbols are expected to rise together

### Strategy 4: Sell Symbol 1, Sell Symbol 2
- **Symbol 1**: Open SELL position
- **Symbol 2**: Open SELL position
- **Use Case**: When both symbols are expected to fall together

---

## Strategy Selection Methods

### Method 1: Manual Strategy Selection (Default)
- **How it works**: You manually select which strategy to use (1, 2, 3, or 4)
- **Setting**: Set `Use Dynamic Strategy = false`
- **Best for**: When you know exactly which direction you want to trade

### Method 2: Dynamic Strategy Selection (Recommended)
- **How it works**: EA automatically selects the best strategy based on relative strength analysis
- **Setting**: Set `Use Dynamic Strategy = true`
- **Best for**: Automated trading without manual intervention

#### How Dynamic Selection Works
The EA calculates **percentage returns** for both symbols over a lookback period (default: 48 bars) and compares them:

**For Positive Correlation:**
- If Symbol 1 outperformed Symbol 2 → Select Strategy 2 (SELL Symbol1, BUY Symbol2)
  - Logic: Buy the underperformer, sell the outperformer (mean reversion)
- If Symbol 2 outperformed Symbol 1 → Select Strategy 1 (BUY Symbol1, SELL Symbol2)
  - Logic: Buy the underperformer, sell the outperformer (mean reversion)

**For Negative Correlation:**
- If both symbols are rising → Select Strategy 3 (BUY both)
  - Logic: Ride the uptrend when correlation weakens
- If both symbols are falling → Select Strategy 4 (SELL both)
  - Logic: Ride the downtrend when correlation weakens
- If mixed signals → Follow the stronger mover's direction

---

## Entry System

### Option 1: Auto Start Entry (Backtesting)
- **Setting**: `Auto Start = true` AND `Enable Correlation Entry = false`
- **How it works**: Places trades immediately when EA starts
- **Use case**: For backtesting purposes

### Option 2: Correlation-Based Entry (Live Trading)
- **Setting**: `Enable Correlation Entry = true`
- **How it works**: Waits for correlation to break threshold before entering

#### Correlation Entry Logic

**Positive Correlation Mode:**
- EA measures correlation between Symbol 1 and Symbol 2
- **Entry Trigger**: When correlation drops BELOW the threshold (e.g., 0.25)
- **Logic**: Pairs normally move together, but when they diverge (low correlation), there's a trading opportunity
- **Example**: If threshold is 0.25, EA enters when correlation falls below 0.25

**Negative Correlation Mode:**
- EA measures correlation between Symbol 1 and Symbol 2
- **Entry Trigger**: When correlation rises ABOVE the threshold (e.g., -0.25 to 0.25)
- **Logic**: Pairs normally move opposite, but when they start moving together (correlation increases), there's a trading opportunity
- **Example**: If threshold is 0.25, EA enters when correlation rises above 0.25

#### Correlation Parameters
- **Correlation Type**: Positive or Negative
- **Entry Threshold**: The correlation value that triggers entry (e.g., 0.25)
- **Timeframe**: Timeframe to calculate correlation (e.g., D1 = Daily)
- **Correlation Period**: Number of bars to use for correlation calculation (e.g., 30 bars)

---

## Risk Management

### Take Profit System
- **Take Profit Amount**: Dollar amount (e.g., $100)
- **How it works**: When the **combined profit** of both positions reaches this amount, ALL positions are closed
- **Example**: If TP = $100, and Symbol1 profit = $60 and Symbol2 profit = $40, total = $100 → Close all trades

### Martingale / Averaging System
- **Loss Amount to Repeat**: Base loss amount for averaging (e.g., $50)
- **How it works**: If trades go into loss, EA places additional trade sets at specific loss levels
- **Formula**: Additional trades are placed at loss levels = `(2^n - 1) × Loss Amount`

#### Example (Loss Amount = $50):
1. **Initial Trade Set**: Placed at start
2. **2nd Trade Set**: Placed when loss reaches $50 [= (2^1 - 1) × $50]
3. **3rd Trade Set**: Placed when loss reaches $150 [= (2^2 - 1) × $50]
4. **4th Trade Set**: Placed when loss reaches $350 [= (2^3 - 1) × $50]
5. **5th Trade Set**: Placed when loss reaches $750 [= (2^4 - 1) × $50]

#### Important Notes
- Each new trade set uses the SAME lot sizes as the initial set
- All positions are managed together
- When TP is hit, ALL positions close simultaneously
- This creates an exponential averaging system

---

## Trade Management Flow

### Step 1: Entry
- **Correlation Entry Enabled**: Wait for correlation signal
- **Auto Start Enabled**: Enter immediately
- **Dynamic Strategy**: Calculate which strategy to use based on relative strength
- **Manual Strategy**: Use the pre-selected strategy

### Step 2: Position Monitoring
- Calculate combined P/L every tick
- Display current P/L on dashboard
- Check for Take Profit condition
- Check for Loss Amount to Repeat condition

### Step 3: Exit Conditions

#### Exit Condition 1: Take Profit Hit
- Combined P/L reaches Take Profit amount
- Close ALL positions
- Reset counters
- If correlation entry is enabled, wait for next signal

#### Exit Condition 2: Manual Close
- User clicks "Close All Trades" button
- Close ALL positions immediately
- Reset counters

### Step 4: Additional Trade Sets (Averaging)
- If combined P/L is negative and reaches loss threshold
- Place additional trade set using the SAME strategy
- Continue monitoring combined P/L
- When TP is eventually hit, all positions close together

---

## Input Parameters Summary

### Basic Settings
- **Magic Number**: Unique identifier for EA's trades (default: 123456)

### Trading Pair Configuration
- **Symbol 1**: First currency pair
- **Symbol 2**: Second currency pair
- **Lot Size 1**: Position size for Symbol 1
- **Lot Size 2**: Position size for Symbol 2

### Entry System (Correlation-Based)
- **Enable Correlation Entry**: true/false
- **Correlation Type**: Positive or Negative
- **Entry Threshold**: Correlation value to trigger entry (e.g., 0.25)
- **Timeframe**: Timeframe for correlation calculation (e.g., D1)
- **Correlation Period**: Number of bars for correlation (e.g., 30)

### Strategy Selection
- **Use Dynamic Strategy**: true (automatic) or false (manual)
- **Relative Strength Period**: Bars for relative strength calculation (default: 48)
- **Manual Strategy**: Which strategy to use if dynamic is disabled (1-4)

### Risk Management
- **Take Profit Amount**: Dollar profit target (e.g., $100)
- **Loss Amount to Repeat**: Base loss for martingale system (e.g., $50)

### Backtesting Options
- **Auto Start**: Place trades immediately on EA start (ignored if correlation entry enabled)

---

## Important Strategy Rules

### Rule 1: Combined Position Management
- Both trades are ALWAYS managed together
- P/L is calculated as the SUM of both positions
- Exit conditions apply to COMBINED P/L, not individual trades

### Rule 2: Same Strategy for All Trade Sets
- Initial trade set uses calculated/selected strategy
- ALL additional trade sets use the SAME strategy
- Strategy does NOT change mid-sequence

### Rule 3: Correlation Entry Only Triggers Once
- When correlation entry is enabled, EA enters when signal is triggered
- After entry, EA manages positions until TP or manual close
- After exit, EA waits for the NEXT correlation signal

### Rule 4: Exponential Loss Levels
- Loss levels increase exponentially: $50, $150, $350, $750, etc.
- This allows for recovery with larger combined positions
- Risk increases with each additional trade set

---

## Strategy Logic Flow

```
START
  ↓
Is Correlation Entry Enabled?
  ↓ YES                           ↓ NO
Wait for Signal                Auto Start?
  ↓                                ↓ YES
Signal Triggered?              Enter Immediately
  ↓ YES
Use Dynamic Strategy?
  ↓ YES                           ↓ NO
Calculate Best Strategy        Use Manual Strategy
  ↓                                ↓
Place Initial Trade Set ←──────────┘
  ↓
Monitor Combined P/L
  ↓
Check Conditions Every Tick
  ├─→ P/L >= Take Profit? → Close All → WAIT FOR NEXT SIGNAL
  ├─→ Loss >= Next Level? → Place Additional Trade Set → Continue Monitoring
  └─→ Manual Close? → Close All → WAIT FOR NEXT SIGNAL
```

---

## Key Strategy Features

### 1. Correlation Analysis
- Measures how two symbols move together
- Uses Pearson correlation coefficient
- Detects divergence (positive correlation) or convergence (negative correlation)

### 2. Relative Strength Analysis
- Calculates percentage returns for both symbols
- Determines which symbol is stronger/weaker
- Automatically selects optimal strategy based on strength

### 3. Portfolio Approach
- Treats both positions as one portfolio
- Risk and profit are measured on combined basis
- Reduces dependency on single pair performance

### 4. Adaptive Entry
- Entry timing based on correlation breakdown
- Avoids entering when pairs are moving normally
- Enters when abnormal relationship is detected

### 5. Progressive Averaging
- Adds positions at increasing loss intervals
- Exponential growth in loss thresholds
- Designed to recover from drawdowns

---

## Example Trading Scenario

### Scenario: EURUSD and GBPUSD with Positive Correlation

**Settings:**
- Symbol 1: EURUSD (Lot 0.01)
- Symbol 2: GBPUSD (Lot 0.01)
- Correlation Type: Positive
- Threshold: 0.25
- Dynamic Strategy: Enabled
- Relative Strength Period: 48 bars
- Take Profit: $100
- Loss to Repeat: $50

**Trading Flow:**

1. **Waiting Phase**
   - EA monitors correlation on D1 timeframe
   - Current correlation: 0.85 (high positive correlation)
   - Status: Waiting for signal

2. **Entry Signal**
   - Correlation drops to 0.20 (below threshold 0.25)
   - EA calculates relative strength over 48 bars
   - EURUSD return: +1.2%
   - GBPUSD return: -0.5%
   - EURUSD outperformed → Strategy: SELL EURUSD, BUY GBPUSD
   - EA enters with initial trade set

3. **Scenario A: Profit Target Hit**
   - EURUSD falls, GBPUSD rises (mean reversion)
   - Combined P/L reaches $100
   - EA closes all positions
   - Waits for next correlation signal

4. **Scenario B: Averaging Sequence**
   - EURUSD continues rising, GBPUSD continues falling
   - Loss reaches -$50 → Place 2nd trade set
   - Loss reaches -$150 → Place 3rd trade set
   - Market reverses, combined P/L reaches $100
   - EA closes all 3 trade sets
   - Waits for next correlation signal

---

## Dashboard Features

The EA includes a visual dashboard with:

- **Symbol Settings**: Editable symbols and lot sizes
- **Strategy Selection**: Radio buttons for manual strategy selection
- **Risk Settings**: Take Profit and Loss to Repeat inputs
- **Place Trades Button**: Manually trigger trade entry
- **Real-time P/L Display**: Shows current combined profit/loss
- **Close All Button**: Emergency exit for all positions

---

## Developer Implementation Notes

### Core Components Required

1. **Correlation Calculator**
   - Pearson correlation coefficient
   - Configurable timeframe and period
   - Returns value between -1 and +1

2. **Relative Strength Calculator**
   - Percentage return calculation
   - Compare two symbols over lookback period
   - Determine stronger/weaker symbol

3. **Strategy Selector**
   - Dynamic mode: Calculate based on correlation type and relative strength
   - Manual mode: Use user-selected strategy
   - Return strategy number (1-4)

4. **Trade Set Manager**
   - Place paired trades simultaneously
   - Track all positions with magic number
   - Calculate combined P/L
   - Close all positions together

5. **Averaging System**
   - Track current trade set count
   - Calculate next loss threshold using exponential formula
   - Trigger new trade sets at correct levels

6. **Entry Signal Monitor**
   - Check correlation on new bar
   - Compare against threshold based on correlation type
   - Trigger entry when conditions met

### Critical Logic Points

- **ONE entry per correlation signal**: Don't re-enter while trades are active
- **Combined P/L calculation**: Always sum ALL positions with matching magic number
- **Exponential loss formula**: Correctly implement (2^n - 1) × Loss Amount
- **Strategy consistency**: Use same strategy for all trade sets in a sequence
- **Correlation direction**: Positive = enter when low, Negative = enter when high

---

## Risk Warnings

1. **Exponential Risk Growth**: Loss levels increase rapidly (50, 150, 350, 750...)
2. **Multiple Positions**: Can accumulate many positions during drawdowns
3. **Correlation Changes**: Correlation can shift unexpectedly
4. **Double Exposure**: Trading two symbols simultaneously doubles exposure
5. **Spread Costs**: Each averaging adds spreads from 2 symbols

---

## Summary

This is a **correlation-based pair trading strategy** that:
- Trades two currency pairs simultaneously
- Uses correlation analysis to time entries
- Employs relative strength to select trade direction
- Manages positions as a combined portfolio
- Implements progressive averaging for recovery
- Exits all positions when combined profit target is reached

The strategy is designed to profit from temporary divergences (positive correlation) or convergences (negative correlation) between two related currency pairs.
