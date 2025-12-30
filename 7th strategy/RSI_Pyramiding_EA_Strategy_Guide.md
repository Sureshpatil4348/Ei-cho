# RSI Pyramiding EA - Strategy Guide

## Overview

The **RSI Pyramiding EA** is a sophisticated automated trading system for MetaTrader 4 that builds positions gradually (pyramiding) based on RSI (Relative Strength Index) signals and trend filters. The strategy can trade in **BUY ONLY**, **SELL ONLY**, or **BOTH** directions, allowing flexible market participation.

---

## 🎯 Core Concept

### What is Pyramiding?

Pyramiding means **adding multiple positions** as the market moves in your favor. Instead of entering with one large trade, this EA:
- Opens an initial small position based on entry signals
- Adds more positions as price moves favorably
- Manages up to **20 simultaneous trades** in the same direction
- Protects profits with intelligent exit rules

### How It Works (Simple Version)

1. **Wait for Signal**: EA waits for RSI to cross oversold (for BUY) or overbought (for SELL)
2. **Enter First Trade**: Opens Trade #1 with a stop loss
3. **Add More Trades**: As price moves up (BUY) or down (SELL), EA adds more trades
4. **Exit**: Closes all trades when profit target is hit OR profit protection triggers

---

## 📊 Strategy Logic

### BUY Strategy

#### Entry Conditions
- **Trend Filter** (Optional): Price must be above 200 EMA
- **RSI Signal** (Optional): RSI crosses from below 30 (oversold) to above 30
- **Alternative Entry**: If RSI is disabled, enters on new candle when above EMA

#### Trade Sequence
1. **Trade 1**: 
   - Entry: RSI crosses above oversold OR new candle above EMA
   - Stop Loss: Set at entry price minus `First_Trade_SL` (e.g., $50)
   
2. **Trade 2**: 
   - Condition: Trade 1 profit ≥ `Min_Profit_For_Second` (e.g., $30)
   - Condition: Candle closes above Trade 1 entry price
   - Condition: RSI not overbought (< 70)
   - **Action**: Stop loss removed from Trade 1

3. **Trades 3-20** (Pyramiding): 
   - Entry: Each new candle that closes above the previous trade entry
   - No stop loss on these trades

4. **FIFO Mode** (After reaching max trades):
   - When max trades reached (e.g., 20 trades)
   - Closes oldest trade
   - Opens new trade at current level

### SELL Strategy

Same logic as BUY, but **inverted**:
- Price below EMA (trend filter)
- RSI crosses from above 70 (overbought) to below 70
- Pyramiding downward on each candle closing below previous entry

---

## ⏱️ Candle Close vs Tick Data - CRITICAL TIMING INFO

Understanding **when** the EA checks conditions is crucial for proper usage and backtesting.

### 📊 What is Tick Data?

**Tick data** = Real-time price updates. Every time the price moves (even by 0.1 pips), the EA's `OnTick()` function runs.

### 🕯️ What is Candle Close?

**Candle close** = Waits for a complete candle to finish on your chosen timeframe before taking action. More stable, fewer false signals.

---

### ✅ When EA Uses TICK DATA (Real-time Monitoring)

The EA monitors these conditions **on every tick** (instant, no waiting):

#### 1. **Trade 1 Entry Signal (RSI Mode)**
- **What**: RSI crossing oversold/overbought levels
- **Why Tick**: Catches RSI crossovers immediately as they happen
- **Impact**: Can enter mid-candle if RSI crosses during the candle
- **Example**: If RSI crosses above 30 at any moment, Trade 1 opens instantly

#### 2. **Exit Conditions (Profit Target & Protection)**
- **What**: Checking combined profit levels
- **Why Tick**: Protects profits and exits immediately when targets hit
- **Impact**: Can exit mid-candle if profit reaches target
- **Example**: If combined profit hits $500.00, all trades close instantly

#### 3. **Trade 1 Profit Check for Trade 2**
- **What**: Monitoring if Trade 1 profit ≥ `Min_Profit_For_Second`
- **Why Tick**: Real-time profit calculation
- **Impact**: Partial condition for opening Trade 2
- **Note**: Still requires candle close condition (see below)

#### 4. **Dashboard Updates**
- **What**: All displayed statistics
- **Why Tick**: Provides live, real-time information
- **Impact**: You see current P/L updating constantly

#### 5. **EMA & RSI Filter Checks**
- **What**: Market above/below EMA verification
- **Why Tick**: Prevents entries if trend changes mid-process
- **Impact**: Dynamic trend filtering

---

### 🕯️ When EA Uses CANDLE CLOSE (Wait for Completion)

The EA waits for **candle close** on the **`Pyramiding_Timeframe`** for these actions:

#### 1. **Trade 2 Entry**
- **What**: Opening the second trade
- **Candle Close Required**: Previous candle must close **above** Trade 1 entry (BUY) or **below** (SELL)
- **Timeframe Used**: `Pyramiding_Timeframe` parameter
- **Why**: Confirms the upward/downward movement is sustained
- **Example**: 
  - Trade 1 opened at $2,000
  - M5 candle closes at $2,025 (above entry)
  - Trade 1 profit is $35 (≥ $30 required)
  - Trade 2 opens on **next tick after candle close**

#### 2. **Trades 3-20 (Pyramiding Entries)**
- **What**: All pyramiding trades after Trade 2
- **Candle Close Required**: Previous candle must close **above** last entry (BUY) or **below** (SELL)
- **Timeframe Used**: `Pyramiding_Timeframe` parameter
- **Why**: Prevents over-trading and requires price confirmation
- **Example**: 
  - Last trade at $2,025
  - M5 candle closes at $2,030 (above last entry)
  - New pyramiding trade opens

#### 3. **FIFO Rotation**
- **What**: Close oldest + open newest when max trades reached
- **Candle Close Required**: Same as pyramiding logic
- **Timeframe Used**: `Pyramiding_Timeframe` parameter
- **Why**: Consistent with pyramiding timing

#### 4. **EMA-Only Mode Entry (Trade 1)**
- **What**: When RSI filter is disabled
- **Candle Close Required**: Waits for new candle to start
- **Timeframe Used**: `Pyramiding_Timeframe` parameter
- **Why**: Prevents multiple entries on same candle

---

### 🎯 Practical Examples

#### Example 1: M5 Pyramiding Timeframe (Fast Trading)

**Setup:**
- Pyramiding_Timeframe = M5
- Trade 1 opened at $2,000

**Timeline:**
- 10:00:00 - Trade 1 opens (RSI crossed on tick)
- 10:04:30 - Price at $2,028, Trade 1 profit = $28 (not enough for Trade 2 yet)
- 10:04:59 - M5 candle closes at $2,030, Trade 1 profit = $35 ✅
- **10:05:00 - Trade 2 opens** (new M5 candle starts, all conditions met)
- **10:05:01 - Price continues, but EA won't open Trade 3 until next M5 candle close**
- 10:09:59 - M5 candle closes at $2,035 (above $2,030)
- **10:10:00 - Trade 3 opens**

---

#### Example 2: H1 Pyramiding Timeframe (Slower Trading)

**Setup:**
- Pyramiding_Timeframe = H1
- Trade 1 opened at $2,000

**Timeline:**
- 09:15:22 - Trade 1 opens (RSI crossed on tick)
- 09:30:00 - Price at $2,040, Trade 1 profit = $45 ✅
- **09:59:58 - Still waiting for H1 candle close** (even though profit condition met)
- 09:59:59 - H1 candle closes at $2,045
- **10:00:00 - Trade 2 opens** (new H1 candle starts)
- 10:30:00 - Price continues, reaches $2,060
- **10:59:59 - Still waiting for H1 candle close**
- 10:59:59 - H1 candle closes at $2,055
- **11:00:00 - Trade 3 opens**

**Result**: Much slower pyramiding, more conservative, fewer trades

---

### 🔧 How to Choose Your Timeframe

| Timeframe | Speed | Trades/Day | Best For | Risk Level |
|-----------|-------|------------|----------|----------|
| **M1** | Very Fast | 50-200+ | Scalping, very active trends | ⚠️ High |
| **M5** | Fast | 20-100 | Day trading, strong trends | Medium-High |
| **M15** | Moderate | 10-40 | Intraday, clear trends | Medium |
| **M30** | Slower | 5-20 | Position building | Medium-Low |
| **H1** | Slow | 2-10 | Swing trading, daily trends | ⚠️ Low |
| **H4** | Very Slow | 1-5 | Multi-day positions | Very Low |

---

### ⚡ Key Rules Summary

| Action | Timing | Controlled By |
|--------|--------|---------------|
| **Trade 1 Entry (RSI)** | ⚡ Tick (immediate) | RSI crossing happens anytime |
| **Trade 1 Entry (EMA-Only)** | 🕯️ Candle Close | `Pyramiding_Timeframe` |
| **Trade 2 Entry** | 🕯️ Candle Close | `Pyramiding_Timeframe` |
| **Pyramiding (3-20)** | 🕯️ Candle Close | `Pyramiding_Timeframe` |
| **FIFO Rotation** | 🕯️ Candle Close | `Pyramiding_Timeframe` |
| **Profit Target Exit** | ⚡ Tick (immediate) | Real-time profit monitoring |
| **Profit Protection Exit** | ⚡ Tick (immediate) | Real-time profit monitoring |
| **Dashboard Display** | ⚡ Tick (updates live) | Real-time |

---

### 💡 Important Implications

#### For Live Trading:
- **Trade 1**: Can open mid-candle (based on RSI tick)
- **Exits**: Happen immediately when profit conditions met
- **Pyramiding**: Always waits for candle close for confirmation
- **No Lag**: Exit conditions are instant (no waiting)

#### For Backtesting:
- **Tick Data Required**: For accurate Trade 1 RSI entry timing
- **Every Tick Mode**: Use "Every tick" in Strategy Tester for best accuracy
- **Candle Close Mode**: May produce different results (less accurate for this EA)
- **Timeframe Matters**: Results will vary significantly between M5 and H1 pyramiding

#### For Strategy Planning:
- **Lower Timeframe** = More trades, faster gains/losses, more active monitoring needed
- **Higher Timeframe** = Fewer trades, smoother equity curve, more patient strategy
- **Mixed Approach**: Can use H1 EMA filter + M5 pyramiding for balanced approach

---

## 🛡️ Risk Management

### Stop Loss
- **First Trade Only**: Protected with price-based stop loss (e.g., $50)
- **Trade 2+**: Stop loss removed from Trade 1 once Trade 2 opens
- **Subsequent Trades**: No individual stop losses

### Exit Conditions

The EA has **TWO automatic exit conditions**:

#### 1. Profit Target 🎯
- **Trigger**: When combined profit from all trades ≥ `Profit_Target` (e.g., $500)
- **Action**: **CLOSE ALL TRADES** - Lock in profits!

#### 2. Profit Protection 🛡️
- **Trigger**: When combined profit falls below `Combined_Profit_Protection` (e.g., $5)
- **Only Active**: After opening 2 or more trades
- **Action**: **CLOSE ALL TRADES** - Protect gains from reversal

#### 3. Manual Close Button
- Dashboard includes **"CLOSE ALL TRADES"** button
- Instantly closes all open positions
- Resets EA to watch for fresh entry

---

## ⚙️ Input Parameters Explained

### Trend Filter Settings

| Parameter | Default | Description |
|-----------|---------|-------------|
| `Use_EMA_Filter` | true | Enable/disable 200 EMA trend filter |
| `EMA_Period` | 200 | Period for EMA calculation |
| `EMA_Timeframe` | H1 | Timeframe for EMA (H1, H4, D1, etc.) |
| `EMA_Price` | PRICE_CLOSE | Which price to use for EMA (Close, Open, High, Low) |

**💡 TIP**: The EMA filter ensures you only trade in the direction of the major trend.

---

### RSI Settings

| Parameter | Default | Description |
|-----------|---------|-------------|
| `Use_RSI_Filter` | true | Enable/disable RSI entry signals |
| `RSI_Period` | 14 | Period for RSI calculation (standard is 14) |
| `RSI_Oversold` | 30 | Oversold level for BUY signals |
| `RSI_Overbought` | 70 | Overbought level for SELL signals |
| `RSI_Timeframe` | CURRENT | Timeframe for RSI analysis |
| `RSI_Price` | PRICE_CLOSE | Which price to use for RSI |

**💡 TIP**: RSI < 30 = Oversold (potential BUY). RSI > 70 = Overbought (potential SELL).

---

### Trade Direction

| Parameter | Default | Options | Description |
|-----------|---------|---------|-------------|
| `Trade_Direction` | BUY_ONLY | BUY_ONLY / SELL_ONLY / BOTH | Choose trading direction |

**💡 TIP**: 
- **BUY_ONLY**: Only opens buy positions (uptrend trading)
- **SELL_ONLY**: Only opens sell positions (downtrend trading)
- **BOTH**: Trades both directions independently (advanced)

---

### Pyramiding Settings

| Parameter | Default | Description |
|-----------|---------|-------------|
| `Pyramiding_Timeframe` | CURRENT | Timeframe used to check for new candle closes |

**💡 TIP**: This controls how often the EA checks for pyramiding entries. M5 = more frequent, H1 = less frequent.

---

### Position Sizing

| Parameter | Default | Description |
|-----------|---------|-------------|
| `Lot_Size` | 0.01 | Lot size for EACH trade (all trades same size) |

**💡 TIP**: With 20 max trades × 0.01 lots = 0.20 lots total maximum exposure.

---

### Risk Management

| Parameter | Default | Description |
|-----------|---------|-------------|
| `First_Trade_SL` | 50.0 | Stop loss for first trade in price units (e.g., $50) |
| `Min_Profit_For_Second` | 30.0 | Minimum profit required from Trade 1 before opening Trade 2 |
| `Combined_Profit_Protection` | 5.0 | Minimum combined profit to maintain (exit if below) |

**💡 TIP**: These create a safety net - first trade has SL, and profit protection prevents giving back gains.

---

### Targets & Limits

| Parameter | Default | Description |
|-----------|---------|-------------|
| `Profit_Target` | 500.0 | Target profit to close all trades (in account currency) |
| `Max_Trades` | 20 | Maximum number of simultaneous trades allowed |

**💡 TIP**: Profit target should be realistic based on your lot size and market volatility.

---

### System Settings

| Parameter | Default | Description |
|-----------|---------|-------------|
| `Magic_Number` | 123456 | Unique identifier for this EA's trades |
| `Comment_Text` | "RSI_Pyramid" | Comment added to each trade |
| `Show_Dashboard` | true | Display live statistics on chart |
| `Profit_Color` | Lime | Color for positive profit display |
| `Loss_Color` | Red | Color for negative profit display |
| `Dashboard_X` | 20 | Dashboard horizontal position |
| `Dashboard_Y` | 30 | Dashboard vertical position |

---

## 📈 Dashboard Display

The EA shows a **live dashboard** on your chart with:

### Information Displayed
- **Mode**: BUY ONLY / SELL ONLY / BOTH
- **Status**: Number of active trades
- **BUY Trades**: Current count / Max (e.g., 5/20)
- **BUY Profit**: Live combined profit/loss
- **SELL Trades**: Current count / Max (if applicable)
- **SELL Profit**: Live combined profit/loss (if applicable)
- **Protection Level**: Distance to profit protection trigger
- **Target**: Your profit target
- **RSI Level**: Current RSI value with oversold/overbought zones
- **EMA Status**: Above/Below EMA indicator
- **Pyramid TF**: Active timeframe for pyramiding
- **CLOSE ALL TRADES Button**: Manual emergency exit

---

## 🔄 Strategy Modes

### Mode 1: RSI + EMA (Full Strategy)
- **EMA Filter**: ✅ Enabled
- **RSI Filter**: ✅ Enabled
- **Best For**: Trending markets with clear RSI signals

### Mode 2: EMA Only (Aggressive)
- **EMA Filter**: ✅ Enabled
- **RSI Filter**: ❌ Disabled
- **Behavior**: Enters immediately on new candle when above/below EMA
- **Best For**: Strong trending markets

### Mode 3: RSI Only
- **EMA Filter**: ❌ Disabled
- **RSI Filter**: ✅ Enabled
- **Best For**: Range-bound markets (not recommended for beginners)

---

## 📋 Trading Example (BUY)

### Scenario: XAUUSD (Gold) with defaults

1. **Market Condition**: Price above 200 EMA (H1)
2. **RSI Signal**: RSI drops to 28, then crosses back above 30
3. **Trade 1**: Opens at $2,000 with SL at $1,950 (0.01 lots)
4. **Price Moves Up**: Gold rises to $2,020
5. **Trade 1 Profit**: Now showing $20+ profit (≥ $30 needed)
6. **New Candle**: Candle closes at $2,025 (above entry)
7. **Trade 2**: Opens at $2,025 (0.01 lots), Trade 1 SL removed
8. **Pyramiding**: Gold continues up
   - $2,030 close → Trade 3 opens
   - $2,035 close → Trade 4 opens
   - ... and so on
9. **Exit**: Combined profit reaches $500 → **ALL TRADES CLOSED** ✅

---

## ⚠️ Important Considerations

### Risk Factors
1. **Drawdown Risk**: All trades (except first) have no stop loss
2. **Reversal Risk**: If market reverses after pyramiding, profits can turn to losses quickly
3. **Profit Protection**: This is your safety net - don't set it too low

### Best Practices
1. **Start Small**: Use 0.01 lot size until comfortable
2. **Test First**: Backtest on demo account
3. **Monitor**: Check dashboard regularly
4. **Adjust Targets**: Set realistic profit targets based on timeframe
5. **Use EMA Filter**: Helps avoid counter-trend trades

### Recommended Settings by Timeframe

#### M5 (5-Minute) - Scalping
- Profit Target: $50-$100
- Max Trades: 10
- Pyramiding TF: M5

#### H1 (1-Hour) - Intraday
- Profit Target: $200-$500
- Max Trades: 15-20
- Pyramiding TF: H1

#### H4/D1 (Swing Trading)
- Profit Target: $500-$1000
- Max Trades: 20
- Pyramiding TF: H4

---

## 🚀 Quick Start Guide

1. **Load EA** on your chart
2. **Set Trade Direction** (BUY_ONLY recommended for beginners)
3. **Configure Lot Size** (start with 0.01)
4. **Set Profit Target** (e.g., $100 for small account)
5. **Enable Dashboard** to monitor performance
6. **Let it Run** - EA will automatically manage entries and exits
7. **Monitor** via dashboard and close manually if needed

---

## 🎓 Strategy Philosophy

This EA is based on the principle of **"letting winners run"**:
- Small initial risk (Trade 1 has SL)
- Add to winning positions (pyramiding)
- Exit when target reached or profit protection triggers
- **Key**: The longer the trend, the more profit potential

**Remember**: This is a trend-following strategy. It performs best in trending markets and may struggle in choppy, ranging conditions.

---

## 📞 Support & Troubleshooting

### Common Issues

**Q: EA not opening trades?**
- Check EMA filter - price might not be above/below EMA
- Check RSI - might not have crossed yet
- Verify lot size is valid for your account

**Q: Trades closing too early?**
- Increase `Profit_Target` for longer holding
- Increase `Combined_Profit_Protection` threshold

**Q: Too many trades opening?**
- Reduce `Max_Trades` parameter
- Use higher pyramiding timeframe (H1 instead of M5)

**Q: Dashboard not showing?**
- Enable `Show_Dashboard = true`
- Adjust X/Y position if off-screen

---

## 📊 Parameter Validation

The EA automatically validates your inputs:
- RSI Period: 5-50
- RSI Oversold: 10-40
- RSI Overbought: 60-90
- Lot Size: 0.01-10.0
- Max Trades: 2-50
- EMA Period: 10-500

If any parameter is out of range, EA will alert you and prevent initialization.

---

## 🔑 Key Takeaways

✅ **Pyramiding** = Adding positions as market moves in your favor  
✅ **EMA Filter** = Only trade with the major trend  
✅ **RSI Signal** = Timing entries at oversold/overbought levels  
✅ **Profit Protection** = Prevents giving back all gains  
✅ **Profit Target** = Locks in profits automatically  
✅ **Manual Control** = Close all trades button for emergencies  

---

*This EA requires careful risk management. Always test on demo first and never risk more than you can afford to lose.*
