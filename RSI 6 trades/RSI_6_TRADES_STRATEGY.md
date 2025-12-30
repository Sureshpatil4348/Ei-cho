# RSI 6 Trades - Laddered Martingale Strategy

## 📋 Table of Contents
- [Overview](#overview)
- [Strategy Logic](#strategy-logic)
- [Entry Conditions](#entry-conditions)
- [Exit Conditions](#exit-conditions)
- [Position Sizing](#position-sizing)
- [Parameters](#parameters)
- [Risk Management](#risk-management)
- [Visual Flow](#visual-flow)
- [Examples](#examples)
- [Important Notes](#important-notes)

---

## Overview

**RSI 6 Trades** is an automated martingale-based trading strategy that uses RSI (Relative Strength Index) signals from multiple timeframes to enter positions. The strategy scales into trades using ATR-based grid spacing and exits when price retraces to the first entry or reaches an ATR-derived profit target.

### Key Features
- ✅ Dual timeframe RSI confirmation (Lower + Higher timeframe)
- ✅ Independent buy and sell baskets (can run simultaneously)
- ✅ ATR-based dynamic grid spacing for entries
- ✅ ATR-based take profit targets
- ✅ Martingale position sizing (doubling on each scale-in)
- ✅ Zone permission system (prevents multiple sequences in same RSI zone)
- ✅ Retrace-to-first exit mechanism

---

## Strategy Logic

### Core Concept
The strategy identifies extreme RSI conditions (overbought/oversold) on two timeframes and enters counter-trend positions. If price continues against the position, it adds to the position using ATR-based spacing and martingale sizing.

### Trading Sides
1. **BUY Side**: Triggered when RSI is oversold (≤ 30)
2. **SELL Side**: Triggered when RSI is overbought (≥ 70)

Both sides can operate independently and simultaneously if `InpAllowBothSides = true`.

---

## Entry Conditions

### First Entry Requirements

#### SELL Side First Entry
All conditions must be met:
1. ✅ No existing SELL positions for this strategy
2. ✅ RSI on **Lower Timeframe** (LTF) ≥ Overbought threshold (default: 70)
3. ✅ RSI on **Higher Timeframe** (HTF) ≥ Overbought threshold (default: 70)
4. ✅ Zone not already used (prevents re-entry in same RSI stay)
5. ✅ If `InpAllowBothSides = false`, no BUY positions exist
6. ✅ If `InpFirstEntryWaitClose = true`, uses closed candle RSI; else uses current candle RSI

**Action**: Open SELL position with `InpInitialLot` size

**State Changes**:
- Freeze first entry price
- Capture ATR values (grid spacing & take profit distance)
- Mark zone as "used"
- Set next index to 1

#### BUY Side First Entry
All conditions must be met:
1. ✅ No existing BUY positions for this strategy
2. ✅ RSI on **Lower Timeframe** (LTF) ≤ Oversold threshold (default: 30)
3. ✅ RSI on **Higher Timeframe** (HTF) ≤ Oversold threshold (default: 30)
4. ✅ Zone not already used (prevents re-entry in same RSI stay)
5. ✅ If `InpAllowBothSides = false`, no SELL positions exist
6. ✅ If `InpFirstEntryWaitClose = true`, uses closed candle RSI; else uses current candle RSI

**Action**: Open BUY position with `InpInitialLot` size

**State Changes**:
- Freeze first entry price
- Capture ATR values (grid spacing & take profit distance)
- Mark zone as "used"
- Set next index to 1

---

### Scaling Entries (Adding to Position)

#### SELL Scaling
Conditions:
1. ✅ At least 1 SELL position exists
2. ✅ Total SELL positions < `InpMaxTrades` (default: 8)
3. ✅ Current ASK price ≥ `FirstPrice + (NextIndex × ATR_Grid)`
4. ✅ Current ASK price > Latest trade entry price

**Action**: 
- Open additional SELL with lot size = `LastLot × InpMartingaleMultiplier`
- Increment NextIndex
- Update target price for next scale-in

**Note**: This happens tick-by-tick (real-time), not waiting for candle close

#### BUY Scaling
Conditions:
1. ✅ At least 1 BUY position exists
2. ✅ Total BUY positions < `InpMaxTrades` (default: 8)
3. ✅ Current BID price ≤ `FirstPrice - (NextIndex × ATR_Grid)`
4. ✅ Current BID price < Latest trade entry price

**Action**: 
- Open additional BUY with lot size = `LastLot × InpMartingaleMultiplier`
- Increment NextIndex
- Update target price for next scale-in

**Note**: This happens tick-by-tick (real-time), not waiting for candle close

---

## Exit Conditions

The strategy has **two exit mechanisms** for each side:

### 1. Retrace to First Entry (Priority Exit)
**Triggers when positions > 1**

#### SELL Basket Exit
- **Condition**: Current ASK price ≤ First SELL entry price
- **Action**: Close all SELL positions immediately
- **Logic**: Price has retraced back to breakeven point

#### BUY Basket Exit
- **Condition**: Current BID price ≥ First BUY entry price
- **Action**: Close all BUY positions immediately
- **Logic**: Price has retraced back to breakeven point

---

### 2. ATR-Based Take Profit
**Triggers when positions < InpMaxTrades**

#### SELL Basket TP
- **Calculation**: 
  - VWAP = Volume-Weighted Average Price of all SELL positions
  - Target distance = `ATR_TP_Sell` (frozen at first entry)
- **Condition**: `(VWAP - Current_ASK) ≥ ATR_TP_Sell`
- **Action**: Close all SELL positions

#### BUY Basket TP
- **Calculation**: 
  - VWAP = Volume-Weighted Average Price of all BUY positions
  - Target distance = `ATR_TP_Buy` (frozen at first entry)
- **Condition**: `(Current_BID - VWAP) ≥ ATR_TP_Buy`
- **Action**: Close all BUY positions

---

### Zone Re-Arming (Permission Reset)

After closing all positions on a side, the zone permission resets when:
- **SELL**: RSI on LTF exits overbought zone (drops below 70)
- **BUY**: RSI on LTF exits oversold zone (rises above 30)

This allows a new sequence to start on the next RSI extreme.

---

## Position Sizing

### Martingale Progression
Starting with `InpInitialLot = 0.01` and `InpMartingaleMultiplier = 2.0`:

| Entry # | Lot Size | Cumulative Lots |
|---------|----------|-----------------|
| 1       | 0.01     | 0.01            |
| 2       | 0.02     | 0.03            |
| 3       | 0.04     | 0.07            |
| 4       | 0.08     | 0.15            |
| 5       | 0.16     | 0.31            |
| 6       | 0.32     | 0.63            |
| 7       | 0.64     | 1.27            |
| 8       | 1.28     | 2.55            |

**⚠️ WARNING**: Exposure grows exponentially. With max 8 trades, total exposure is **255x** initial lot size!

---

## Parameters

### RSI Configuration
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `InpRSITimeframe` | Timeframe | M5 | Lower timeframe for RSI calculation |
| `InpRSIHigherTimeframe` | Timeframe | H1 | Higher timeframe for RSI filter |
| `InpRSIPeriod` | Integer | 14 | RSI period (standard Wilder's) |
| `InpRSIOverbought` | Double | 70.0 | Threshold for SELL entries |
| `InpRSIOversold` | Double | 30.0 | Threshold for BUY entries |
| `InpFirstEntryWaitClose` | Boolean | true | Wait for candle close for first entry |

### Position Management
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `InpInitialLot` | Double | 0.01 | Starting position size |
| `InpMartingaleMultiplier` | Double | 2.0 | Multiplier for each scale-in |
| `InpMaxTrades` | Integer | 8 | Maximum positions per side |
| `InpAllowBothSides` | Boolean | true | Allow simultaneous BUY & SELL baskets |
| `InpMagicNumber` | Integer | 123456 | Unique identifier for orders |
| `InpOrderComment` | String | "RSI_Ladder_Martingale" | Order comment tag |

### ATR Configuration
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `InpATRGridTimeframe` | Timeframe | H4 | Timeframe for grid spacing ATR |
| `InpATRGridPeriod` | Integer | 14 | ATR period for grid spacing |
| `InpATRTPTimeframe` | Timeframe | H1 | Timeframe for take profit ATR |
| `InpATRTPPeriod` | Integer | 14 | ATR period for take profit |

---

## Risk Management

### Maximum Risk Per Sequence
With default settings (8 trades, 2x multiplier):
- **Total lots**: 2.55x initial lot
- **If initial = 0.01**: Maximum 0.0255 lots total
- **If initial = 0.10**: Maximum 2.55 lots total

### Grid Spacing Risk
- **ATR-based spacing**: Adapts to volatility
- **H4 ATR default**: Typically 50-150 pips on major pairs
- **Maximum drawdown**: `ATR × (1+2+3+4+5+6+7) = 28 × ATR`

### Capital Requirements
**Minimum recommended capital per pair**:
- Conservative: $10,000 (0.01 initial lot)
- Moderate: $5,000 (0.01 initial lot, careful pair selection)
- Aggressive: $2,000 (0.01 initial lot, HIGH RISK)

**⚠️ This is a martingale strategy - account can blow up rapidly if market trends strongly**

---

## Visual Flow

### Strategy Execution Flow
```
┌─────────────────────────────────────────────────────┐
│              ON EACH TICK                           │
└─────────────────────────────────────────────────────┘
                        ↓
┌─────────────────────────────────────────────────────┐
│  1. Check Exit Conditions (Priority)                │
│     • Retrace to first entry?                       │
│     • ATR take profit reached?                      │
│     → Close all positions if conditions met         │
└─────────────────────────────────────────────────────┘
                        ↓
┌─────────────────────────────────────────────────────┐
│  2. Calculate RSI (LTF & HTF)                       │
│     • Current candle RSI                            │
│     • Closed candle RSI                             │
└─────────────────────────────────────────────────────┘
                        ↓
┌─────────────────────────────────────────────────────┐
│  3. Update Zone Permissions                         │
│     • If RSI exits overbought → reset sell zone     │
│     • If RSI exits oversold → reset buy zone        │
└─────────────────────────────────────────────────────┘
                        ↓
┌─────────────────────────────────────────────────────┐
│  4. Check First Entry Conditions                    │
│     • No positions on side?                         │
│     • Both LTF & HTF RSI extreme?                   │
│     • Zone not used?                                │
│     → Open first position, freeze ATR values        │
└─────────────────────────────────────────────────────┘
                        ↓
┌─────────────────────────────────────────────────────┐
│  5. Check Scaling Conditions                        │
│     • Active positions < max?                       │
│     • Price reached next grid level?                │
│     → Add position with martingale sizing           │
└─────────────────────────────────────────────────────┘
                        ↓
┌─────────────────────────────────────────────────────┐
│  6. Draw TP Lines on Chart                          │
│     • Show visual take profit targets               │
└─────────────────────────────────────────────────────┘
```

### Entry Sequence Example (SELL Side)

```
Price moves UP →

RSI >= 70 on M5 & H1
        ↓
    Entry #1 @ 1.1000 (0.01 lots)
    [Freeze ATR_Grid = 50 pips]
    [Freeze ATR_TP = 100 pips]
        ↓
Price +50 pips (1.1050)
        ↓
    Entry #2 @ 1.1050 (0.02 lots)
        ↓
Price +50 pips (1.1100)
        ↓
    Entry #3 @ 1.1100 (0.04 lots)
        ↓
Price +50 pips (1.1150)
        ↓
    Entry #4 @ 1.1150 (0.08 lots)

Total: 0.15 lots
VWAP: ~1.1093
TP Target: 1.1093 - 0.0100 = 1.0993
Retrace Exit: 1.1000 (first entry price)
```

---

## Examples

### Example 1: Successful BUY Sequence

**Market Context**: EURUSD drops rapidly, RSI hits oversold

1. **T1**: RSI M5 = 28, RSI H1 = 25 → **Entry #1** @ 1.0900 (0.01 lots)
2. **T2**: Price drops to 1.0850 → **Entry #2** @ 1.0850 (0.02 lots)
3. **T3**: Price drops to 1.0800 → **Entry #3** @ 1.0800 (0.04 lots)
4. **T4**: Price reverses, moves to 1.0900
   - **Retrace Exit Triggered**: Price ≥ First Entry (1.0900)
   - **Close all BUY positions**
   - **Result**: Small profit/breakeven (saved by exit mechanism)

---

### Example 2: Full Grid to Max Trades

**Market Context**: Strong downtrend, SELL side activated

1. **Entry #1** @ 1.2000 (0.01 lots) - RSI M5 = 72, H1 = 74
2. **Entry #2** @ 1.2050 (0.02 lots)
3. **Entry #3** @ 1.2100 (0.04 lots)
4. **Entry #4** @ 1.2150 (0.08 lots)
5. **Entry #5** @ 1.2200 (0.16 lots)
6. **Entry #6** @ 1.2250 (0.32 lots)
7. **Entry #7** @ 1.2300 (0.64 lots)
8. **Entry #8** @ 1.2350 (1.28 lots) - **MAX REACHED**

**Total Exposure**: 2.55 lots  
**VWAP**: ~1.2259  
**ATR TP Target**: 1.2259 - 100 pips = 1.2159

Price reverses to 1.2150:
- **ATR TP Triggered**: (1.2259 - 1.2150) = 109 pips ≥ 100 pips
- **All positions closed at profit**

---

### Example 3: Zone Permission System

**Scenario**: Preventing multiple sequences in same zone

```
Time    RSI M5    RSI H1    Action
───────────────────────────────────────
09:00   72        75        Open SELL #1
09:05   74        76        (zone_used = true)
09:10   69        74        RSI exits overbought
09:15   73        76        (zone_used = false) ← Reset
09:20   76        78        Open SELL #1 ← New sequence allowed
```

Without this system, it would re-enter at 09:15, potentially creating overlapping sequences.

---

## Important Notes

### ⚠️ Critical Warnings

1. **Martingale Risk**: This strategy can suffer catastrophic losses in strong trends
2. **Capital Requirements**: Ensure sufficient capital for full 8-trade sequence
3. **Leverage Considerations**: High leverage amplifies both gains and losses
4. **Broker Compatibility**: 
   - Requires low spread
   - Fast execution
   - No restrictions on grid/hedging
5. **Market Conditions**: 
   - Best in ranging/mean-reverting markets
   - Dangerous in strong trending markets
   - News events can trigger rapid drawdowns

### ✅ Best Practices

1. **Test thoroughly** on demo account (minimum 3 months)
2. **Use appropriate lot sizes** relative to account balance
3. **Monitor correlation** if running multiple pairs
4. **Set hard stop loss** at account level (e.g., 30% drawdown)
5. **Consider time-based filters** (avoid high-impact news)
6. **Regular monitoring** required - not fully hands-off
7. **Keep records** of all sequences for post-analysis

### 🔧 Optimization Tips

1. **Adjust RSI thresholds**: Test 65/35 or 75/25 for your pair
2. **Modify timeframes**: Try M15/H4 for longer-term approach
3. **Reduce max trades**: 6 trades max = less risk (max 0.63 lots)
4. **Adjust martingale**: Use 1.5x instead of 2.0x for slower growth
5. **Symbol selection**: Best on major pairs with tight spreads (EUR/USD, GBP/USD)

---

## Technical Implementation

### MT4 Functions Used
- `iRSI()`: RSI indicator calculation
- `iATR()`: ATR indicator calculation
- `OrderSend()`: Place market orders
- `OrderClose()`: Close positions
- `OrderSelect()`: Iterate through orders
- `MarketInfo()`: Get symbol specifications

### State Variables Tracked
- First entry price per side
- Frozen ATR values (grid spacing & TP)
- Next grid index
- Zone usage flags
- RSI values (current & closed)

### Chart Objects
- Blue horizontal line: BUY basket take profit target
- Red horizontal line: SELL basket take profit target

---

## Comparison: MT4 vs Python Implementation

| Feature | MT4 (MQ4) | Python (MT5) | Status |
|---------|-----------|--------------|--------|
| Core Logic | ✅ Complete | ✅ Complete | ✅ Identical |
| RSI Calculation | Native `iRSI()` | Custom implementation | ✅ Verified |
| ATR Calculation | Native `iATR()` | Custom implementation | ✅ Verified |
| Position Management | OrderSend/Close | MT5 API | ✅ Equivalent |
| State Tracking | Global variables | Dataclass state | ✅ Better in Python |
| Error Handling | Basic | Comprehensive | ✅ Better in Python |
| Multi-pair Support | Manual | Requires architecture | ⚠️ Needs work |
| Logging | Limited | Full logging | ✅ Better in Python |

---

## Version History

**v1.0** (Current)
- Dual timeframe RSI confirmation
- ATR-based grid spacing and take profit
- Zone permission system
- Retrace-to-first exit mechanism
- Martingale position sizing
- Independent buy/sell baskets

---

## Support & Disclaimer

**⚠️ DISCLAIMER**: This strategy involves significant risk. Past performance does not guarantee future results. Martingale strategies can lead to substantial losses. Use at your own risk. Always test thoroughly on demo accounts before live trading.

**Strategy Type**: Martingale Grid System  
**Risk Level**: HIGH  
**Recommended Experience**: Advanced traders only  
**Minimum Capital**: $5,000 per pair  

---

*For questions, optimization suggestions, or implementation support, refer to the code files:*
- `RSI  6 Trades.mq4` - MT4 Expert Advisor
- `rsi_six_trades_strategy.py` - Python MT5 Implementation
