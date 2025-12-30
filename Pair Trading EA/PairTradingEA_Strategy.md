# PairTradingEA Strategy Explanation

This document provides a detailed breakdown of the trading strategy implemented in `PairTradingEA.mq5`. The Expert Advisor (EA) is designed for **manual entry pair trading** with an **automated recovery system** and **global profit management**.

---

## 1. Overview
The **PairTradingEA** allows you to simultaneously trade two different currency pairs (or assets) using a custom dashboard. Once you initiate a trade, the EA manages the position automatically by:
1.  Monitoring the **Global P/L** (Profit/Loss) of all open trades.
2.  Closing all trades when a specific **Profit Target** is reached.
3.  Automatically adding new "Recovery Trades" if the position goes into drawdown, based on a specific mathematical sequence.

---

## 2. Dashboard Inputs & Configuration
The EA features an on-chart dashboard that allows you to configure the strategy before placing a trade.

| Input Field | Description |
| :--- | :--- |
| **Symbol 1** | The first currency pair to trade (e.g., `EURUSD`). |
| **Symbol 2** | The second currency pair to trade (e.g., `GBPUSD`). |
| **Lot Size 1** | The fixed volume for Symbol 1 per trade. |
| **Lot Size 2** | The fixed volume for Symbol 2 per trade. |
| **Select Strategy** | Determine the direction of trade for the pair (see Section 3). |
| **Take Profit ($)** | The combined profit target in dollars. Example: `$100`. |
| **Loss to Repeat ($)** | The base loss amount used to calculate when to add recovery trades. Example: `$50`. |

---

## 3. Strategy Modes (Direction)
You can choose how the two symbols interact. This allows for correlation or hedging strategies.

1.  **Buy Symbol 1, Sell Symbol 2**: Standard pair trading (betting on spread convergence/divergence).
2.  **Sell Symbol 1, Buy Symbol 2**: Inverse of strategy 1.
3.  **Buy Symbol 1, Buy Symbol 2**: Long exposure on both assets.
4.  **Sell Symbol 1, Sell Symbol 2**: Short exposure on both assets.

---

## 4. Workflows & Logic

### A. Entry (Manual Trigger)
*   **Action**: You click the **"PLACE TRADES"** button on the dashboard.
*   **Logic**:
    1.  The EA validates that all inputs are correct (valid symbols, lot sizes > 0).
    2.  If any previous trades are active, they are closed immediately.
    3.  The **Initial Trade Set (Set #0)** is opened immediately for both symbols according to the selected strategy.
    4.  Monitoring begins.

### B. Exit: Take Profit (Automated)
*   **Condition**: The EA constantly sums the Profit + Swap of all open trades under its management.
*   **Trigger**: When `Total P/L >= Take Profit ($)`.
*   **Action**:
    *   **Close All Trades** immediately.
    *   Reset the trade counter and loss levels.
    *   Alert the user: "Take Profit achieved!"

### C. Manual Exit
*   **Action**: You click the **"CLOSE ALL TRADES"** button.
*   **Logic**: Forces an immediate closure of all positions and resets the EA state.

---

## 5. Recovery Mechanism (The "Grid")
This is the most complex part of the strategy. If the trade moves against you, the EA does **not** use a Stop Loss. Instead, it adds *more* trades (same direction and lot size) at specific accumulating loss levels to average out the entry price.

### The Formula
The EA uses a **Geometric Progression** based on the "Loss to Repeat" value to determine when to add the next set of trades.

**Formula for Next Trigger Level:**
```math
Threshold = LossAmount * (2^n - 1)
```
*Where `n` is the next trade set number (1, 2, 3...).*

### Step-by-Step Example
*   **Scenario**:
    *   `Loss to Repeat` = **$100**
    *   `Take Profit` = **$100**
    *   `Initial Trade` (Set 0) is open.

1.  **First Recovery Level (Set #2)**
    *   Trigger: When Drawdown hits **$100** `(100 * (2^1 - 1))`.
    *   Action: Open another set of trades (Same lots as initial).
    *   *Total Sets Running: 2*

2.  **Second Recovery Level (Set #3)**
    *   Trigger: When Drawdown hits **$300** `(100 * (2^2 - 1))`.
    *   *Note: This is $200 deeper than the previous level.*
    *   Action: Open another set of trades.
    *   *Total Sets Running: 3*

3.  **Third Recovery Level (Set #4)**
    *   Trigger: When Drawdown hits **$700** `(100 * (2^3 - 1))`.
    *   *Note: This is $400 deeper than the previous level.*
    *   Action: Open another set of trades.
    *   *Total Sets Running: 4*

4.  **Fourth Recovery Level (Set #5)**
    *   Trigger: When Drawdown hits **$1500** `(100 * (2^4 - 1))`.
    *   *Note: This is $800 deeper than the previous level.*

**Summary of Intervals:** The gap between trade additions doubles each time ($100, $200, $400, $800...). This approach gives trades more "breathing room" as the drawdown increases, rather than stacking trades too closely.

---

## 6. Important Technical Details
*   **Magic Number**: The EA uses ID `123456`. It will only manage and calculate P/L for trades with match this specific ID.
*   **Timer**: The dashboard updates P/L every 1 second (and on every price tick).
*   **Swap Costs**: The P/L calculation **includes Swap**, so overnight holding costs are factored into the Take Profit and Loss Trigger targets.
