# Gold Buy Dip Strategy Documentation

## Overview

The "Gold Buy Dip" is a sophisticated, mean-reversion trading strategy designed primarily for assets like Gold (XAU/USD), but adaptable to other instruments. The core idea is to identify significant price movements away from a recent range and then enter a trade in the opposite direction, anticipating a correction back towards the mean.

The strategy combines two main signals for trade entry: a significant percentage-based price move followed by a Z-score confirmation. It incorporates a robust grid trading system to manage positions and a portfolio-level maximum drawdown check for risk management.

## Recommended Lot Sizing

For effective risk management, it is recommended to use a **lot size of 0.01 for every $1000 of capital** in the trading account. This helps to ensure that the risk per trade is a small fraction of the total equity, which is crucial when employing a grid-based system.

---

## Core Trading Logic

The strategy's logic for entering a trade is a two-step process designed to filter for high-probability mean-reversion scenarios.

### 1. Percentage Movement Trigger

First, the EA waits for the price to make a significant move away from its recent high or low. It does this by:
1.  Identifying the highest and lowest closing prices over a `LookbackCandles` period (e.g., 50 candles).
2.  Calculating the percentage change of the current price from that highest high and lowest low.
3.  If the price moves up by more than the `PercentageThreshold` (e.g., 2%) from the recent low, it triggers a **potential SELL setup**.
4.  If the price moves down by more than the `PercentageThreshold` (e.g., 2%) from the recent high, it triggers a **potential BUY setup**.

When a trigger occurs, the EA does not open a trade immediately. Instead, it enters a "waiting" state for Z-Score confirmation.

### 2. Z-Score Confirmation

After the percentage movement trigger, the EA waits for a limited number of candles (`ZScoreWaitCandles`) for a Z-score confirmation. The Z-score measures how far the current price is from the recent average price in terms of standard deviations.

-   For a **SELL setup** (after a +2% move up), the EA waits for the Z-score to become highly positive (e.g., `> 3.0`), indicating the price is statistically overbought and likely to revert lower.
-   For a **BUY setup** (after a -2% move down), the EA waits for the Z-score to become highly negative (e.g., `< -3.0`), indicating the price is statistically oversold and likely to revert higher.

If the Z-score crosses its threshold within the waiting period, the initial trade is executed. If not, the setup is invalidated, and the EA waits for a new percentage movement trigger.

---

## Indicators Used

The strategy relies on custom calculations and standard statistical indicators to make trading decisions.

### Z-Score

The Z-score is a statistical measure that tells you how many standard deviations a data point is from the mean (average) of a set of data.

**Formula:**
\[ Z = \frac{(P - \mu)}{\sigma} \]
Where:
-   \( P \) = Current Price (`Close[0]`)
-   \( \mu \) = Mean (average) of closing prices over the `ZScorePeriod`
-   \( \sigma \) = Standard Deviation of closing prices over the `ZScorePeriod`

A high positive Z-score suggests the price is unusually high compared to its recent average, while a low negative Z-score suggests it is unusually low.

### Average True Range (ATR)

The ATR is used to determine the distance between grid trades. It measures market volatility.

**Calculation Steps:**
1.  **True Range (TR):** For each candle, the TR is the greatest of the following:
    -   `High - Low`
    -   `|High - Previous Close|`
    -   `|Low - Previous Close|`
2.  **Average True Range (ATR):** The ATR is the moving average of the True Range values over the `ATRPeriod`.

A higher ATR value means higher volatility, which will result in wider spacing between grid trades.

---

## Grid Trading System

If `UseGridTrading` is enabled, the EA will manage the initial trade as part of a grid.

### Grid Initiation

-   The grid is initiated when the first trade is opened based on the core logic.
-   The EA tracks the initial trade's price and direction.

### Adding Grid Trades

-   If the market moves against the initial position, the EA will open additional trades.
-   **Grid Spacing:** The distance between grid trades is determined by either:
    -   **ATR:** `ATR Value * GridATRMultiplier` (e.g., 1.0 * ATR).
    -   **Percentage:** `Last Trade Price * (GridPercent / 100)`.
-   **Lot Sizing:** The lot size for subsequent grid trades can be:
    -   A simple multiplier of the initial lot (`GridLotMultiplier`).
    -   Progressively increased using the `LotProgressionFactor`.
-   The EA will continue adding trades up to `MaxGridTrades`.

### Grid Take Profit

-   Individual trades in the grid do not have their own Take Profit.
-   The EA calculates a **single Take Profit level for the entire grid**.
-   This level is calculated based on the **volume-weighted average price** of all open trades.
-   The target profit is then added to this average price, based on either `TakeProfit` (in points) or `TakeProfitPercent`.
-   When the market price reaches this calculated level, all trades in the grid are closed simultaneously.

---

## Risk Management

### Maximum Drawdown

The primary safety mechanism is the `MaxDrawdownPercent` input.
-   The EA monitors the account's equity.
-   If the floating loss (drawdown) of the account exceeds the specified percentage of the initial balance, the EA will **force-close all open grid positions**.
-   This acts as an ultimate stop-loss for the entire strategy, preventing catastrophic losses.

---

## Input Parameters

| Parameter               | Description                                                                                             | Default Value   |
| ----------------------- | ------------------------------------------------------------------------------------------------------- | --------------- |
| `LotSize`               | The initial lot size for the first trade in a series.                                                   | `0.1`           |
| `TakeProfit`            | The take profit in points, applied to the average price of the grid. Used if `UseTakeProfitPercent` is false. | `200`           |
| `PercentageThreshold`   | The percentage move from a recent high/low required to trigger a trade setup.                           | `2.0`           |
| `LookbackCandles`       | The number of past candles to analyze for the highest/lowest price for the percentage movement trigger.     | `50`            |
| `ZScoreWaitCandles`     | The maximum number of candles to wait for a Z-score confirmation after a percentage trigger.              | `10`            |
| `ZScoreThresholdSell`   | The Z-score value that must be exceeded to confirm a SELL trade.                                        | `3.0`           |
| `ZScoreThresholdBuy`    | The Z-score value that the price must fall below to confirm a BUY trade.                                | `-3.0`          |
| `ZScorePeriod`          | The number of candles used to calculate the Z-score's mean and standard deviation.                      | `20`            |
| `MagicNumber`           | A unique ID to ensure the EA only manages its own trades.                                               | `12345`         |
| `UseTakeProfitPercent`  | If `true`, `TakeProfitPercent` is used for the TP calculation; otherwise, `TakeProfit` (points) is used. | `false`         |
| `TakeProfitPercent`     | The take profit as a percentage of the average entry price.                                             | `1.0`           |
| `UseGridTrading`        | Enables or disables the grid trading functionality.                                                     | `true`          |
| `MaxGridTrades`         | The maximum number of trades allowed in a single grid series.                                           | `5`             |
| `UseGridPercent`        | If `true`, grid spacing is based on `GridPercent`; otherwise, it's based on ATR.                        | `false`         |
| `GridPercent`           | The distance between grid trades as a percentage of the last trade's price.                             | `0.5`           |
| `GridATRMultiplier`     | A multiplier for the ATR value to determine grid spacing.                                               | `1.0`           |
| `ATRPeriod`             | The period for calculating the Average True Range (ATR).                                                | `14`            |
| `GridLotMultiplier`     | A simple multiplier for the lot size of subsequent grid trades. Ignored if `UseProgressiveLots` is true. | `1.0`           |
| `UseProgressiveLots`    | If `true`, the lot size of each grid trade increases by the `LotProgressionFactor`.                     | `false`         |
| `LotProgressionFactor`  | The factor by which the lot size is multiplied for each new grid level (e.g., 1.5).                     | `1.5`           |
| `MaxDrawdownPercent`    | The maximum allowed account drawdown percentage before all trades are force-closed.                     | `50.0`          |

