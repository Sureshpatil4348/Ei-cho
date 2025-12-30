//+------------------------------------------------------------------+
//|                                    Tiger.mq4 |
//|                                  Copyright 2024, Trading Strategy  |
//|                                             https://www.mql5.com |
//+------------------------------------------------------------------+
#property copyright "Copyright 2024, Trading Strategy"
#property link      "https://www.mql5.com"
#property version   "1.00"
#property strict

//--- Input Parameters
input double LotSize = 0.1;                    // Lot size for trades
input int TakeProfit = 200;                   // Take Profit in points (for average TP calculation)
input double PercentageThreshold = 2.0;       // Percentage movement threshold (2%)
input int LookbackCandles = 50;               // Look back period for price analysis
input int ZScoreWaitCandles = 10;             // Candles to wait for Z-score confirmation
input double ZScoreThresholdSell = 3.0;       // Z-score threshold for SELL trades (mean reversion from high)
input double ZScoreThresholdBuy = -3.0;       // Z-score threshold for BUY trades (mean reversion from low)
input int ZScorePeriod = 20;                  // Period for Z-score calculation
input int MagicNumber = 12345;                // Magic number for trades

//--- Take Profit Parameters
input bool UseTakeProfitPercent = false;      // Use percentage-based take profit instead of points
input double TakeProfitPercent = 1.0;         // Take profit as percentage of average entry price

//--- Grid Trading Parameters
input bool UseGridTrading = true;             // Enable grid trading
input int MaxGridTrades = 5;                  // Maximum number of grid trades
input bool UseGridPercent = false;            // Use percentage-based grid spacing instead of ATR
input double GridPercent = 0.5;               // Grid spacing as percentage of price
input double GridATRMultiplier = 1.0;         // ATR multiplier for grid spacing
input int ATRPeriod = 14;                     // ATR calculation period
input double GridLotMultiplier = 1.0;         // Lot size multiplier for grid trades
input bool UseProgressiveLots = false;        // Increase lot size with each grid level
input double LotProgressionFactor = 1.5;     // Multiplication factor for progressive lots
input double MaxDrawdownPercent = 50.0;      // Maximum drawdown percentage before force close

//--- Global Variables
datetime LastBarTime = 0;
bool ConditionTriggered = false;
datetime ConditionTriggerTime = 0;
bool WaitingForZScore = false;
int WaitingDirection = 0; // 1 for sell condition, -1 for buy condition
double TriggerPrice = 0;
int ConditionCandles = 0;

//--- Grid Trading Variables
struct GridTrade {
    int ticket;
    double openPrice;
    double lotSize;
    int gridLevel;
    datetime openTime;
    int orderType;
};

GridTrade GridTrades[10];  // Array to store grid trade info
int CurrentGridLevel = 0;
double LastGridPrice = 0;
bool GridActive = false;
double InitialTradePrice = 0;
int InitialTradeTicket = 0;
int GridDirection = 0; // 1 for buy grid, -1 for sell grid
double ATRValue = 0;
double InitialAccountBalance = 0;

//+------------------------------------------------------------------+
//| Expert initialization function                                   |
//+------------------------------------------------------------------+
int OnInit()
{
    Print("ZScore Percentage Movement EA with Grid Trading initialized");
    Print("Percentage Threshold: ", PercentageThreshold, "%");
    Print("Lookback Candles: ", LookbackCandles);
    Print("Z-Score Wait Candles: ", ZScoreWaitCandles);
    Print("Z-Score Threshold SELL: ", ZScoreThresholdSell);
    Print("Z-Score Threshold BUY: ", ZScoreThresholdBuy);
    Print("Take Profit Mode: ", UseTakeProfitPercent ? "Percentage" : "Points");
    if(UseTakeProfitPercent) Print("Take Profit Percentage: ", TakeProfitPercent, "%");
    else Print("Take Profit Points: ", TakeProfit);
    Print("Grid Trading: ", UseGridTrading ? "Enabled" : "Disabled");
    Print("Max Grid Trades: ", MaxGridTrades);
    Print("Grid Spacing Mode: ", UseGridPercent ? "Percentage" : "ATR");
    if(UseGridPercent) Print("Grid Spacing Percentage: ", GridPercent, "%");
    else Print("ATR Period: ", ATRPeriod, " | ATR Multiplier: ", GridATRMultiplier);
    
    // Initialize grid variables
    InitializeGridVariables();
    InitialAccountBalance = AccountBalance();
    
    return(INIT_SUCCEEDED);
}

//+------------------------------------------------------------------+
//| Expert deinitialization function                                 |
//+------------------------------------------------------------------+
void OnDeinit(const int reason)
{
    Print("ZScore Percentage Movement EA deinitialized");
}

//+------------------------------------------------------------------+
//| Expert tick function                                             |
//+------------------------------------------------------------------+
void OnTick()
{
    // Update ATR value
    ATRValue = CalculateATR(ATRPeriod);
    
    // Check if new candle has formed
    if(Time[0] != LastBarTime)
    {
        LastBarTime = Time[0];
        
        // Main strategy logic on new candle
        CheckPercentageMovement();
        
        // Check Z-score conditions if waiting
        if(WaitingForZScore)
        {
            CheckZScoreConditions();
        }
    }
    
    // Handle grid trading if active
    if(GridActive && UseGridTrading)
    {
        MonitorGridConditions();
        CheckGridExit();
        CheckMaxDrawdown();
    }
}

//+------------------------------------------------------------------+
//| Check for percentage movement conditions                         |
//+------------------------------------------------------------------+
void CheckPercentageMovement()
{
    if(Bars < LookbackCandles + 2) return; // Need at least LookbackCandles + 2 bars
    
    // Get current close price
    double currentClose = Close[1]; // Previous candle close
    
    // Find highest and lowest close in last 50 candles
    double highestClose = 0;
    double lowestClose = 999999;
    
    // Ensure we don't exceed available bars
    int lookbackLimit = MathMin(LookbackCandles + 1, Bars - 1);
    
    for(int i = 2; i <= lookbackLimit; i++)
    {
        if(Close[i] > highestClose) highestClose = Close[i];
        if(Close[i] < lowestClose) lowestClose = Close[i];
    }
    
    // Calculate percentage changes
    double percentageFromHigh = ((currentClose - highestClose) / highestClose) * 100;
    double percentageFromLow = ((currentClose - lowestClose) / lowestClose) * 100;
    
    Print("Current Close: ", currentClose, " Highest: ", highestClose, " Lowest: ", lowestClose);
    Print("% from High: ", percentageFromHigh, " % from Low: ", percentageFromLow);
    
    // Check for +2% condition (market moved up significantly)
    if(percentageFromLow >= PercentageThreshold && !WaitingForZScore)
    {
        Print("Market moved +", PercentageThreshold, "% from recent low. Waiting for Z-score > ", ZScoreThresholdSell, " to SELL");
        WaitingForZScore = true;
        WaitingDirection = 1; // Wait for sell condition
        ConditionTriggerTime = Time[0];
        TriggerPrice = currentClose;
        ConditionCandles = 0;
    }
    // Check for -2% condition (market moved down significantly) 
    else if(percentageFromHigh <= -PercentageThreshold && !WaitingForZScore)
    {
        Print("Market moved -", PercentageThreshold, "% from recent high. Waiting for Z-score < ", ZScoreThresholdBuy, " to BUY");
        WaitingForZScore = true;
        WaitingDirection = -1; // Wait for buy condition
        ConditionTriggerTime = Time[0];
        TriggerPrice = currentClose;
        ConditionCandles = 0;
    }
}

//+------------------------------------------------------------------+
//| Check Z-score conditions and execute trades                      |
//+------------------------------------------------------------------+
void CheckZScoreConditions()
{
    ConditionCandles++;
    
    // Reset if we've waited too long
    if(ConditionCandles > ZScoreWaitCandles)
    {
        Print("Z-score condition timeout. Resetting wait state.");
        WaitingForZScore = false;
        WaitingDirection = 0;
        ConditionCandles = 0;
        return;
    }
    
    // Calculate Z-score
    double zscore = CalculateZScore(ZScorePeriod);
    Print("Current Z-score: ", zscore, " (Candle ", ConditionCandles, " of ", ZScoreWaitCandles, ")");
    
    // Check sell condition (market moved up +2%, waiting for Z-score > threshold)
    if(WaitingDirection == 1 && zscore > ZScoreThresholdSell)
    {
        Print("Z-score condition met for SELL trade. Z-score: ", zscore, " (threshold: ", ZScoreThresholdSell, ")");
        if(OpenSellTrade())
        {
            WaitingForZScore = false;
            WaitingDirection = 0;
            ConditionCandles = 0;
        }
    }
    // Check buy condition (market moved down -2%, waiting for Z-score < threshold)
    else if(WaitingDirection == -1 && zscore < ZScoreThresholdBuy)
    {
        Print("Z-score condition met for BUY trade. Z-score: ", zscore, " (threshold: ", ZScoreThresholdBuy, ")");
        if(OpenBuyTrade())
        {
            WaitingForZScore = false;
            WaitingDirection = 0;
            ConditionCandles = 0;
        }
    }
}

//+------------------------------------------------------------------+
//| Calculate Z-Score                                                |
//+------------------------------------------------------------------+
double CalculateZScore(int period)
{
    if(Bars < period + 1) return 0;
    
    double currentPrice = Close[0];
    double sum = 0;
    double sumSquares = 0;
    
    // Ensure we don't exceed available bars
    int actualPeriod = MathMin(period, Bars - 1);
    
    // Calculate mean
    for(int i = 1; i <= actualPeriod; i++)
    {
        sum += Close[i];
    }
    double mean = sum / actualPeriod;
    
    // Calculate standard deviation
    for(int i = 1; i <= actualPeriod; i++)
    {
        double diff = Close[i] - mean;
        sumSquares += diff * diff;
    }
    double stdDev = MathSqrt(sumSquares / actualPeriod);
    
    // Calculate Z-score
    if(stdDev == 0) return 0;
    double zscore = (currentPrice - mean) / stdDev;
    
    return zscore;
}

//+------------------------------------------------------------------+
//| Open Buy Trade                                                   |
//+------------------------------------------------------------------+
bool OpenBuyTrade()
{
    if(GridActive) return false; // Don't open new initial trade if grid is active
    
    double price = Ask;
    // No stop loss - grid system will handle risk management
    double tp = 0; // No individual TP - will use average TP
    
    int ticket = OrderSend(Symbol(), OP_BUY, LotSize, price, 3, 0, tp, 
                          "ZScore Buy Initial", MagicNumber, 0, clrGreen);
    
    if(ticket > 0)
    {
        Print("BUY trade opened successfully. Ticket: ", ticket);
        
        // Initialize grid system
        if(UseGridTrading)
        {
            InitializeGrid(ticket, OP_BUY, price, LotSize);
        }
        return true;
    }
    else
    {
        Print("Failed to open BUY trade. Error: ", GetLastError());
        return false;
    }
}

//+------------------------------------------------------------------+
//| Open Sell Trade                                                  |
//+------------------------------------------------------------------+
bool OpenSellTrade()
{
    if(GridActive) return false; // Don't open new initial trade if grid is active
    
    double price = Bid;
    // No stop loss - grid system will handle risk management
    double tp = 0; // No individual TP - will use average TP
    
    int ticket = OrderSend(Symbol(), OP_SELL, LotSize, price, 3, 0, tp, 
                          "ZScore Sell Initial", MagicNumber, 0, clrRed);
    
    if(ticket > 0)
    {
        Print("SELL trade opened successfully. Ticket: ", ticket);
        
        // Initialize grid system
        if(UseGridTrading)
        {
            InitializeGrid(ticket, OP_SELL, price, LotSize);
        }
        return true;
    }
    else
    {
        Print("Failed to open SELL trade. Error: ", GetLastError());
        return false;
    }
}

//+------------------------------------------------------------------+
//| Count Open Trades                                               |
//+------------------------------------------------------------------+
int CountOpenTrades()
{
    int count = 0;
    for(int i = 0; i < OrdersTotal(); i++)
    {
        if(OrderSelect(i, SELECT_BY_POS) && OrderSymbol() == Symbol() && 
           OrderMagicNumber() == MagicNumber)
        {
            count++;
        }
    }
    return count;
}

//+------------------------------------------------------------------+
//| Initialize Grid Variables                                        |
//+------------------------------------------------------------------+
void InitializeGridVariables()
{
    for(int i = 0; i < 10; i++)
    {
        GridTrades[i].ticket = 0;
        GridTrades[i].openPrice = 0;
        GridTrades[i].lotSize = 0;
        GridTrades[i].gridLevel = 0;
        GridTrades[i].openTime = 0;
        GridTrades[i].orderType = -1;
    }
    CurrentGridLevel = 0;
    LastGridPrice = 0;
    GridActive = false;
    InitialTradePrice = 0;
    InitialTradeTicket = 0;
    GridDirection = 0;
}

//+------------------------------------------------------------------+
//| Initialize Grid System                                           |
//+------------------------------------------------------------------+
void InitializeGrid(int ticket, int orderType, double price, double lotSize)
{
    GridActive = true;
    InitialTradeTicket = ticket;
    InitialTradePrice = price;
    LastGridPrice = price;
    GridDirection = (orderType == OP_BUY) ? 1 : -1;
    CurrentGridLevel = 0;
    
    // Store initial trade in grid array
    GridTrades[0].ticket = ticket;
    GridTrades[0].openPrice = price;
    GridTrades[0].lotSize = lotSize;
    GridTrades[0].gridLevel = 0;
    GridTrades[0].openTime = TimeCurrent();
    GridTrades[0].orderType = orderType;
    
    Print("Grid system initialized. Direction: ", (GridDirection == 1) ? "BUY" : "SELL", 
          " | Initial Price: ", price, " | ATR: ", ATRValue);
}

//+------------------------------------------------------------------+
//| Calculate ATR (Average True Range)                               |
//+------------------------------------------------------------------+
double CalculateATR(int period)
{
    if(Bars < period + 2) return 0.001; // Need at least period + 2 bars for ATR calculation
    
    double atrSum = 0;
    
    // Ensure we don't exceed available bars (need i+1 for Close[i+1])
    int actualPeriod = MathMin(period, Bars - 2);
    
    for(int i = 1; i <= actualPeriod; i++)
    {
        double tr1 = High[i] - Low[i];
        double tr2 = MathAbs(High[i] - Close[i+1]);
        double tr3 = MathAbs(Low[i] - Close[i+1]);
        
        double trueRange = MathMax(tr1, MathMax(tr2, tr3));
        atrSum += trueRange;
    }
    
    return atrSum / actualPeriod;
}

//+------------------------------------------------------------------+
//| Monitor Grid Conditions                                          |
//+------------------------------------------------------------------+
void MonitorGridConditions()
{
    // Safety check: ensure we don't exceed array bounds or max grid trades
    if(CurrentGridLevel >= MaxGridTrades || CurrentGridLevel >= 9) return;
    
    double currentPrice = (GridDirection == 1) ? Bid : Ask;
    double gridDistance = 0;
    
    // Calculate grid distance based on user preference
    if(UseGridPercent) // Use percentage-based grid spacing
    {
        gridDistance = LastGridPrice * (GridPercent / 100.0);
    }
    else // Use ATR-based grid spacing
    {
        gridDistance = ATRValue * GridATRMultiplier;
    }
    
    bool shouldAddGrid = false;
    
    if(GridDirection == 1) // BUY grid - add on price drops
    {
        if(currentPrice <= LastGridPrice - gridDistance)
        {
            shouldAddGrid = true;
        }
    }
    else // SELL grid - add on price rises
    {
        if(currentPrice >= LastGridPrice + gridDistance)
        {
            shouldAddGrid = true;
        }
    }
    
    if(shouldAddGrid)
    {
        AddGridTrade(GridDirection, currentPrice);
    }
}

//+------------------------------------------------------------------+
//| Add Grid Trade                                                   |
//+------------------------------------------------------------------+
bool AddGridTrade(int direction, double currentPrice)
{
    // Safety check: ensure we don't exceed array bounds (GridTrades[10] = indices 0-9)
    if(CurrentGridLevel >= MaxGridTrades || CurrentGridLevel >= 9) return false;
    
    // Calculate lot size for grid trade
    double gridLotSize = LotSize * GridLotMultiplier;
    if(UseProgressiveLots)
    {
        gridLotSize = LotSize * MathPow(LotProgressionFactor, CurrentGridLevel + 1);
    }
    
    int orderType = (direction == 1) ? OP_BUY : OP_SELL;
    double price = (direction == 1) ? Ask : Bid;
    
    string comment = StringConcatenate("Grid ", (CurrentGridLevel + 1), 
                                      (direction == 1) ? " Buy" : " Sell");
    
    int ticket = OrderSend(Symbol(), orderType, gridLotSize, price, 3, 0, 0, 
                          comment, MagicNumber, 0, 
                          (direction == 1) ? clrGreen : clrRed);
    
    if(ticket > 0)
    {
        CurrentGridLevel++;
        LastGridPrice = price;
        
        // Ensure array bounds safety
        if(CurrentGridLevel < 10)
        {
            // Store grid trade info
            GridTrades[CurrentGridLevel].ticket = ticket;
            GridTrades[CurrentGridLevel].openPrice = price;
            GridTrades[CurrentGridLevel].lotSize = gridLotSize;
            GridTrades[CurrentGridLevel].gridLevel = CurrentGridLevel;
            GridTrades[CurrentGridLevel].openTime = TimeCurrent();
            GridTrades[CurrentGridLevel].orderType = orderType;
        }
        
        Print("Grid trade ", CurrentGridLevel, " added. Ticket: ", ticket, 
              " | Price: ", price, " | Lot: ", gridLotSize);
        
        return true;
    }
    else
    {
        Print("Failed to add grid trade. Error: ", GetLastError());
        return false;
    }
}

//+------------------------------------------------------------------+
//| Calculate Average Take Profit Price                              |
//+------------------------------------------------------------------+
double CalculateAverageTakeProfit()
{
    double totalLots = 0;
    double weightedPrice = 0;
    
    // Calculate weighted average entry price
    int maxLevel = MathMin(CurrentGridLevel, 9); // Ensure array bounds safety
    for(int i = 0; i <= maxLevel; i++)
    {
        if(GridTrades[i].ticket > 0)
        {
            totalLots += GridTrades[i].lotSize;
            weightedPrice += GridTrades[i].openPrice * GridTrades[i].lotSize;
        }
    }
    
    if(totalLots == 0) return 0;
    
    double averagePrice = weightedPrice / totalLots;
    double targetPrice = 0;
    
    if(UseTakeProfitPercent) // Use percentage-based take profit
    {
        if(GridDirection == 1) // BUY positions
        {
            targetPrice = averagePrice * (1 + TakeProfitPercent / 100.0);
        }
        else // SELL positions
        {
            targetPrice = averagePrice * (1 - TakeProfitPercent / 100.0);
        }
    }
    else // Use points-based take profit
    {
        if(GridDirection == 1) // BUY positions
        {
            targetPrice = averagePrice + (TakeProfit * Point);
        }
        else // SELL positions
        {
            targetPrice = averagePrice - (TakeProfit * Point);
        }
    }
    
    return targetPrice;
}

//+------------------------------------------------------------------+
//| Check Grid Exit Conditions                                       |
//+------------------------------------------------------------------+
void CheckGridExit()
{
    double averageTP = CalculateAverageTakeProfit();
    double currentPrice = (GridDirection == 1) ? Bid : Ask;
    
    bool shouldExit = false;
    
    if(GridDirection == 1 && currentPrice >= averageTP) // BUY grid profit target reached
    {
        shouldExit = true;
        Print("BUY grid profit target reached. Current: ", currentPrice, " | Target: ", averageTP);
    }
    else if(GridDirection == -1 && currentPrice <= averageTP) // SELL grid profit target reached
    {
        shouldExit = true;
        Print("SELL grid profit target reached. Current: ", currentPrice, " | Target: ", averageTP);
    }
    else if(CurrentGridLevel >= MaxGridTrades) // Max grid trades reached
    {
        shouldExit = true;
        Print("Maximum grid trades reached. Force closing all positions.");
    }
    
    if(shouldExit)
    {
        CloseAllGridTrades();
    }
}

//+------------------------------------------------------------------+
//| Close All Grid Trades                                            |
//+------------------------------------------------------------------+
void CloseAllGridTrades()
{
    Print("Closing all grid trades...");
    
    int maxLevel = MathMin(CurrentGridLevel, 9); // Ensure array bounds safety
    for(int i = 0; i <= maxLevel; i++)
    {
        if(GridTrades[i].ticket > 0)
        {
            if(OrderSelect(GridTrades[i].ticket, SELECT_BY_TICKET))
            {
                double closePrice = (OrderType() == OP_BUY) ? Bid : Ask;
                bool closed = OrderClose(GridTrades[i].ticket, OrderLots(), closePrice, 3, clrYellow);
                
                if(closed)
                {
                    Print("Grid trade closed. Ticket: ", GridTrades[i].ticket, 
                          " | Level: ", GridTrades[i].gridLevel);
                }
                else
                {
                    Print("Failed to close grid trade. Ticket: ", GridTrades[i].ticket, 
                          " | Error: ", GetLastError());
                }
            }
        }
    }
    
    // Reset grid system
    InitializeGridVariables();
    Print("Grid system reset. Ready for next signal.");
}

//+------------------------------------------------------------------+
//| Check Maximum Drawdown                                           |
//+------------------------------------------------------------------+
void CheckMaxDrawdown()
{
    double currentBalance = AccountBalance();
    double currentEquity = AccountEquity();
    double drawdownPercent = ((InitialAccountBalance - currentEquity) / InitialAccountBalance) * 100;
    
    if(drawdownPercent > MaxDrawdownPercent)
    {
        Print("Maximum drawdown exceeded: ", drawdownPercent, "%. Force closing all positions.");
        CloseAllGridTrades();
    }
} 