//+------------------------------------------------------------------+
//|                                          RSI_Pyramiding_EA.mq4   |
//|                RSI Pyramiding Strategy - BUY/SELL/BOTH           |
//|                                   MetaTrader 4 Expert Advisor    |
//+------------------------------------------------------------------+
#property copyright "RSI Pyramiding Strategy"
#property link      ""
#property version   "1.00"
#property strict

//+------------------------------------------------------------------+
//| Input Parameters - ALL CONFIGURABLE                              |
//+------------------------------------------------------------------+
input group "========== Trend Filter ==========";
input bool     Use_EMA_Filter = true;              // Use 200 EMA Trend Filter
input int      EMA_Period = 200;                   // EMA Period
input ENUM_TIMEFRAMES EMA_Timeframe = PERIOD_H1;   // EMA Timeframe
input ENUM_APPLIED_PRICE EMA_Price = PRICE_CLOSE;  // EMA Applied Price

input group "========== RSI Settings ==========";
input bool     Use_RSI_Filter = true;              // Use RSI for Entry Signal
input int      RSI_Period = 14;                    // RSI Period
input int      RSI_Oversold = 30;                  // RSI Oversold Level
input int      RSI_Overbought = 70;                // RSI Overbought Level
input ENUM_TIMEFRAMES RSI_Timeframe = PERIOD_CURRENT; // RSI Timeframe
input ENUM_APPLIED_PRICE RSI_Price = PRICE_CLOSE;  // RSI Applied Price

input group "========== Trade Direction ==========";
enum ENUM_TRADE_DIRECTION
{
   TRADE_BUY_ONLY = 0,    // Buy Only
   TRADE_SELL_ONLY = 1,   // Sell Only
   TRADE_BOTH = 2         // Both
};
input ENUM_TRADE_DIRECTION Trade_Direction = TRADE_BUY_ONLY; // Trade Direction

input group "========== Pyramiding Settings ==========";
input ENUM_TIMEFRAMES Pyramiding_Timeframe = PERIOD_CURRENT; // Pyramiding Timeframe (Candle Close Check)

input group "========== Position Sizing ==========";
input double   Lot_Size = 0.01;                    // Lot Size per Trade

input group "========== Risk Management ==========";
input double   First_Trade_SL = 50.0;              // First Trade SL (in Price, e.g., $50)
input double   Min_Profit_For_Second = 30.0;       // Min Profit to Open Trade 2 ($)
input double   Combined_Profit_Protection = 5.0;   // Combined Profit Min Protection ($)

input group "========== Targets & Limits ==========";
input double   Profit_Target = 500.0;              // Profit Target to Close All ($)
input int      Max_Trades = 20;                    // Maximum Simultaneous Trades

input group "========== System Settings ==========";
input int      Magic_Number = 123456;              // Magic Number
input string   Comment_Text = "RSI_Pyramid";       // Trade Comment
input bool     Show_Dashboard = true;              // Show Dashboard on Chart
input color    Profit_Color = clrLime;             // Profit Text Color
input color    Loss_Color = clrRed;                // Loss Text Color
input int      Dashboard_X = 20;                   // Dashboard X Position
input int      Dashboard_Y = 30;                   // Dashboard Y Position

//+------------------------------------------------------------------+
//| Global Variables                                                  |
//+------------------------------------------------------------------+
// Buy Side Variables
double g_previousRSI = 0;           // Previous bar RSI value
bool g_rsiWasBelowOversold = false; // Tracks if RSI was below oversold level
int g_firstTradeTicket = -1;        // Ticket number of first buy trade
double g_firstTradePrice = 0;       // Entry price of first buy trade
bool g_firstTradeSLRemoved = false; // Whether first buy trade SL was removed
double g_lastBuyPrice = 0;          // Price of most recent buy trade
int g_totalTrades = 0;              // Count of open buy trades

// Sell Side Variables
bool g_rsiWasAboveOverbought = false; // Tracks if RSI was above overbought level
int g_firstSellTradeTicket = -1;     // Ticket number of first sell trade
double g_firstSellTradePrice = 0;    // Entry price of first sell trade
bool g_firstSellTradeSLRemoved = false; // Whether first sell trade SL was removed
double g_lastSellPrice = 0;          // Price of most recent sell trade
int g_totalSellTrades = 0;           // Count of open sell trades

// Dashboard object names
string g_dashPrefix = "RSI_Dash_";

//+------------------------------------------------------------------+
//| Expert initialization function                                    |
//+------------------------------------------------------------------+
int OnInit()
{
   // Validate inputs
   if(RSI_Period < 5 || RSI_Period > 50)
   {
      Alert("RSI Period must be between 5 and 50");
      return(INIT_PARAMETERS_INCORRECT);
   }
   
   if(RSI_Oversold < 10 || RSI_Oversold > 40)
   {
      Alert("RSI Oversold level must be between 10 and 40");
      return(INIT_PARAMETERS_INCORRECT);
   }
   
   if(RSI_Overbought < 60 || RSI_Overbought > 90)
   {
      Alert("RSI Overbought level must be between 60 and 90");
      return(INIT_PARAMETERS_INCORRECT);
   }
   
   if(Lot_Size < 0.01 || Lot_Size > 10.0)
   {
      Alert("Lot Size must be between 0.01 and 10.0");
      return(INIT_PARAMETERS_INCORRECT);
   }
   
   if(Max_Trades < 2 || Max_Trades > 50)
   {
      Alert("Max Trades must be between 2 and 50");
      return(INIT_PARAMETERS_INCORRECT);
   }
   
   if(Use_EMA_Filter && (EMA_Period < 10 || EMA_Period > 500))
   {
      Alert("EMA Period must be between 10 and 500");
      return(INIT_PARAMETERS_INCORRECT);
   }
   
   // Initialize state
   g_previousRSI = 0;
   g_rsiWasBelowOversold = false;
   g_firstTradeTicket = -1;
   g_firstTradePrice = 0;
   g_firstTradeSLRemoved = false;
   g_lastBuyPrice = 0;
   
   // Initialize sell state
   g_rsiWasAboveOverbought = false;
   g_firstSellTradeTicket = -1;
   g_firstSellTradePrice = 0;
   g_firstSellTradeSLRemoved = false;
   g_lastSellPrice = 0;
   
   // Count existing trades (in case EA is restarted)
   UpdateTradeState();
   
   Print("RSI Pyramiding EA initialized successfully");
   Print("Magic Number: ", Magic_Number);
   Print("Max Trades: ", Max_Trades);
   Print("Profit Target: $", Profit_Target);
   
   return(INIT_SUCCEEDED);
}

//+------------------------------------------------------------------+
//| Expert deinitialization function                                  |
//+------------------------------------------------------------------+
void OnDeinit(const int reason)
{
   // Remove dashboard objects
   if(Show_Dashboard)
      RemoveDashboard();
   
   Print("RSI Pyramiding EA deinitialized. Reason: ", reason);
}

//+------------------------------------------------------------------+
//| Chart Event handler                                               |
//+------------------------------------------------------------------+
void OnChartEvent(const int id, const long &lparam, const double &dparam, const string &sparam)
{
   // Check for button click event
   if(id == CHARTEVENT_OBJECT_CLICK)
   {
      // Check if Close All Trades button was clicked
      if(sparam == g_dashPrefix + "CloseAllBtn")
      {
         // Reset button state
         ObjectSetInteger(0, g_dashPrefix + "CloseAllBtn", OBJPROP_STATE, false);
         
         // Close all trades
         CloseAllTradesManual();
      }
   }
}

//+------------------------------------------------------------------+
//| Expert tick function                                              |
//+------------------------------------------------------------------+
void OnTick()
{
   // Update trade state
   UpdateTradeState();
   
   // === BUY SIDE LOGIC ===
   if(Trade_Direction == TRADE_BUY_ONLY || Trade_Direction == TRADE_BOTH)
   {
      // Check exit conditions first
      if(g_totalTrades > 0)
      {
         CheckExitConditions();
      }
      
      // Check for new buy entries
      if(g_totalTrades == 0)
      {
         // Look for initial entry signal
         CheckInitialEntry();
      }
      else if(g_totalTrades == 1)
      {
         // Check for second trade entry
         CheckSecondTradeEntry();
      }
      else if(g_totalTrades < Max_Trades)
      {
         // Check for pyramiding entry (trades 3-20)
         CheckPyramidingEntry();
      }
      else if(g_totalTrades >= Max_Trades)
      {
         // Check for FIFO entry (close oldest, open new)
         CheckFIFOEntry();
      }
   }
   
   // === SELL SIDE LOGIC ===
   if(Trade_Direction == TRADE_SELL_ONLY || Trade_Direction == TRADE_BOTH)
   {
      // Check exit conditions first
      if(g_totalSellTrades > 0)
      {
         CheckSellExitConditions();
      }
      
      // Check for new sell entries
      if(g_totalSellTrades == 0)
      {
         // Look for initial sell entry signal
         CheckInitialSellEntry();
      }
      else if(g_totalSellTrades == 1)
      {
         // Check for second sell trade entry
         CheckSecondSellTradeEntry();
      }
      else if(g_totalSellTrades < Max_Trades)
      {
         // Check for pyramiding entry (trades 3-20)
         CheckSellPyramidingEntry();
      }
      else if(g_totalSellTrades >= Max_Trades)
      {
         // Check for FIFO entry (close oldest, open new)
         CheckSellFIFOEntry();
      }
   }
   
   // Update dashboard
   if(Show_Dashboard)
      UpdateDashboard();
}

//+------------------------------------------------------------------+
//| Update trade state and variables                                  |
//+------------------------------------------------------------------+
void UpdateTradeState()
{
   g_totalTrades = 0;
   g_firstTradeTicket = -1;
   g_firstTradePrice = 0;
   g_lastBuyPrice = 0;
   
   g_totalSellTrades = 0;
   g_firstSellTradeTicket = -1;
   g_firstSellTradePrice = 0;
   g_lastSellPrice = 0;
   
   int oldestBuyTicket = -1;
   int oldestSellTicket = -1;
   double highestBuyPrice = 0;
   double lowestSellPrice = 0;
   
   for(int i = OrdersTotal() - 1; i >= 0; i--)
   {
      if(OrderSelect(i, SELECT_BY_POS, MODE_TRADES))
      {
         if(OrderSymbol() == Symbol() && OrderMagicNumber() == Magic_Number)
         {
            // Track BUY trades
            if(OrderType() == OP_BUY)
            {
               g_totalTrades++;
               
               // Track first buy trade (oldest ticket with lowest number)
               if(oldestBuyTicket == -1 || OrderTicket() < oldestBuyTicket)
               {
                  oldestBuyTicket = OrderTicket();
                  g_firstTradeTicket = OrderTicket();
                  g_firstTradePrice = OrderOpenPrice();
               }
               
               // Track last buy price (highest entry price)
               if(OrderOpenPrice() > highestBuyPrice)
               {
                  highestBuyPrice = OrderOpenPrice();
                  g_lastBuyPrice = OrderOpenPrice();
               }
            }
            // Track SELL trades
            else if(OrderType() == OP_SELL)
            {
               g_totalSellTrades++;
               
               // Track first sell trade (oldest ticket with lowest number)
               if(oldestSellTicket == -1 || OrderTicket() < oldestSellTicket)
               {
                  oldestSellTicket = OrderTicket();
                  g_firstSellTradeTicket = OrderTicket();
                  g_firstSellTradePrice = OrderOpenPrice();
               }
               
               // Track last sell price (lowest entry price)
               if(lowestSellPrice == 0 || OrderOpenPrice() < lowestSellPrice)
               {
                  lowestSellPrice = OrderOpenPrice();
                  g_lastSellPrice = OrderOpenPrice();
               }
            }
         }
      }
   }
   
   // Check if first buy trade SL was removed
   if(g_firstTradeTicket > 0 && OrderSelect(g_firstTradeTicket, SELECT_BY_TICKET))
   {
      g_firstTradeSLRemoved = (OrderStopLoss() == 0);
   }
   
   // Check if first sell trade SL was removed
   if(g_firstSellTradeTicket > 0 && OrderSelect(g_firstSellTradeTicket, SELECT_BY_TICKET))
   {
      g_firstSellTradeSLRemoved = (OrderStopLoss() == 0);
   }
}

//+------------------------------------------------------------------+
//| Check for initial entry signal (Trade 1)                          |
//+------------------------------------------------------------------+
void CheckInitialEntry()
{
   // Check EMA filter first (if enabled)
   if(Use_EMA_Filter)
   {
      if(!IsMarketAboveEMA())
      {
         // Market not above EMA - don't check signals
         g_rsiWasBelowOversold = false;  // Reset RSI state
         return;
      }
   }
   
   // MODE 1: RSI-based entry (original strategy)
   if(Use_RSI_Filter)
   {
      // Calculate current RSI on selected timeframe
      double currentRSI = iRSI(Symbol(), RSI_Timeframe, RSI_Period, RSI_Price, 0);
      double previousRSI = iRSI(Symbol(), RSI_Timeframe, RSI_Period, RSI_Price, 1);
      
      // Check if RSI was below oversold
      if(previousRSI < RSI_Oversold)
      {
         g_rsiWasBelowOversold = true;
      }
      
      // Entry signal: RSI crosses from below oversold to above
      if(g_rsiWasBelowOversold && currentRSI > RSI_Oversold)
      {
         // Open first trade
         double price = Ask;
         double sl = price - First_Trade_SL;  // Price-based SL
         double tp = 0;  // No TP
         
         int ticket = OrderSend(Symbol(), OP_BUY, Lot_Size, price, 3, sl, tp, 
                                Comment_Text + " #1", Magic_Number, 0, clrBlue);
         
         if(ticket > 0)
         {
            Print("Trade 1 opened at ", price, " with SL at ", sl, " (RSI Entry)");
            g_firstTradeTicket = ticket;
            g_firstTradePrice = price;
            g_lastBuyPrice = price;
            g_rsiWasBelowOversold = false;  // Reset flag
            g_firstTradeSLRemoved = false;
         }
         else
         {
            Print("Error opening Trade 1: ", GetLastError());
         }
      }
      
      g_previousRSI = currentRSI;
   }
   // MODE 2: EMA-only entry (instant entry on new candle if above EMA)
   else
   {
      // Check on new bar only (on pyramiding timeframe)
      static datetime lastBarTime = 0;
      datetime currentBarTime = iTime(Symbol(), Pyramiding_Timeframe, 0);
      
      if(currentBarTime == lastBarTime)
         return;  // Wait for new bar
      
      lastBarTime = currentBarTime;
      
      // Verify EMA filter is enabled and market is above EMA
      if(!Use_EMA_Filter)
      {
         Print("Warning: RSI disabled but EMA filter also disabled. No entry conditions available.");
         return;
      }
      
      // Market is above EMA (already checked above), enter immediately
      double price = Ask;
      double sl = price - First_Trade_SL;  // Price-based SL
      double tp = 0;  // No TP
      
      int ticket = OrderSend(Symbol(), OP_BUY, Lot_Size, price, 3, sl, tp, 
                             Comment_Text + " #1", Magic_Number, 0, clrBlue);
      
      if(ticket > 0)
      {
         Print("Trade 1 opened at ", price, " with SL at ", sl, " (EMA-Only Entry)");
         g_firstTradeTicket = ticket;
         g_firstTradePrice = price;
         g_lastBuyPrice = price;
         g_firstTradeSLRemoved = false;
      }
      else
      {
         Print("Error opening Trade 1: ", GetLastError());
      }
   }
}

//+------------------------------------------------------------------+
//| Check for second trade entry                                      |
//+------------------------------------------------------------------+
void CheckSecondTradeEntry()
{
   // Verify first trade exists
   if(g_firstTradeTicket <= 0 || !OrderSelect(g_firstTradeTicket, SELECT_BY_TICKET))
      return;
   
   // Check conditions for Trade 2:
   // 1. First trade profit >= $30
   double firstTradeProfit = OrderProfit() + OrderSwap() + OrderCommission();
   
   if(firstTradeProfit < Min_Profit_For_Second)
      return;
   
   // 2. Current candle closed above first trade entry (check on new bar on pyramiding timeframe)
   static datetime lastBarTime = 0;
   datetime currentBarTime = iTime(Symbol(), Pyramiding_Timeframe, 0);
   
   if(currentBarTime == lastBarTime)
      return;  // Wait for new bar
   
   lastBarTime = currentBarTime;
   
   double closePrice = iClose(Symbol(), Pyramiding_Timeframe, 1);  // Previous bar close on pyramiding TF
   
   if(closePrice <= g_firstTradePrice)
      return;
   
   // 3. RSI not overbought (optional check on RSI timeframe)
   double currentRSI = iRSI(Symbol(), RSI_Timeframe, RSI_Period, RSI_Price, 0);
   if(currentRSI > 70)
      return;
   
   // Open Trade 2
   double price = Ask;
   int ticket = OrderSend(Symbol(), OP_BUY, Lot_Size, price, 3, 0, 0, 
                          Comment_Text + " #2", Magic_Number, 0, clrGreen);
   
   if(ticket > 0)
   {
      Print("Trade 2 opened at ", price);
      
      // Remove SL from Trade 1
      if(OrderSelect(g_firstTradeTicket, SELECT_BY_TICKET))
      {
         bool modified = OrderModify(g_firstTradeTicket, OrderOpenPrice(), 0, 0, 0, clrNONE);
         if(modified)
         {
            Print("Trade 1 Stop Loss removed");
            g_firstTradeSLRemoved = true;
         }
         else
         {
            Print("Error removing Trade 1 SL: ", GetLastError());
         }
      }
   }
   else
   {
      Print("Error opening Trade 2: ", GetLastError());
   }
}

//+------------------------------------------------------------------+
//| Check for pyramiding entry (Trades 3-20)                          |
//+------------------------------------------------------------------+
void CheckPyramidingEntry()
{
   // Check on new bar only (on pyramiding timeframe)
   static datetime lastBarTime = 0;
   datetime currentBarTime = iTime(Symbol(), Pyramiding_Timeframe, 0);
   
   if(currentBarTime == lastBarTime)
      return;
   
   lastBarTime = currentBarTime;
   
   // Check if previous candle closed above last buy level on pyramiding timeframe
   double closePrice = iClose(Symbol(), Pyramiding_Timeframe, 1);
   
   if(closePrice <= g_lastBuyPrice)
      return;
   
   // Open new pyramiding trade
   double price = Ask;
   int tradeNumber = g_totalTrades + 1;
   
   int ticket = OrderSend(Symbol(), OP_BUY, Lot_Size, price, 3, 0, 0, 
                          Comment_Text + " #" + IntegerToString(tradeNumber), 
                          Magic_Number, 0, clrGreen);
   
   if(ticket > 0)
   {
      Print("Trade ", tradeNumber, " opened at ", price, " (Pyramiding)");
   }
   else
   {
      Print("Error opening pyramiding trade: ", GetLastError());
   }
}

//+------------------------------------------------------------------+
//| Check for FIFO entry (when max trades reached)                    |
//+------------------------------------------------------------------+
void CheckFIFOEntry()
{
   // Check on new bar only (on pyramiding timeframe)
   static datetime lastBarTime = 0;
   datetime currentBarTime = iTime(Symbol(), Pyramiding_Timeframe, 0);
   
   if(currentBarTime == lastBarTime)
      return;
   
   lastBarTime = currentBarTime;
   
   // Check if previous candle closed above last buy level on pyramiding timeframe
   double closePrice = iClose(Symbol(), Pyramiding_Timeframe, 1);
   
   if(closePrice <= g_lastBuyPrice)
      return;
   
   // Find oldest trade to close
   int oldestTicket = -1;
   
   for(int i = OrdersTotal() - 1; i >= 0; i--)
   {
      if(OrderSelect(i, SELECT_BY_POS, MODE_TRADES))
      {
         if(OrderSymbol() == Symbol() && OrderMagicNumber() == Magic_Number && OrderType() == OP_BUY)
         {
            if(oldestTicket == -1 || OrderTicket() < oldestTicket)
            {
               oldestTicket = OrderTicket();
            }
         }
      }
   }
   
   // Close oldest trade
   if(oldestTicket > 0 && OrderSelect(oldestTicket, SELECT_BY_TICKET))
   {
      bool closed = OrderClose(oldestTicket, OrderLots(), Bid, 3, clrRed);
      if(closed)
      {
         Print("FIFO: Closed oldest trade #", oldestTicket, " to maintain max ", Max_Trades, " trades");
      }
      else
      {
         Print("Error closing oldest trade: ", GetLastError());
         return;  // Don't open new trade if close failed
      }
   }
   
   // Open new trade
   double price = Ask;
   int ticket = OrderSend(Symbol(), OP_BUY, Lot_Size, price, 3, 0, 0, 
                          Comment_Text + " #FIFO", Magic_Number, 0, clrGreen);
   
   if(ticket > 0)
   {
      Print("FIFO: New trade opened at ", price);
   }
   else
   {
      Print("Error opening FIFO trade: ", GetLastError());
   }
}

//+------------------------------------------------------------------+
//| Check exit conditions                                             |
//+------------------------------------------------------------------+
void CheckExitConditions()
{
   // Calculate combined profit
   double combinedProfit = GetCombinedProfit();
   
   // Exit Condition 1: Profit Target Reached
   if(combinedProfit >= Profit_Target)
   {
      Print("Profit Target reached: $", combinedProfit, " >= $", Profit_Target);
      CloseAllTrades("Profit Target Reached");
      return;
   }
   
   // Exit Condition 2: Combined Profit Protection (only after 2+ trades)
   if(g_totalTrades >= 2 && combinedProfit < Combined_Profit_Protection)
   {
      Print("Combined Profit Protection triggered: $", combinedProfit, " < $", Combined_Profit_Protection);
      CloseAllTrades("Profit Protection");
      return;
   }
}

//+------------------------------------------------------------------+
//| Calculate combined profit of all open trades                      |
//+------------------------------------------------------------------+
double GetCombinedProfit()
{
   double totalProfit = 0;
   
   for(int i = OrdersTotal() - 1; i >= 0; i--)
   {
      if(OrderSelect(i, SELECT_BY_POS, MODE_TRADES))
      {
         if(OrderSymbol() == Symbol() && OrderMagicNumber() == Magic_Number && OrderType() == OP_BUY)
         {
            totalProfit += OrderProfit() + OrderSwap() + OrderCommission();
         }
      }
   }
   
   return totalProfit;
}

//+------------------------------------------------------------------+
//| Close all trades                                                  |
//+------------------------------------------------------------------+
void CloseAllTrades(string reason)
{
   Print("Closing all trades - Reason: ", reason);
   
   int closed = 0;
   
   for(int i = OrdersTotal() - 1; i >= 0; i--)
   {
      if(OrderSelect(i, SELECT_BY_POS, MODE_TRADES))
      {
         if(OrderSymbol() == Symbol() && OrderMagicNumber() == Magic_Number && OrderType() == OP_BUY)
         {
            bool success = OrderClose(OrderTicket(), OrderLots(), Bid, 3, clrRed);
            if(success)
            {
               closed++;
            }
            else
            {
               Print("Error closing ticket ", OrderTicket(), ": ", GetLastError());
            }
         }
      }
   }
   
   Print("Closed ", closed, " trades. Resetting strategy.");
   
   // Reset state
   g_rsiWasBelowOversold = false;
   g_firstTradeTicket = -1;
   g_firstTradePrice = 0;
   g_firstTradeSLRemoved = false;
   g_lastBuyPrice = 0;
}

//+------------------------------------------------------------------+
//| Update dashboard display                                          |
//+------------------------------------------------------------------+
void UpdateDashboard()
{
   double combinedProfit = GetCombinedProfit();
   double combinedSellProfit = GetCombinedSellProfit();
   double currentRSI = iRSI(Symbol(), RSI_Timeframe, RSI_Period, RSI_Price, 0);
   
   // Determine status text
   string status = "Watching for Entry";
   if(Trade_Direction == TRADE_BUY_ONLY && g_totalTrades > 0)
      status = "BUY Active (" + IntegerToString(g_totalTrades) + " trades)";
   else if(Trade_Direction == TRADE_SELL_ONLY && g_totalSellTrades > 0)
      status = "SELL Active (" + IntegerToString(g_totalSellTrades) + " trades)";
   else if(Trade_Direction == TRADE_BOTH && (g_totalTrades > 0 || g_totalSellTrades > 0))
      status = "BUY:" + IntegerToString(g_totalTrades) + " | SELL:" + IntegerToString(g_totalSellTrades);
   
   // Create/update text labels
   int yOffset = Dashboard_Y;
   
   CreateLabel(g_dashPrefix + "Title", "=== RSI PYRAMIDING EA ===", Dashboard_X, yOffset, clrWhite, 10);
   yOffset += 20;
   
   // Trade Direction Mode
   string directionText = (Trade_Direction == TRADE_BUY_ONLY) ? "BUY ONLY" : 
                          (Trade_Direction == TRADE_SELL_ONLY) ? "SELL ONLY" : "BOTH";
   color directionColor = (Trade_Direction == TRADE_BUY_ONLY) ? clrLime : 
                          (Trade_Direction == TRADE_SELL_ONLY) ? clrOrange : clrAqua;
   CreateLabel(g_dashPrefix + "Direction", "Mode: " + directionText, Dashboard_X, yOffset, directionColor, 9);
   yOffset += 18;
   
   CreateLabel(g_dashPrefix + "Status", "Status: " + status, Dashboard_X, yOffset, clrYellow, 9);
   yOffset += 18;
   
   // === BUY SIDE INFO ===
   if(Trade_Direction == TRADE_BUY_ONLY || Trade_Direction == TRADE_BOTH)
   {
      CreateLabel(g_dashPrefix + "BuyTrades", "BUY Trades: " + IntegerToString(g_totalTrades) + "/" + IntegerToString(Max_Trades), 
                  Dashboard_X, yOffset, clrLime, 9);
      yOffset += 18;
      
      color buyProfitColor = (combinedProfit >= 0) ? Profit_Color : Loss_Color;
      CreateLabel(g_dashPrefix + "BuyProfit", "BUY Profit: $" + DoubleToString(combinedProfit, 2), 
                  Dashboard_X, yOffset, buyProfitColor, 9);
      yOffset += 18;
      
      if(g_totalTrades >= 2)
      {
         double protectionLevel = Combined_Profit_Protection;
         color protColor = (combinedProfit >= protectionLevel) ? clrLime : clrOrange;
         CreateLabel(g_dashPrefix + "BuyProtection", "BUY Protection: $" + DoubleToString(combinedProfit, 2) + " / $" + DoubleToString(protectionLevel, 2), 
                     Dashboard_X, yOffset, protColor, 8);
         yOffset += 18;
      }
   }
   
   // === SELL SIDE INFO ===
   if(Trade_Direction == TRADE_SELL_ONLY || Trade_Direction == TRADE_BOTH)
   {
      CreateLabel(g_dashPrefix + "SellTrades", "SELL Trades: " + IntegerToString(g_totalSellTrades) + "/" + IntegerToString(Max_Trades), 
                  Dashboard_X, yOffset, clrOrange, 9);
      yOffset += 18;
      
      color sellProfitColor = (combinedSellProfit >= 0) ? Profit_Color : Loss_Color;
      CreateLabel(g_dashPrefix + "SellProfit", "SELL Profit: $" + DoubleToString(combinedSellProfit, 2), 
                  Dashboard_X, yOffset, sellProfitColor, 9);
      yOffset += 18;
      
      if(g_totalSellTrades >= 2)
      {
         double protectionLevel = Combined_Profit_Protection;
         color protColor = (combinedSellProfit >= protectionLevel) ? clrLime : clrOrange;
         CreateLabel(g_dashPrefix + "SellProtection", "SELL Protection: $" + DoubleToString(combinedSellProfit, 2) + " / $" + DoubleToString(protectionLevel, 2), 
                     Dashboard_X, yOffset, protColor, 8);
         yOffset += 18;
      }
   }
   
   // === COMBINED INFO (for BOTH mode) ===
   if(Trade_Direction == TRADE_BOTH)
   {
      double totalProfit = combinedProfit + combinedSellProfit;
      color totalProfitColor = (totalProfit >= 0) ? Profit_Color : Loss_Color;
      CreateLabel(g_dashPrefix + "TotalProfit", "TOTAL Profit: $" + DoubleToString(totalProfit, 2), 
                  Dashboard_X, yOffset, totalProfitColor, 10);
      yOffset += 18;
   }
   
   // === GENERAL INFO ===
   CreateLabel(g_dashPrefix + "Target", "Target: $" + DoubleToString(Profit_Target, 0), 
               Dashboard_X, yOffset, clrWhite, 8);
   yOffset += 18;
   
   string rsiTF = TimeframeToString(RSI_Timeframe);
   color rsiColor = clrWhite;
   if(currentRSI <= RSI_Oversold)
      rsiColor = clrLime;
   else if(currentRSI >= RSI_Overbought)
      rsiColor = clrOrange;
   CreateLabel(g_dashPrefix + "RSI", "RSI (" + rsiTF + "): " + DoubleToString(currentRSI, 1) + 
               " [" + IntegerToString(RSI_Oversold) + "/" + IntegerToString(RSI_Overbought) + "]", 
               Dashboard_X, yOffset, rsiColor, 8);
   yOffset += 18;
   
   // EMA Filter Status
   if(Use_EMA_Filter)
   {
      bool aboveEMA = IsMarketAboveEMA();
      bool belowEMA = IsMarketBelowEMA();
      string emaTF = TimeframeToString(EMA_Timeframe);
      string emaStatus = "";
      color emaColor = clrWhite;
      
      if(Trade_Direction == TRADE_BUY_ONLY)
      {
         emaStatus = aboveEMA ? "ABOVE (Active)" : "BELOW (No Entry)";
         emaColor = aboveEMA ? clrLime : clrRed;
      }
      else if(Trade_Direction == TRADE_SELL_ONLY)
      {
         emaStatus = belowEMA ? "BELOW (Active)" : "ABOVE (No Entry)";
         emaColor = belowEMA ? clrOrange : clrRed;
      }
      else  // BOTH
      {
         if(aboveEMA && belowEMA)
            emaStatus = "BOTH Active";
         else if(aboveEMA)
            emaStatus = "BUY Active";
         else if(belowEMA)
            emaStatus = "SELL Active";
         else
            emaStatus = "Neutral";
         emaColor = (aboveEMA || belowEMA) ? clrAqua : clrGray;
      }
      
      CreateLabel(g_dashPrefix + "EMA", "EMA (" + emaTF + "): " + emaStatus, 
                  Dashboard_X, yOffset, emaColor, 8);
      yOffset += 18;
   }
   
   // Pyramiding Timeframe Display
   string pyramidTF = TimeframeToString(Pyramiding_Timeframe);
   CreateLabel(g_dashPrefix + "PyramidTF", "Pyramid TF: " + pyramidTF, 
               Dashboard_X, yOffset, clrCyan, 8);
   yOffset += 25;
   
   // Close All Trades Button
   CreateButton(g_dashPrefix + "CloseAllBtn", "CLOSE ALL TRADES", Dashboard_X, yOffset, 180, 30, clrRed, clrWhite);
}

//+------------------------------------------------------------------+
//| Create or update text label                                       |
//+------------------------------------------------------------------+
void CreateLabel(string name, string text, int x, int y, color clr, int fontSize)
{
   if(ObjectFind(0, name) < 0)
   {
      ObjectCreate(0, name, OBJ_LABEL, 0, 0, 0);
      ObjectSetInteger(0, name, OBJPROP_CORNER, CORNER_LEFT_UPPER);
      ObjectSetInteger(0, name, OBJPROP_ANCHOR, ANCHOR_LEFT_UPPER);
      ObjectSetInteger(0, name, OBJPROP_XDISTANCE, x);
      ObjectSetInteger(0, name, OBJPROP_YDISTANCE, y);
      ObjectSetInteger(0, name, OBJPROP_FONTSIZE, fontSize);
      ObjectSetString(0, name, OBJPROP_FONT, "Consolas");
   }
   
   ObjectSetString(0, name, OBJPROP_TEXT, text);
   ObjectSetInteger(0, name, OBJPROP_COLOR, clr);
}

//+------------------------------------------------------------------+
//| Create or update button                                           |
//+------------------------------------------------------------------+
void CreateButton(string name, string text, int x, int y, int width, int height, color bgColor, color textColor)
{
   if(ObjectFind(0, name) < 0)
   {
      ObjectCreate(0, name, OBJ_BUTTON, 0, 0, 0);
      ObjectSetInteger(0, name, OBJPROP_CORNER, CORNER_LEFT_UPPER);
      ObjectSetInteger(0, name, OBJPROP_XDISTANCE, x);
      ObjectSetInteger(0, name, OBJPROP_YDISTANCE, y);
      ObjectSetInteger(0, name, OBJPROP_XSIZE, width);
      ObjectSetInteger(0, name, OBJPROP_YSIZE, height);
      ObjectSetInteger(0, name, OBJPROP_FONTSIZE, 9);
      ObjectSetString(0, name, OBJPROP_FONT, "Arial Bold");
   }
   
   ObjectSetString(0, name, OBJPROP_TEXT, text);
   ObjectSetInteger(0, name, OBJPROP_BGCOLOR, bgColor);
   ObjectSetInteger(0, name, OBJPROP_COLOR, textColor);
   ObjectSetInteger(0, name, OBJPROP_STATE, false);
}

//+------------------------------------------------------------------+
//| Remove dashboard objects                                          |
//+------------------------------------------------------------------+
void RemoveDashboard()
{
   ObjectDelete(0, g_dashPrefix + "Title");
   ObjectDelete(0, g_dashPrefix + "Direction");
   ObjectDelete(0, g_dashPrefix + "Status");
   ObjectDelete(0, g_dashPrefix + "Trades");  // Legacy - may not exist
   ObjectDelete(0, g_dashPrefix + "BuyTrades");
   ObjectDelete(0, g_dashPrefix + "BuyProfit");
   ObjectDelete(0, g_dashPrefix + "BuyProtection");
   ObjectDelete(0, g_dashPrefix + "SellTrades");
   ObjectDelete(0, g_dashPrefix + "SellProfit");
   ObjectDelete(0, g_dashPrefix + "SellProtection");
   ObjectDelete(0, g_dashPrefix + "TotalProfit");
   ObjectDelete(0, g_dashPrefix + "Profit");  // Legacy - may not exist
   ObjectDelete(0, g_dashPrefix + "Target");
   ObjectDelete(0, g_dashPrefix + "RSI");
   ObjectDelete(0, g_dashPrefix + "SL");  // Legacy - may not exist
   ObjectDelete(0, g_dashPrefix + "Protection");  // Legacy - may not exist
   ObjectDelete(0, g_dashPrefix + "EMA");
   ObjectDelete(0, g_dashPrefix + "PyramidTF");
   ObjectDelete(0, g_dashPrefix + "CloseAllBtn");
}

//+------------------------------------------------------------------+
//| Check if market is above EMA                                      |
//+------------------------------------------------------------------+
bool IsMarketAboveEMA()
{
   double currentPrice = iClose(Symbol(), EMA_Timeframe, 0);
   double emaValue = iMA(Symbol(), EMA_Timeframe, EMA_Period, 0, MODE_EMA, EMA_Price, 0);
   
   return (currentPrice > emaValue);
}

//+------------------------------------------------------------------+
//| Check if market is below EMA                                      |
//+------------------------------------------------------------------+
bool IsMarketBelowEMA()
{
   double currentPrice = iClose(Symbol(), EMA_Timeframe, 0);
   double emaValue = iMA(Symbol(), EMA_Timeframe, EMA_Period, 0, MODE_EMA, EMA_Price, 0);
   
   return (currentPrice < emaValue);
}

//+------------------------------------------------------------------+
//| Check for initial SELL entry signal (Trade 1)                     |
//+------------------------------------------------------------------+
void CheckInitialSellEntry()
{
   // Check EMA filter first (if enabled)
   if(Use_EMA_Filter)
   {
      if(!IsMarketBelowEMA())
      {
         // Market not below EMA - don't check signals
         g_rsiWasAboveOverbought = false;  // Reset RSI state
         return;
      }
   }
   
   // MODE 1: RSI-based entry (original strategy)
   if(Use_RSI_Filter)
   {
      // Calculate current RSI on selected timeframe
      double currentRSI = iRSI(Symbol(), RSI_Timeframe, RSI_Period, RSI_Price, 0);
      double previousRSI = iRSI(Symbol(), RSI_Timeframe, RSI_Period, RSI_Price, 1);
      
      // Check if RSI was above overbought
      if(previousRSI > RSI_Overbought)
      {
         g_rsiWasAboveOverbought = true;
      }
      
      // Entry signal: RSI crosses from above overbought to below
      if(g_rsiWasAboveOverbought && currentRSI < RSI_Overbought)
      {
         // Open first sell trade
         double price = Bid;
         double sl = price + First_Trade_SL;  // Price-based SL
         double tp = 0;  // No TP
         
         int ticket = OrderSend(Symbol(), OP_SELL, Lot_Size, price, 3, sl, tp, 
                                Comment_Text + " SELL #1", Magic_Number, 0, clrRed);
         
         if(ticket > 0)
         {
            Print("SELL Trade 1 opened at ", price, " with SL at ", sl, " (RSI Entry)");
            g_firstSellTradeTicket = ticket;
            g_firstSellTradePrice = price;
            g_lastSellPrice = price;
            g_rsiWasAboveOverbought = false;  // Reset flag
            g_firstSellTradeSLRemoved = false;
         }
         else
         {
            Print("Error opening SELL Trade 1: ", GetLastError());
         }
      }
      
      g_previousRSI = currentRSI;
   }
   // MODE 2: EMA-only entry (instant entry on new candle if below EMA)
   else
   {
      // Check on new bar only (on pyramiding timeframe)
      static datetime lastBarTime = 0;
      datetime currentBarTime = iTime(Symbol(), Pyramiding_Timeframe, 0);
      
      if(currentBarTime == lastBarTime)
         return;  // Wait for new bar
      
      lastBarTime = currentBarTime;
      
      // Verify EMA filter is enabled and market is below EMA
      if(!Use_EMA_Filter)
      {
         Print("Warning: RSI disabled but EMA filter also disabled. No entry conditions available.");
         return;
      }
      
      // Market is below EMA (already checked above), enter immediately
      double price = Bid;
      double sl = price + First_Trade_SL;  // Price-based SL
      double tp = 0;  // No TP
      
      int ticket = OrderSend(Symbol(), OP_SELL, Lot_Size, price, 3, sl, tp, 
                             Comment_Text + " SELL #1", Magic_Number, 0, clrRed);
      
      if(ticket > 0)
      {
         Print("SELL Trade 1 opened at ", price, " with SL at ", sl, " (EMA-Only Entry)");
         g_firstSellTradeTicket = ticket;
         g_firstSellTradePrice = price;
         g_lastSellPrice = price;
         g_firstSellTradeSLRemoved = false;
      }
      else
      {
         Print("Error opening SELL Trade 1: ", GetLastError());
      }
   }
}

//+------------------------------------------------------------------+
//| Check for second SELL trade entry                                 |
//+------------------------------------------------------------------+
void CheckSecondSellTradeEntry()
{
   // Verify first sell trade exists
   if(g_firstSellTradeTicket <= 0 || !OrderSelect(g_firstSellTradeTicket, SELECT_BY_TICKET))
      return;
   
   // Check conditions for Sell Trade 2:
   // 1. First sell trade profit >= $30
   double firstSellTradeProfit = OrderProfit() + OrderSwap() + OrderCommission();
   
   if(firstSellTradeProfit < Min_Profit_For_Second)
      return;
   
   // 2. Current candle closed below first sell trade entry (check on new bar on pyramiding timeframe)
   static datetime lastBarTime = 0;
   datetime currentBarTime = iTime(Symbol(), Pyramiding_Timeframe, 0);
   
   if(currentBarTime == lastBarTime)
      return;  // Wait for new bar
   
   lastBarTime = currentBarTime;
   
   double closePrice = iClose(Symbol(), Pyramiding_Timeframe, 1);  // Previous bar close on pyramiding TF
   
   if(closePrice >= g_firstSellTradePrice)
      return;
   
   // 3. RSI not oversold (optional check on RSI timeframe)
   double currentRSI = iRSI(Symbol(), RSI_Timeframe, RSI_Period, RSI_Price, 0);
   if(currentRSI < RSI_Oversold)
      return;
   
   // Open Sell Trade 2
   double price = Bid;
   int ticket = OrderSend(Symbol(), OP_SELL, Lot_Size, price, 3, 0, 0, 
                          Comment_Text + " SELL #2", Magic_Number, 0, clrOrange);
   
   if(ticket > 0)
   {
      Print("SELL Trade 2 opened at ", price);
      
      // Remove SL from Sell Trade 1
      if(OrderSelect(g_firstSellTradeTicket, SELECT_BY_TICKET))
      {
         bool modified = OrderModify(g_firstSellTradeTicket, OrderOpenPrice(), 0, 0, 0, clrNONE);
         if(modified)
         {
            Print("SELL Trade 1 Stop Loss removed");
            g_firstSellTradeSLRemoved = true;
         }
         else
         {
            Print("Error removing SELL Trade 1 SL: ", GetLastError());
         }
      }
   }
   else
   {
      Print("Error opening SELL Trade 2: ", GetLastError());
   }
}

//+------------------------------------------------------------------+
//| Check for SELL pyramiding entry (Trades 3-20)                     |
//+------------------------------------------------------------------+
void CheckSellPyramidingEntry()
{
   // Check on new bar only (on pyramiding timeframe)
   static datetime lastBarTime = 0;
   datetime currentBarTime = iTime(Symbol(), Pyramiding_Timeframe, 0);
   
   if(currentBarTime == lastBarTime)
      return;
   
   lastBarTime = currentBarTime;
   
   // Check if previous candle closed below last sell level on pyramiding timeframe
   double closePrice = iClose(Symbol(), Pyramiding_Timeframe, 1);
   
   if(closePrice >= g_lastSellPrice)
      return;
   
   // Open new sell pyramiding trade
   double price = Bid;
   int tradeNumber = g_totalSellTrades + 1;
   
   int ticket = OrderSend(Symbol(), OP_SELL, Lot_Size, price, 3, 0, 0, 
                          Comment_Text + " SELL #" + IntegerToString(tradeNumber), 
                          Magic_Number, 0, clrOrange);
   
   if(ticket > 0)
   {
      Print("SELL Trade ", tradeNumber, " opened at ", price, " (Pyramiding)");
   }
   else
   {
      Print("Error opening sell pyramiding trade: ", GetLastError());
   }
}

//+------------------------------------------------------------------+
//| Check for SELL FIFO entry (when max trades reached)               |
//+------------------------------------------------------------------+
void CheckSellFIFOEntry()
{
   // Check on new bar only (on pyramiding timeframe)
   static datetime lastBarTime = 0;
   datetime currentBarTime = iTime(Symbol(), Pyramiding_Timeframe, 0);
   
   if(currentBarTime == lastBarTime)
      return;
   
   lastBarTime = currentBarTime;
   
   // Check if previous candle closed below last sell level on pyramiding timeframe
   double closePrice = iClose(Symbol(), Pyramiding_Timeframe, 1);
   
   if(closePrice >= g_lastSellPrice)
      return;
   
   // Find oldest sell trade to close
   int oldestTicket = -1;
   
   for(int i = OrdersTotal() - 1; i >= 0; i--)
   {
      if(OrderSelect(i, SELECT_BY_POS, MODE_TRADES))
      {
         if(OrderSymbol() == Symbol() && OrderMagicNumber() == Magic_Number && OrderType() == OP_SELL)
         {
            if(oldestTicket == -1 || OrderTicket() < oldestTicket)
            {
               oldestTicket = OrderTicket();
            }
         }
      }
   }
   
   // Close oldest sell trade
   if(oldestTicket > 0 && OrderSelect(oldestTicket, SELECT_BY_TICKET))
   {
      bool closed = OrderClose(oldestTicket, OrderLots(), Ask, 3, clrRed);
      if(closed)
      {
         Print("FIFO: Closed oldest SELL trade #", oldestTicket, " to maintain max ", Max_Trades, " trades");
      }
      else
      {
         Print("Error closing oldest sell trade: ", GetLastError());
         return;  // Don't open new trade if close failed
      }
   }
   
   // Open new sell trade
   double price = Bid;
   int ticket = OrderSend(Symbol(), OP_SELL, Lot_Size, price, 3, 0, 0, 
                          Comment_Text + " SELL #FIFO", Magic_Number, 0, clrOrange);
   
   if(ticket > 0)
   {
      Print("FIFO: New SELL trade opened at ", price);
   }
   else
   {
      Print("Error opening SELL FIFO trade: ", GetLastError());
   }
}

//+------------------------------------------------------------------+
//| Check SELL exit conditions                                        |
//+------------------------------------------------------------------+
void CheckSellExitConditions()
{
   // Calculate combined sell profit
   double combinedProfit = GetCombinedSellProfit();
   
   // Exit Condition 1: Profit Target Reached
   if(combinedProfit >= Profit_Target)
   {
      Print("SELL Profit Target reached: $", combinedProfit, " >= $", Profit_Target);
      CloseAllSellTrades("Profit Target Reached");
      return;
   }
   
   // Exit Condition 2: Combined Profit Protection (only after 2+ trades)
   if(g_totalSellTrades >= 2 && combinedProfit < Combined_Profit_Protection)
   {
      Print("SELL Combined Profit Protection triggered: $", combinedProfit, " < $", Combined_Profit_Protection);
      CloseAllSellTrades("Profit Protection");
      return;
   }
}

//+------------------------------------------------------------------+
//| Calculate combined profit of all open SELL trades                 |
//+------------------------------------------------------------------+
double GetCombinedSellProfit()
{
   double totalProfit = 0;
   
   for(int i = OrdersTotal() - 1; i >= 0; i--)
   {
      if(OrderSelect(i, SELECT_BY_POS, MODE_TRADES))
      {
         if(OrderSymbol() == Symbol() && OrderMagicNumber() == Magic_Number && OrderType() == OP_SELL)
         {
            totalProfit += OrderProfit() + OrderSwap() + OrderCommission();
         }
      }
   }
   
   return totalProfit;
}

//+------------------------------------------------------------------+
//| Close all SELL trades                                             |
//+------------------------------------------------------------------+
void CloseAllSellTrades(string reason)
{
   Print("Closing all SELL trades - Reason: ", reason);
   
   int closed = 0;
   
   for(int i = OrdersTotal() - 1; i >= 0; i--)
   {
      if(OrderSelect(i, SELECT_BY_POS, MODE_TRADES))
      {
         if(OrderSymbol() == Symbol() && OrderMagicNumber() == Magic_Number && OrderType() == OP_SELL)
         {
            bool success = OrderClose(OrderTicket(), OrderLots(), Ask, 3, clrRed);
            if(success)
            {
               closed++;
            }
            else
            {
               Print("Error closing SELL ticket ", OrderTicket(), ": ", GetLastError());
            }
         }
      }
   }
   
   Print("Closed ", closed, " SELL trades. Resetting sell strategy.");
   
   // Reset sell state
   g_rsiWasAboveOverbought = false;
   g_firstSellTradeTicket = -1;
   g_firstSellTradePrice = 0;
   g_firstSellTradeSLRemoved = false;
   g_lastSellPrice = 0;
}


//+------------------------------------------------------------------+
//| Close all trades manually (button click)                          |
//+------------------------------------------------------------------+
void CloseAllTradesManual()
{
   Print("=== MANUAL CLOSE ALL TRADES TRIGGERED ===");
   
   int totalClosed = 0;
   int buyTradesClosed = 0;
   int sellTradesClosed = 0;
   
   // Close all BUY trades
   for(int i = OrdersTotal() - 1; i >= 0; i--)
   {
      if(OrderSelect(i, SELECT_BY_POS, MODE_TRADES))
      {
         if(OrderSymbol() == Symbol() && OrderMagicNumber() == Magic_Number && OrderType() == OP_BUY)
         {
            bool success = OrderClose(OrderTicket(), OrderLots(), Bid, 3, clrRed);
            if(success)
            {
               buyTradesClosed++;
               totalClosed++;
            }
            else
            {
               Print("Error closing BUY ticket ", OrderTicket(), ": ", GetLastError());
            }
         }
      }
   }
   
   // Close all SELL trades
   for(int i = OrdersTotal() - 1; i >= 0; i--)
   {
      if(OrderSelect(i, SELECT_BY_POS, MODE_TRADES))
      {
         if(OrderSymbol() == Symbol() && OrderMagicNumber() == Magic_Number && OrderType() == OP_SELL)
         {
            bool success = OrderClose(OrderTicket(), OrderLots(), Ask, 3, clrRed);
            if(success)
            {
               sellTradesClosed++;
               totalClosed++;
            }
            else
            {
               Print("Error closing SELL ticket ", OrderTicket(), ": ", GetLastError());
            }
         }
      }
   }
   
   Print("Manual Close Complete: BUY trades closed: ", buyTradesClosed, ", SELL trades closed: ", sellTradesClosed, ", Total: ", totalClosed);
   
   // Reset all state variables for fresh start
   // Buy state
   g_rsiWasBelowOversold = false;
   g_firstTradeTicket = -1;
   g_firstTradePrice = 0;
   g_firstTradeSLRemoved = false;
   g_lastBuyPrice = 0;
   g_totalTrades = 0;
   
   // Sell state
   g_rsiWasAboveOverbought = false;
   g_firstSellTradeTicket = -1;
   g_firstSellTradePrice = 0;
   g_firstSellTradeSLRemoved = false;
   g_lastSellPrice = 0;
   g_totalSellTrades = 0;
   
   Print("EA Reset Complete. Ready to start fresh.");
   Alert("All trades closed. Total: ", totalClosed, " (BUY: ", buyTradesClosed, ", SELL: ", sellTradesClosed, "). EA reset and ready to start fresh.");
}

//+------------------------------------------------------------------+
//| Convert timeframe enum to string                                  |
//+------------------------------------------------------------------+
string TimeframeToString(ENUM_TIMEFRAMES tf)
{
   switch(tf)
   {
      case PERIOD_CURRENT: return "CURRENT";
      case PERIOD_M1:  return "M1";
      case PERIOD_M5:  return "M5";
      case PERIOD_M15: return "M15";
      case PERIOD_M30: return "M30";
      case PERIOD_H1:  return "H1";
      case PERIOD_H4:  return "H4";
      case PERIOD_D1:  return "D1";
      case PERIOD_W1:  return "W1";
      case PERIOD_MN1: return "MN1";
      default: return "UNKNOWN";
   }
}

//+------------------------------------------------------------------+
