//+------------------------------------------------------------------+
//|                                              PairTradingEA.mq5   |
//|                                                                  |
//+------------------------------------------------------------------+
#property copyright "Pair Trading EA"
#property version   "1.00"
#property strict

#include <Trade\Trade.mqh>

// Input parameters
input int MagicNumber = 123456;

// Global variables
CTrade trade;
string symbol1 = "";
string symbol2 = "";
double lotSize1 = 0.01;
double lotSize2 = 0.01;
int strategySelected = 0; // 0=none, 1-4=strategies
double takeProfitAmount = 100.0;
double lossAmountToRepeat = 50.0;
bool tradesActive = false;
double initialLossLevel = 0;
int tradeSetCount = 0;
double lastLossLevel = 0;

// UI object names
string btnPlaceTrades = "btnPlace";
string btnCloseAll = "btnClose";
string lblPnL = "lblPnL";
string editSymbol1 = "editSym1";
string editSymbol2 = "editSym2";
string editLot1 = "editLot1";
string editLot2 = "editLot2";
string editTP = "editTP";
string editLoss = "editLoss";
string radioStrat1 = "radioStrat1";
string radioStrat2 = "radioStrat2";
string radioStrat3 = "radioStrat3";
string radioStrat4 = "radioStrat4";

//+------------------------------------------------------------------+
//| Expert initialization function                                     |
//+------------------------------------------------------------------+
int OnInit()
{
   trade.SetExpertMagicNumber(MagicNumber);
   CreateDashboard();
   EventSetTimer(1);
   return(INIT_SUCCEEDED);
}

//+------------------------------------------------------------------+
//| Expert deinitialization function                                   |
//+------------------------------------------------------------------+
void OnDeinit(const int reason)
{
   EventKillTimer();
   DeleteAllObjects();
}

//+------------------------------------------------------------------+
//| Expert tick function                                               |
//+------------------------------------------------------------------+
void OnTick()
{
   if(tradesActive)
   {
      double currentPnL = GetTotalPnL();
      UpdatePnLDisplay(currentPnL);
      
      // Check for take profit
      if(currentPnL >= takeProfitAmount)
      {
         CloseAllTrades();
         tradesActive = false;
         tradeSetCount = 0;
         lastLossLevel = 0;
         Alert("Take Profit achieved! All positions closed. Profit: $", DoubleToString(currentPnL, 2));
      }
      
      // Check for loss threshold to place new trades
      // Formula: (2^n - 1) * lossAmountToRepeat where n is the next trade set number
      // Example with lossAmountToRepeat = 100:
      //   2nd set at: (2^1 - 1) * 100 = 100
      //   3rd set at: (2^2 - 1) * 100 = 300
      //   4th set at: (2^3 - 1) * 100 = 700
      //   5th set at: (2^4 - 1) * 100 = 1500
      if(lossAmountToRepeat > 0 && currentPnL < 0)
      {
         double absLoss = MathAbs(currentPnL);
         int nextTradeSetNumber = tradeSetCount + 1;
         double nextLossLevel = lossAmountToRepeat * (MathPow(2, nextTradeSetNumber) - 1);
         
         if(absLoss >= nextLossLevel && absLoss > lastLossLevel)
         {
            lastLossLevel = nextLossLevel;
            tradeSetCount++;
            PlaceTradeSet();
            Alert("Loss level reached: $", DoubleToString(currentPnL, 2), " (threshold: $", DoubleToString(nextLossLevel, 2), "). Placing additional trade set #", tradeSetCount + 1);
         }
      }
   }
}

//+------------------------------------------------------------------+
//| Timer function                                                     |
//+------------------------------------------------------------------+
void OnTimer()
{
   if(tradesActive)
   {
      double currentPnL = GetTotalPnL();
      UpdatePnLDisplay(currentPnL);
   }
}

//+------------------------------------------------------------------+
//| ChartEvent function                                                |
//+------------------------------------------------------------------+
void OnChartEvent(const int id, const long &lparam, const double &dparam, const string &sparam)
{
   if(id == CHARTEVENT_OBJECT_CLICK)
   {
      // Place Trades button
      if(sparam == btnPlaceTrades)
      {
         ObjectSetInteger(0, btnPlaceTrades, OBJPROP_STATE, false);
         OnPlaceTradesClick();
      }
      
      // Close All button
      if(sparam == btnCloseAll)
      {
         ObjectSetInteger(0, btnCloseAll, OBJPROP_STATE, false);
         OnCloseAllClick();
      }
      
      // Strategy selection
      if(sparam == radioStrat1)
      {
         SelectStrategy(1);
      }
      if(sparam == radioStrat2)
      {
         SelectStrategy(2);
      }
      if(sparam == radioStrat3)
      {
         SelectStrategy(3);
      }
      if(sparam == radioStrat4)
      {
         SelectStrategy(4);
      }
   }
}

//+------------------------------------------------------------------+
//| Create Dashboard UI                                               |
//+------------------------------------------------------------------+
void CreateDashboard()
{
   int x = 15;
   int y = 20;
   int panelWidth = 330;
   int panelHeight = 560;
   
   // Create background panel
   CreatePanel("panelBG", x - 5, y - 5, panelWidth, panelHeight, clrDarkSlateGray, clrWhite);
   
   // Title
   CreateLabel("lblTitle", x + 10, y + 5, "PAIR TRADING DASHBOARD", clrYellow, 11, true);
   y += 35;
   
   // Symbol 1
   CreateLabel("lblSym1", x + 5, y, "Symbol 1:", clrWhite, 9);
   CreateEditBox(editSymbol1, x + 80, y - 3, 230, 22, "EURUSD");
   y += 32;
   
   // Symbol 2
   CreateLabel("lblSym2", x + 5, y, "Symbol 2:", clrWhite, 9);
   CreateEditBox(editSymbol2, x + 80, y - 3, 230, 22, "GBPUSD");
   y += 32;
   
   // Lot Size 1
   CreateLabel("lblLot1", x + 5, y, "Lot Size 1:", clrWhite, 9);
   CreateEditBox(editLot1, x + 80, y - 3, 230, 22, "0.01");
   y += 32;
   
   // Lot Size 2
   CreateLabel("lblLot2", x + 5, y, "Lot Size 2:", clrWhite, 9);
   CreateEditBox(editLot2, x + 80, y - 3, 230, 22, "0.01");
   y += 40;
   
   // Strategy options
   CreateLabel("lblStrategy", x + 5, y, "Select Strategy:", clrWhite, 10, true);
   y += 28;
   
   CreateRadioButton(radioStrat1, x + 10, y, "Buy Symbol 1, Sell Symbol 2");
   y += 28;
   CreateRadioButton(radioStrat2, x + 10, y, "Sell Symbol 1, Buy Symbol 2");
   y += 28;
   CreateRadioButton(radioStrat3, x + 10, y, "Buy Symbol 1, Buy Symbol 2");
   y += 28;
   CreateRadioButton(radioStrat4, x + 10, y, "Sell Symbol 1, Sell Symbol 2");
   y += 38;
   
   // Take Profit
   CreateLabel("lblTP", x + 5, y, "Take Profit ($):", clrWhite, 9);
   CreateEditBox(editTP, x + 120, y - 3, 190, 22, "100");
   y += 32;
   
   // Loss Amount to Repeat
   CreateLabel("lblLoss", x + 5, y, "Loss to Repeat ($):", clrWhite, 9);
   CreateEditBox(editLoss, x + 120, y - 3, 190, 22, "50");
   y += 42;
   
   // Place Trades button
   CreateButton(btnPlaceTrades, x + 20, y, 280, 38, "PLACE TRADES", clrWhite, clrGreen);
   y += 50;
   
   // P/L Display
   CreateLabel(lblPnL, x + 60, y, "Current P/L: $0.00", clrYellow, 12, true);
   y += 38;
   
   // Close All button
   CreateButton(btnCloseAll, x + 20, y, 280, 38, "CLOSE ALL TRADES", clrWhite, clrRed);
   
   ChartRedraw();
}

//+------------------------------------------------------------------+
//| Create Panel                                                       |
//+------------------------------------------------------------------+
void CreatePanel(string name, int x, int y, int width, int height, color bgColor, color borderColor)
{
   ObjectCreate(0, name, OBJ_RECTANGLE_LABEL, 0, 0, 0);
   ObjectSetInteger(0, name, OBJPROP_XDISTANCE, x);
   ObjectSetInteger(0, name, OBJPROP_YDISTANCE, y);
   ObjectSetInteger(0, name, OBJPROP_XSIZE, width);
   ObjectSetInteger(0, name, OBJPROP_YSIZE, height);
   ObjectSetInteger(0, name, OBJPROP_BGCOLOR, bgColor);
   ObjectSetInteger(0, name, OBJPROP_BORDER_TYPE, BORDER_FLAT);
   ObjectSetInteger(0, name, OBJPROP_CORNER, CORNER_LEFT_UPPER);
   ObjectSetInteger(0, name, OBJPROP_COLOR, borderColor);
   ObjectSetInteger(0, name, OBJPROP_STYLE, STYLE_SOLID);
   ObjectSetInteger(0, name, OBJPROP_WIDTH, 1);
   ObjectSetInteger(0, name, OBJPROP_BACK, true);
}

//+------------------------------------------------------------------+
//| Create Label                                                       |
//+------------------------------------------------------------------+
void CreateLabel(string name, int x, int y, string text, color clr, int fontSize = 9, bool bold = false)
{
   ObjectCreate(0, name, OBJ_LABEL, 0, 0, 0);
   ObjectSetInteger(0, name, OBJPROP_XDISTANCE, x);
   ObjectSetInteger(0, name, OBJPROP_YDISTANCE, y);
   ObjectSetInteger(0, name, OBJPROP_COLOR, clr);
   ObjectSetInteger(0, name, OBJPROP_FONTSIZE, fontSize);
   ObjectSetString(0, name, OBJPROP_TEXT, text);
   ObjectSetString(0, name, OBJPROP_FONT, bold ? "Arial Bold" : "Arial");
   ObjectSetInteger(0, name, OBJPROP_CORNER, CORNER_LEFT_UPPER);
}

//+------------------------------------------------------------------+
//| Create Edit Box                                                    |
//+------------------------------------------------------------------+
void CreateEditBox(string name, int x, int y, int width, int height, string text)
{
   ObjectCreate(0, name, OBJ_EDIT, 0, 0, 0);
   ObjectSetInteger(0, name, OBJPROP_XDISTANCE, x);
   ObjectSetInteger(0, name, OBJPROP_YDISTANCE, y);
   ObjectSetInteger(0, name, OBJPROP_XSIZE, width);
   ObjectSetInteger(0, name, OBJPROP_YSIZE, height);
   ObjectSetString(0, name, OBJPROP_TEXT, text);
   ObjectSetInteger(0, name, OBJPROP_CORNER, CORNER_LEFT_UPPER);
   ObjectSetInteger(0, name, OBJPROP_COLOR, clrBlack);
   ObjectSetInteger(0, name, OBJPROP_BGCOLOR, clrWhite);
   ObjectSetInteger(0, name, OBJPROP_BORDER_COLOR, clrGray);
   ObjectSetInteger(0, name, OBJPROP_FONTSIZE, 9);
}

//+------------------------------------------------------------------+
//| Create Button                                                      |
//+------------------------------------------------------------------+
void CreateButton(string name, int x, int y, int width, int height, string text, color textColor, color bgColor)
{
   ObjectCreate(0, name, OBJ_BUTTON, 0, 0, 0);
   ObjectSetInteger(0, name, OBJPROP_XDISTANCE, x);
   ObjectSetInteger(0, name, OBJPROP_YDISTANCE, y);
   ObjectSetInteger(0, name, OBJPROP_XSIZE, width);
   ObjectSetInteger(0, name, OBJPROP_YSIZE, height);
   ObjectSetString(0, name, OBJPROP_TEXT, text);
   ObjectSetInteger(0, name, OBJPROP_COLOR, textColor);
   ObjectSetInteger(0, name, OBJPROP_BGCOLOR, bgColor);
   ObjectSetInteger(0, name, OBJPROP_CORNER, CORNER_LEFT_UPPER);
   ObjectSetInteger(0, name, OBJPROP_FONTSIZE, 10);
}

//+------------------------------------------------------------------+
//| Create Radio Button                                                |
//+------------------------------------------------------------------+
void CreateRadioButton(string name, int x, int y, string text)
{
   ObjectCreate(0, name, OBJ_BUTTON, 0, 0, 0);
   ObjectSetInteger(0, name, OBJPROP_XDISTANCE, x);
   ObjectSetInteger(0, name, OBJPROP_YDISTANCE, y);
   ObjectSetInteger(0, name, OBJPROP_XSIZE, 300);
   ObjectSetInteger(0, name, OBJPROP_YSIZE, 25);
   ObjectSetString(0, name, OBJPROP_TEXT, text);
   ObjectSetInteger(0, name, OBJPROP_COLOR, clrBlack);
   ObjectSetInteger(0, name, OBJPROP_BGCOLOR, clrLightGray);
   ObjectSetInteger(0, name, OBJPROP_CORNER, CORNER_LEFT_UPPER);
   ObjectSetInteger(0, name, OBJPROP_FONTSIZE, 9);
}

//+------------------------------------------------------------------+
//| Select Strategy                                                    |
//+------------------------------------------------------------------+
void SelectStrategy(int strategy)
{
   strategySelected = strategy;
   
   // Reset all radio buttons
   ObjectSetInteger(0, radioStrat1, OBJPROP_BGCOLOR, clrLightGray);
   ObjectSetInteger(0, radioStrat2, OBJPROP_BGCOLOR, clrLightGray);
   ObjectSetInteger(0, radioStrat3, OBJPROP_BGCOLOR, clrLightGray);
   ObjectSetInteger(0, radioStrat4, OBJPROP_BGCOLOR, clrLightGray);
   
   // Highlight selected
   string selectedRadio = "";
   if(strategy == 1) selectedRadio = radioStrat1;
   else if(strategy == 2) selectedRadio = radioStrat2;
   else if(strategy == 3) selectedRadio = radioStrat3;
   else if(strategy == 4) selectedRadio = radioStrat4;
   
   ObjectSetInteger(0, selectedRadio, OBJPROP_BGCOLOR, clrLightBlue);
   ObjectSetInteger(0, selectedRadio, OBJPROP_STATE, false);
   ChartRedraw();
}

//+------------------------------------------------------------------+
//| Place Trades Click Handler                                         |
//+------------------------------------------------------------------+
void OnPlaceTradesClick()
{
   // Read inputs
   symbol1 = ObjectGetString(0, editSymbol1, OBJPROP_TEXT);
   symbol2 = ObjectGetString(0, editSymbol2, OBJPROP_TEXT);
   lotSize1 = StringToDouble(ObjectGetString(0, editLot1, OBJPROP_TEXT));
   lotSize2 = StringToDouble(ObjectGetString(0, editLot2, OBJPROP_TEXT));
   takeProfitAmount = StringToDouble(ObjectGetString(0, editTP, OBJPROP_TEXT));
   lossAmountToRepeat = StringToDouble(ObjectGetString(0, editLoss, OBJPROP_TEXT));
   
   // Validate inputs
   if(symbol1 == "" || symbol2 == "")
   {
      Alert("Please enter both symbols!");
      return;
   }
   
   if(lotSize1 <= 0 || lotSize2 <= 0)
   {
      Alert("Lot sizes must be greater than 0!");
      return;
   }
   
   if(strategySelected == 0)
   {
      Alert("Please select a strategy!");
      return;
   }
   
   if(takeProfitAmount <= 0)
   {
      Alert("Take Profit must be greater than 0!");
      return;
   }
   
   // Close any existing trades first
   if(tradesActive)
   {
      CloseAllTrades();
   }
   
   // Reset counters
   tradeSetCount = 0;
   lastLossLevel = 0;
   
   // Place initial trade set
   if(PlaceTradeSet())
   {
      tradesActive = true;
      Alert("Initial trades placed successfully!");
   }
   else
   {
      Alert("Failed to place trades. Check your inputs and try again.");
   }
}

//+------------------------------------------------------------------+
//| Close All Click Handler                                           |
//+------------------------------------------------------------------+
void OnCloseAllClick()
{
   if(tradesActive)
   {
      double finalPnL = GetTotalPnL();
      CloseAllTrades();
      tradesActive = false;
      tradeSetCount = 0;
      lastLossLevel = 0;
      Alert("All trades closed manually. Final P/L: $", DoubleToString(finalPnL, 2));
   }
   else
   {
      Alert("No active trades to close.");
   }
}

//+------------------------------------------------------------------+
//| Place Trade Set                                                    |
//+------------------------------------------------------------------+
bool PlaceTradeSet()
{
   bool success = true;
   
   // Strategy 1: Buy Symbol 1, Sell Symbol 2
   if(strategySelected == 1)
   {
      success = success && trade.Buy(lotSize1, symbol1, 0, 0, 0, "PairTrade");
      success = success && trade.Sell(lotSize2, symbol2, 0, 0, 0, "PairTrade");
   }
   // Strategy 2: Sell Symbol 1, Buy Symbol 2
   else if(strategySelected == 2)
   {
      success = success && trade.Sell(lotSize1, symbol1, 0, 0, 0, "PairTrade");
      success = success && trade.Buy(lotSize2, symbol2, 0, 0, 0, "PairTrade");
   }
   // Strategy 3: Buy Symbol 1, Buy Symbol 2
   else if(strategySelected == 3)
   {
      success = success && trade.Buy(lotSize1, symbol1, 0, 0, 0, "PairTrade");
      success = success && trade.Buy(lotSize2, symbol2, 0, 0, 0, "PairTrade");
   }
   // Strategy 4: Sell Symbol 1, Sell Symbol 2
   else if(strategySelected == 4)
   {
      success = success && trade.Sell(lotSize1, symbol1, 0, 0, 0, "PairTrade");
      success = success && trade.Sell(lotSize2, symbol2, 0, 0, 0, "PairTrade");
   }
   
   return success;
}

//+------------------------------------------------------------------+
//| Close All Trades                                                   |
//+------------------------------------------------------------------+
void CloseAllTrades()
{
   for(int i = PositionsTotal() - 1; i >= 0; i--)
   {
      ulong ticket = PositionGetTicket(i);
      if(ticket > 0)
      {
         if(PositionSelectByTicket(ticket))
         {
            if(PositionGetInteger(POSITION_MAGIC) == MagicNumber)
            {
               trade.PositionClose(ticket);
            }
         }
      }
   }
   
   UpdatePnLDisplay(0);
}

//+------------------------------------------------------------------+
//| Get Total P/L                                                      |
//+------------------------------------------------------------------+
double GetTotalPnL()
{
   double totalPnL = 0;
   
   for(int i = 0; i < PositionsTotal(); i++)
   {
      ulong ticket = PositionGetTicket(i);
      if(ticket > 0)
      {
         if(PositionSelectByTicket(ticket))
         {
            if(PositionGetInteger(POSITION_MAGIC) == MagicNumber)
            {
               totalPnL += PositionGetDouble(POSITION_PROFIT) + PositionGetDouble(POSITION_SWAP);
            }
         }
      }
   }
   
   return totalPnL;
}

//+------------------------------------------------------------------+
//| Update P/L Display                                                 |
//+------------------------------------------------------------------+
void UpdatePnLDisplay(double pnl)
{
   string pnlText = "Current P/L: $" + DoubleToString(pnl, 2);
   color pnlColor = clrYellow;
   
   if(pnl > 0) pnlColor = clrLime;
   else if(pnl < 0) pnlColor = clrRed;
   
   ObjectSetString(0, lblPnL, OBJPROP_TEXT, pnlText);
   ObjectSetInteger(0, lblPnL, OBJPROP_COLOR, pnlColor);
   ChartRedraw();
}

//+------------------------------------------------------------------+
//| Delete All Dashboard Objects                                       |
//+------------------------------------------------------------------+
void DeleteAllObjects()
{
   ObjectDelete(0, "panelBG");
   ObjectDelete(0, "lblTitle");
   ObjectDelete(0, "lblSym1");
   ObjectDelete(0, "lblSym2");
   ObjectDelete(0, "lblLot1");
   ObjectDelete(0, "lblLot2");
   ObjectDelete(0, "lblStrategy");
   ObjectDelete(0, "lblTP");
   ObjectDelete(0, "lblLoss");
   ObjectDelete(0, lblPnL);
   ObjectDelete(0, editSymbol1);
   ObjectDelete(0, editSymbol2);
   ObjectDelete(0, editLot1);
   ObjectDelete(0, editLot2);
   ObjectDelete(0, editTP);
   ObjectDelete(0, editLoss);
   ObjectDelete(0, btnPlaceTrades);
   ObjectDelete(0, btnCloseAll);
   ObjectDelete(0, radioStrat1);
   ObjectDelete(0, radioStrat2);
   ObjectDelete(0, radioStrat3);
   ObjectDelete(0, radioStrat4);
   ChartRedraw();
}
//+------------------------------------------------------------------+

