//+------------------------------------------------------------------+
//|                                       RSI_Daily_Strategy_EA.mq5  |
//|                              Daily RSI Trading Strategy Expert   |
//|                                                                  |
//| Strategy: RSI Cross-based Entry/Exit on Daily Timeframe          |
//| - Long Entry: RSI crosses above oversold threshold (30)          |
//| - Long Exit: RSI exceeds exit threshold (60)                     |
//| - Short Entry: RSI crosses below overbought threshold (70)       |
//| - Short Exit: RSI drops below exit threshold (40)                |
//| Supports multiple concurrent positions per direction             |
//| Multi-Symbol Backtesting Support                                 |
//| Multi-Timeframe Trend Continuation Score Filter                  |
//+------------------------------------------------------------------+
#property copyright "FxLab"
#property version   "3.00"
#property strict

#include <Trade\Trade.mqh>

//+------------------------------------------------------------------+
//| Forex Pairs List - Major and Minor Pairs                         |
//+------------------------------------------------------------------+
#define MAX_SYMBOLS 28

string g_symbolList[MAX_SYMBOLS] = {
   // Major Pairs (7)
   "EURUSD",    // Euro / US Dollar
   "GBPUSD",    // British Pound / US Dollar
   "USDJPY",    // US Dollar / Japanese Yen
   "USDCHF",    // US Dollar / Swiss Franc
   "AUDUSD",    // Australian Dollar / US Dollar
   "USDCAD",    // US Dollar / Canadian Dollar
   "NZDUSD",    // New Zealand Dollar / US Dollar
   
   // Minor/Cross Pairs - EUR Crosses (6)
   "EURGBP",    // Euro / British Pound
   "EURJPY",    // Euro / Japanese Yen
   "EURCHF",    // Euro / Swiss Franc
   "EURAUD",    // Euro / Australian Dollar
   "EURCAD",    // Euro / Canadian Dollar
   "EURNZD",    // Euro / New Zealand Dollar
   
   // Minor/Cross Pairs - GBP Crosses (5)
   "GBPJPY",    // British Pound / Japanese Yen
   "GBPCHF",    // British Pound / Swiss Franc
   "GBPAUD",    // British Pound / Australian Dollar
   "GBPCAD",    // British Pound / Canadian Dollar
   "GBPNZD",    // British Pound / New Zealand Dollar
   
   // Minor/Cross Pairs - AUD Crosses (4)
   "AUDJPY",    // Australian Dollar / Japanese Yen
   "AUDCHF",    // Australian Dollar / Swiss Franc
   "AUDCAD",    // Australian Dollar / Canadian Dollar
   "AUDNZD",    // Australian Dollar / New Zealand Dollar
   
   // Minor/Cross Pairs - Other Crosses (6)
   "NZDJPY",    // New Zealand Dollar / Japanese Yen
   "NZDCHF",    // New Zealand Dollar / Swiss Franc
   "NZDCAD",    // New Zealand Dollar / Canadian Dollar
   "CADJPY",    // Canadian Dollar / Japanese Yen
   "CADCHF",    // Canadian Dollar / Swiss Franc
   "CHFJPY"     // Swiss Franc / Japanese Yen
};

//+------------------------------------------------------------------+
//| Input Parameters - All Strategy Options                          |
//+------------------------------------------------------------------+

//--- Multi-Symbol Mode
input group "=== Multi-Symbol Settings ==="
enum ENUM_SYMBOL_MODE
{
   SYMBOL_CURRENT_ONLY = 0,  // Current Chart Symbol Only
   SYMBOL_MULTI_ALL    = 1,  // All Major & Minor Pairs
   SYMBOL_MAJORS_ONLY  = 2,  // Major Pairs Only (7 pairs)
   SYMBOL_MINORS_ONLY  = 3,  // Minor/Cross Pairs Only (21 pairs)
   SYMBOL_CUSTOM_LIST  = 4   // Custom Symbol List
};
input ENUM_SYMBOL_MODE     InpSymbolMode        = SYMBOL_CURRENT_ONLY; // Symbol Trading Mode
input string               InpCustomSymbols     = "EURUSD,GBPUSD,USDJPY"; // Custom Symbols (comma separated)
input string               InpSymbolSuffix      = "";            // Symbol Suffix (e.g., .raw, .m)
input string               InpSymbolPrefix      = "";            // Symbol Prefix (e.g., FX:)

//--- RSI Settings
input group "=== RSI Indicator Settings ==="
input int                  InpRSIPeriod         = 14;           // RSI Period
input ENUM_APPLIED_PRICE   InpRSIAppliedPrice   = PRICE_CLOSE;  // RSI Applied Price
input ENUM_TIMEFRAMES      InpTimeframe         = PERIOD_D1;    // Trading Timeframe

//--- Entry Thresholds
input group "=== Entry Thresholds ==="
input double               InpLongEntryLevel    = 30.0;         // Long Entry RSI Level (Cross Above)
input double               InpShortEntryLevel   = 70.0;         // Short Entry RSI Level (Cross Below)

//--- Exit Thresholds
input group "=== Exit Thresholds ==="
input double               InpLongExitLevel     = 60.0;         // Long Exit RSI Level (Close Longs Above)
input double               InpShortExitLevel    = 40.0;         // Short Exit RSI Level (Close Shorts Below)

//--- Trade Direction
input group "=== Trade Direction ==="
enum ENUM_TRADE_DIRECTION
{
   TRADE_BOTH      = 0,   // Both Buy and Sell
   TRADE_BUY_ONLY  = 1,   // Buy Only
   TRADE_SELL_ONLY = 2    // Sell Only
};
input ENUM_TRADE_DIRECTION InpTradeDirection    = TRADE_BOTH;   // Trade Direction

//--- Lot Size Settings
input group "=== Lot Size Settings ==="
enum ENUM_LOT_MODE
{
   LOT_FIXED       = 0,   // Fixed Lot Size
   LOT_RISK_PERCENT= 1    // Risk Percentage of Balance
};
input ENUM_LOT_MODE        InpLotMode           = LOT_FIXED;    // Lot Size Mode
input double               InpFixedLotSize      = 0.1;          // Fixed Lot Size
input double               InpRiskPercent       = 1.0;          // Risk Percent of Balance (%)



//--- Position Limits
input group "=== Position Limits ==="
input int                  InpMaxPositionsLong  = 10;           // Max Long Positions Per Symbol
input int                  InpMaxPositionsShort = 10;           // Max Short Positions Per Symbol
input int                  InpMaxTotalPositions = 100;          // Max Total Positions (All Symbols)

//--- Global Profit Target Exit
input group "=== Global Profit Target Exit ==="
input bool                 InpUseGlobalProfitTarget = true;     // Use Global Profit Target Exit
input double               InpGlobalProfitTarget    = 1000.0;   // Global Profit Target ($)
input bool                 InpUseGlobalLossLimit    = false;    // Use Global Loss Limit Exit
input double               InpGlobalLossLimit       = 500.0;    // Global Loss Limit ($)

//--- Order Settings
input group "=== Order Settings ==="
input int                  InpMagicNumber       = 123456;       // Magic Number
input int                  InpSlippage          = 10;           // Slippage (Points)
input string               InpOrderComment      = "RSI_Daily";  // Order Comment

//--- Display Settings
input group "=== Display Settings ==="
input bool                 InpShowDashboard     = true;         // Show Dashboard on Chart
input color                InpDashboardColor    = clrWhite;     // Dashboard Text Color
input int                  InpDashboardFontSize = 10;           // Dashboard Font Size

//--- Multi-Timeframe Trend Score Settings
input group "=== MTF Trend Score Settings ==="
input bool                 InpUseMTFTrendScore     = true;      // Enable MTF Trend Score Filter
enum ENUM_KEY_TIMEFRAME
{
   KEY_TF_M1   = 0,   // 1 Minute
   KEY_TF_M5   = 1,   // 5 Minutes
   KEY_TF_M15  = 2,   // 15 Minutes
   KEY_TF_M30  = 3,   // 30 Minutes
   KEY_TF_H1   = 4,   // 1 Hour
   KEY_TF_H4   = 5,   // 4 Hours
   KEY_TF_D1   = 6    // Daily
};
input ENUM_KEY_TIMEFRAME   InpKeyTimeframe         = KEY_TF_D1; // Key Timeframe for Weighting
input double               InpMinTrendConfidence   = 40.0;      // Minimum Trend Confidence (%)

//+------------------------------------------------------------------+
//| Structure to hold symbol-specific data                           |
//+------------------------------------------------------------------+
struct SymbolData
{
   string   symbol;
   int      rsiHandle;
   double   rsiBuffer[3];
   double   pipValue;
   int      digits;
   datetime lastBarTime;
   bool     isActive;
};

//+------------------------------------------------------------------+
//| MTF Trend Score Structures and Constants                         |
//+------------------------------------------------------------------+
#define MTF_COUNT 7

//--- Timeframe enum array for indexing
ENUM_TIMEFRAMES g_mtfTimeframes[MTF_COUNT] = {
   PERIOD_M1, PERIOD_M5, PERIOD_M15, PERIOD_M30, 
   PERIOD_H1, PERIOD_H4, PERIOD_D1
};

//--- Weights for each key timeframe [keyTF][analysissTF]
double g_tfWeights[7][7] = {
   {0.35, 0.25, 0.15, 0.10, 0.08, 0.05, 0.02}, // 1m key
   {0.10, 0.30, 0.25, 0.15, 0.10, 0.07, 0.03}, // 5m key
   {0.05, 0.15, 0.30, 0.20, 0.15, 0.10, 0.05}, // 15m key
   {0.03, 0.08, 0.15, 0.28, 0.22, 0.16, 0.08}, // 30m key
   {0.02, 0.05, 0.10, 0.15, 0.30, 0.25, 0.13}, // 1h key
   {0.00, 0.03, 0.07, 0.10, 0.15, 0.35, 0.30}, // 4h key
   {0.00, 0.00, 0.00, 0.02, 0.08, 0.25, 0.40}  // 1d key
};

//--- MTF Indicator handles per timeframe
struct MTFHandles
{
   int adxHandle;
   int ema20Handle;
   int ema50Handle;
   int ema200Handle;
   int macdHandle;
   int atrHandle;
};

//--- MTF data per symbol
struct SymbolMTFData
{
   MTFHandles handles[MTF_COUNT];  // Handles for each of 7 timeframes
   bool isInitialized;
};

//+------------------------------------------------------------------+
//| Global Variables                                                 |
//+------------------------------------------------------------------+
CTrade         trade;
SymbolData     g_symbols[];
int            g_activeSymbolCount = 0;
string         g_customSymbols[];
SymbolMTFData  g_mtfData[];  // MTF indicator handles per symbol

//+------------------------------------------------------------------+
//| Expert initialization function                                   |
//+------------------------------------------------------------------+
int OnInit()
{
   //--- Validate inputs
   if(InpRSIPeriod < 1)
   {
      Print("Error: RSI Period must be >= 1");
      return(INIT_PARAMETERS_INCORRECT);
   }
   
   if(InpLongEntryLevel <= 0 || InpLongEntryLevel >= 100)
   {
      Print("Error: Long Entry Level must be between 0 and 100");
      return(INIT_PARAMETERS_INCORRECT);
   }
   
   if(InpShortEntryLevel <= 0 || InpShortEntryLevel >= 100)
   {
      Print("Error: Short Entry Level must be between 0 and 100");
      return(INIT_PARAMETERS_INCORRECT);
   }
   
   if(InpLongExitLevel <= InpLongEntryLevel)
   {
      Print("Warning: Long Exit Level should be greater than Long Entry Level");
   }
   
   if(InpShortExitLevel >= InpShortEntryLevel)
   {
      Print("Warning: Short Exit Level should be less than Short Entry Level");
   }
   
   //--- Setup trade object
   trade.SetExpertMagicNumber(InpMagicNumber);
   trade.SetDeviationInPoints(InpSlippage);
   trade.SetTypeFilling(ORDER_FILLING_IOC);
   
   //--- Initialize symbols based on mode
   if(!InitializeSymbols())
   {
      Print("Error: Failed to initialize symbols");
      return(INIT_FAILED);
   }
   
   Print("RSI Daily Strategy EA initialized successfully");
   Print("Symbol Mode: ", EnumToString(InpSymbolMode));
   Print("Active Symbols: ", g_activeSymbolCount);
   Print("Timeframe: ", EnumToString(InpTimeframe));
   
   //--- List all active symbols
   for(int i = 0; i < g_activeSymbolCount; i++)
   {
      if(g_symbols[i].isActive)
         Print("Symbol ", (i+1), ": ", g_symbols[i].symbol);
   }
   
   return(INIT_SUCCEEDED);
}

//+------------------------------------------------------------------+
//| Initialize symbols based on selected mode                        |
//+------------------------------------------------------------------+
bool InitializeSymbols()
{
   int symbolCount = 0;
   
   switch(InpSymbolMode)
   {
      case SYMBOL_CURRENT_ONLY:
         symbolCount = 1;
         ArrayResize(g_symbols, symbolCount);
         g_symbols[0].symbol = _Symbol;
         break;
         
      case SYMBOL_MULTI_ALL:
         symbolCount = MAX_SYMBOLS;
         ArrayResize(g_symbols, symbolCount);
         for(int i = 0; i < MAX_SYMBOLS; i++)
            g_symbols[i].symbol = InpSymbolPrefix + g_symbolList[i] + InpSymbolSuffix;
         break;
         
      case SYMBOL_MAJORS_ONLY:
         symbolCount = 7;
         ArrayResize(g_symbols, symbolCount);
         for(int i = 0; i < 7; i++)
            g_symbols[i].symbol = InpSymbolPrefix + g_symbolList[i] + InpSymbolSuffix;
         break;
         
      case SYMBOL_MINORS_ONLY:
         symbolCount = 21;
         ArrayResize(g_symbols, symbolCount);
         for(int i = 0; i < 21; i++)
            g_symbols[i].symbol = InpSymbolPrefix + g_symbolList[i + 7] + InpSymbolSuffix;
         break;
         
      case SYMBOL_CUSTOM_LIST:
         ParseCustomSymbols();
         symbolCount = ArraySize(g_customSymbols);
         ArrayResize(g_symbols, symbolCount);
         for(int i = 0; i < symbolCount; i++)
            g_symbols[i].symbol = InpSymbolPrefix + g_customSymbols[i] + InpSymbolSuffix;
         break;
   }
   
   //--- Resize MTF data array to match symbols
   if(InpUseMTFTrendScore)
   {
      ArrayResize(g_mtfData, symbolCount);
      for(int i = 0; i < symbolCount; i++)
         g_mtfData[i].isInitialized = false;
   }
   
   //--- Initialize each symbol
   g_activeSymbolCount = 0;
   for(int i = 0; i < symbolCount; i++)
   {
      g_symbols[i].isActive = false;
      g_symbols[i].rsiHandle = INVALID_HANDLE;
      g_symbols[i].lastBarTime = 0;
      
      //--- Check if symbol exists
      if(!SymbolSelect(g_symbols[i].symbol, true))
      {
         Print("Warning: Symbol not found: ", g_symbols[i].symbol);
         continue;
      }
      
      //--- Create RSI handle for this symbol
      g_symbols[i].rsiHandle = iRSI(g_symbols[i].symbol, InpTimeframe, InpRSIPeriod, InpRSIAppliedPrice);
      if(g_symbols[i].rsiHandle == INVALID_HANDLE)
      {
         Print("Error: Failed to create RSI handle for ", g_symbols[i].symbol);
         continue;
      }
      
      //--- Determine pip value
      g_symbols[i].digits = (int)SymbolInfoInteger(g_symbols[i].symbol, SYMBOL_DIGITS);
      if(g_symbols[i].digits == 3 || g_symbols[i].digits == 5)
         g_symbols[i].pipValue = SymbolInfoDouble(g_symbols[i].symbol, SYMBOL_POINT) * 10;
      else
         g_symbols[i].pipValue = SymbolInfoDouble(g_symbols[i].symbol, SYMBOL_POINT);
      
      ArraySetAsSeries(g_symbols[i].rsiBuffer, true);
      g_symbols[i].isActive = true;
      g_activeSymbolCount++;
      
      //--- Initialize MTF handles for this symbol
      if(InpUseMTFTrendScore)
      {
         if(!InitializeMTFHandles(i))
         {
            Print("Warning: Failed to initialize MTF handles for ", g_symbols[i].symbol, 
                  " - MTF scoring will be disabled for this symbol");
         }
      }
   }
   
   return (g_activeSymbolCount > 0);
}

//+------------------------------------------------------------------+
//| Parse custom symbols from input string                           |
//+------------------------------------------------------------------+
void ParseCustomSymbols()
{
   string tempSymbols = InpCustomSymbols;
   StringReplace(tempSymbols, " ", "");
   
   string result[];
   int count = StringSplit(tempSymbols, ',', result);
   
   ArrayResize(g_customSymbols, count);
   for(int i = 0; i < count; i++)
   {
      g_customSymbols[i] = result[i];
   }
}

//+------------------------------------------------------------------+
//| Expert deinitialization function                                 |
//+------------------------------------------------------------------+
void OnDeinit(const int reason)
{
   //--- Release all indicator handles
   for(int i = 0; i < ArraySize(g_symbols); i++)
   {
      if(g_symbols[i].rsiHandle != INVALID_HANDLE)
      {
         IndicatorRelease(g_symbols[i].rsiHandle);
         g_symbols[i].rsiHandle = INVALID_HANDLE;
      }
   }
   
   //--- Release MTF indicator handles
   if(InpUseMTFTrendScore)
   {
      for(int i = 0; i < ArraySize(g_mtfData); i++)
      {
         ReleaseMTFHandles(i);
      }
   }
   
   //--- Remove dashboard objects
   ObjectsDeleteAll(0, "RSI_Dashboard_");
   
   Print("RSI Daily Strategy EA deinitialized. Reason: ", reason);
}

//+------------------------------------------------------------------+
//| Expert tick function                                             |
//+------------------------------------------------------------------+
void OnTick()
{
   //--- Check global profit/loss target FIRST before any other processing
   if(CheckGlobalProfitTarget())
   {
      //--- Global target hit, all positions closed, skip this tick
      if(InpShowDashboard) UpdateDashboard();
      return;
   }
   
   //--- Process each active symbol
   for(int i = 0; i < ArraySize(g_symbols); i++)
   {
      if(!g_symbols[i].isActive)
         continue;
      
      ProcessSymbol(i);
   }
   

   
   //--- Update dashboard
   if(InpShowDashboard) UpdateDashboard();
}

//+------------------------------------------------------------------+
//| Process a single symbol                                          |
//+------------------------------------------------------------------+
void ProcessSymbol(int idx)
{
   string symbol = g_symbols[idx].symbol;
   
   //--- Check for new bar on the trading timeframe
   datetime currentBarTime = iTime(symbol, InpTimeframe, 0);
   if(currentBarTime == g_symbols[idx].lastBarTime)
      return;
   
   g_symbols[idx].lastBarTime = currentBarTime;
   
   //--- Get RSI values
   if(CopyBuffer(g_symbols[idx].rsiHandle, 0, 0, 3, g_symbols[idx].rsiBuffer) < 3)
   {
      Print("Error copying RSI buffer for ", symbol);
      return;
   }
   
   double rsiCurrent = g_symbols[idx].rsiBuffer[1];
   double rsiPrevious = g_symbols[idx].rsiBuffer[2];
   
   //--- Check exit conditions first
   CheckExitConditions(symbol, rsiCurrent);
   
   //--- Check entry conditions
   if(IsTradeAllowed())
   {
      CheckEntryConditions(idx, symbol, rsiCurrent, rsiPrevious);
   }
}

//+------------------------------------------------------------------+
//| Check if trading is allowed                                      |
//+------------------------------------------------------------------+
bool IsTradeAllowed()
{
   //--- Check day of week (no trading on weekends)
   MqlDateTime dt;
   TimeToStruct(TimeCurrent(), dt);
   
   if(dt.day_of_week == 0 || dt.day_of_week == 6)
      return false;
   
   //--- Check total position limit
   if(CountAllPositions() >= InpMaxTotalPositions)
      return false;
   
   return true;
}

//+------------------------------------------------------------------+
//| Check entry conditions for a symbol                              |
//+------------------------------------------------------------------+
void CheckEntryConditions(int idx, string symbol, double rsiCurrent, double rsiPrevious)
{
   //--- Long Entry: RSI crosses above oversold level
   if(InpTradeDirection != TRADE_SELL_ONLY)
   {
      if(rsiPrevious < InpLongEntryLevel && rsiCurrent >= InpLongEntryLevel)
      {
         int currentLongs = CountPositions(symbol, POSITION_TYPE_BUY);
         
         if(currentLongs < InpMaxPositionsLong)
         {
            //--- Check MTF Bearish Trend Score (confirming downtrend before reversal)
            if(InpUseMTFTrendScore)
            {
               double bearishScore = GetMTFBearishScore(idx);
               if(bearishScore < InpMinTrendConfidence)
               {
                  Print(symbol, " LONG SIGNAL REJECTED: MTF Bearish Score ", 
                        DoubleToString(bearishScore, 1), "% < ", 
                        DoubleToString(InpMinTrendConfidence, 1), "% threshold");
                  return;
               }
               Print(symbol, " MTF Bearish Score: ", DoubleToString(bearishScore, 1), "% - APPROVED for BUY");
            }
            
            Print(symbol, " LONG SIGNAL: RSI crossed above ", InpLongEntryLevel, 
                  " (Prev: ", DoubleToString(rsiPrevious, 2), ", Curr: ", DoubleToString(rsiCurrent, 2), ")");
            OpenTrade(idx, symbol, ORDER_TYPE_BUY);
         }
      }
   }
   
   //--- Short Entry: RSI crosses below overbought level
   if(InpTradeDirection != TRADE_BUY_ONLY)
   {
      if(rsiPrevious > InpShortEntryLevel && rsiCurrent <= InpShortEntryLevel)
      {
         int currentShorts = CountPositions(symbol, POSITION_TYPE_SELL);
         
         if(currentShorts < InpMaxPositionsShort)
         {
            //--- Check MTF Bullish Trend Score (confirming uptrend before reversal)
            if(InpUseMTFTrendScore)
            {
               double bullishScore = GetMTFBullishScore(idx);
               if(bullishScore < InpMinTrendConfidence)
               {
                  Print(symbol, " SHORT SIGNAL REJECTED: MTF Bullish Score ", 
                        DoubleToString(bullishScore, 1), "% < ", 
                        DoubleToString(InpMinTrendConfidence, 1), "% threshold");
                  return;
               }
               Print(symbol, " MTF Bullish Score: ", DoubleToString(bullishScore, 1), "% - APPROVED for SELL");
            }
            
            Print(symbol, " SHORT SIGNAL: RSI crossed below ", InpShortEntryLevel,
                  " (Prev: ", DoubleToString(rsiPrevious, 2), ", Curr: ", DoubleToString(rsiCurrent, 2), ")");
            OpenTrade(idx, symbol, ORDER_TYPE_SELL);
         }
      }
   }
}

//+------------------------------------------------------------------+
//| Check exit conditions for a symbol                               |
//+------------------------------------------------------------------+
void CheckExitConditions(string symbol, double rsiCurrent)
{
   //--- Long Exit: RSI exceeds exit level
   if(rsiCurrent > InpLongExitLevel)
   {
      int closedCount = CloseAllPositions(symbol, POSITION_TYPE_BUY);
      if(closedCount > 0)
      {
         Print(symbol, " LONG EXIT: RSI (", DoubleToString(rsiCurrent, 2), 
               ") > ", InpLongExitLevel, " | Closed ", closedCount, " positions");
      }
   }
   
   //--- Short Exit: RSI drops below exit level
   if(rsiCurrent < InpShortExitLevel)
   {
      int closedCount = CloseAllPositions(symbol, POSITION_TYPE_SELL);
      if(closedCount > 0)
      {
         Print(symbol, " SHORT EXIT: RSI (", DoubleToString(rsiCurrent, 2), 
               ") < ", InpShortExitLevel, " | Closed ", closedCount, " positions");
      }
   }
}

//+------------------------------------------------------------------+
//| Open a trade for a specific symbol                               |
//+------------------------------------------------------------------+
bool OpenTrade(int idx, string symbol, ENUM_ORDER_TYPE orderType)
{
   double lotSize = CalculateLotSize(symbol);
   if(lotSize <= 0)
   {
      Print("Error: Invalid lot size calculated for ", symbol);
      return false;
   }
   
   double price = 0;
   
   if(orderType == ORDER_TYPE_BUY)
      price = SymbolInfoDouble(symbol, SYMBOL_ASK);
   else if(orderType == ORDER_TYPE_SELL)
      price = SymbolInfoDouble(symbol, SYMBOL_BID);
   
   bool result = trade.PositionOpen(symbol, orderType, lotSize, price, 0, 0, InpOrderComment);
   
   if(result)
   {
      Print(symbol, " Trade Opened: ", (orderType == ORDER_TYPE_BUY ? "BUY" : "SELL"), 
            " | Lot: ", lotSize, " | Price: ", price);
   }
   else
   {
      Print(symbol, " Trade Failed: Error ", GetLastError());
   }
   
   return result;
}

//+------------------------------------------------------------------+
//| Calculate lot size for a specific symbol                         |
//+------------------------------------------------------------------+
double CalculateLotSize(string symbol)
{
   double lotSize = InpFixedLotSize;
   
   //--- Validate lot size
   double minLot = SymbolInfoDouble(symbol, SYMBOL_VOLUME_MIN);
   double maxLot = SymbolInfoDouble(symbol, SYMBOL_VOLUME_MAX);
   double stepLot = SymbolInfoDouble(symbol, SYMBOL_VOLUME_STEP);
   
   if(lotSize < minLot) lotSize = minLot;
   if(lotSize > maxLot) lotSize = maxLot;
   
   lotSize = MathFloor(lotSize / stepLot) * stepLot;
   
   return NormalizeDouble(lotSize, 2);
}

//+------------------------------------------------------------------+
//| Count positions for a specific symbol and type                   |
//+------------------------------------------------------------------+
int CountPositions(string symbol, ENUM_POSITION_TYPE posType)
{
   int count = 0;
   for(int i = PositionsTotal() - 1; i >= 0; i--)
   {
      if(PositionSelectByTicket(PositionGetTicket(i)))
      {
         if(PositionGetString(POSITION_SYMBOL) == symbol &&
            PositionGetInteger(POSITION_MAGIC) == InpMagicNumber &&
            PositionGetInteger(POSITION_TYPE) == posType)
         {
            count++;
         }
      }
   }
   return count;
}

//+------------------------------------------------------------------+
//| Count all positions for this EA                                  |
//+------------------------------------------------------------------+
int CountAllPositions()
{
   int count = 0;
   for(int i = PositionsTotal() - 1; i >= 0; i--)
   {
      if(PositionSelectByTicket(PositionGetTicket(i)))
      {
         if(PositionGetInteger(POSITION_MAGIC) == InpMagicNumber)
         {
            count++;
         }
      }
   }
   return count;
}

//+------------------------------------------------------------------+
//| Close all positions for a specific symbol and type               |
//+------------------------------------------------------------------+
int CloseAllPositions(string symbol, ENUM_POSITION_TYPE posType)
{
   int closedCount = 0;
   for(int i = PositionsTotal() - 1; i >= 0; i--)
   {
      if(PositionSelectByTicket(PositionGetTicket(i)))
      {
         if(PositionGetString(POSITION_SYMBOL) == symbol &&
            PositionGetInteger(POSITION_MAGIC) == InpMagicNumber &&
            PositionGetInteger(POSITION_TYPE) == posType)
         {
            ulong ticket = PositionGetTicket(i);
            if(trade.PositionClose(ticket))
            {
               closedCount++;
            }
            else
            {
               Print("Failed to close position #", ticket, " Error: ", GetLastError());
            }
         }
      }
   }
   return closedCount;
}



//+------------------------------------------------------------------+
//| Update dashboard                                                 |
//+------------------------------------------------------------------+
void UpdateDashboard()
{
   string dashPrefix = "RSI_Dashboard_";
   int startX = 10;
   int startY = 30;
   int lineHeight = 18;
   int line = 0;
   
   //--- Title
   CreateLabel(dashPrefix + "Title", startX, startY + lineHeight * line++, 
               "=== RSI Daily Strategy v2.0 ===", InpDashboardColor, InpDashboardFontSize + 2);
   
   //--- Symbol Mode
   string modeStr = "";
   switch(InpSymbolMode)
   {
      case SYMBOL_CURRENT_ONLY: modeStr = "Current Only"; break;
      case SYMBOL_MULTI_ALL:    modeStr = "All Pairs (28)"; break;
      case SYMBOL_MAJORS_ONLY:  modeStr = "Majors (7)"; break;
      case SYMBOL_MINORS_ONLY:  modeStr = "Minors (21)"; break;
      case SYMBOL_CUSTOM_LIST:  modeStr = "Custom"; break;
   }
   
   CreateLabel(dashPrefix + "Mode", startX, startY + lineHeight * line++,
               "Mode: " + modeStr + " | Active: " + IntegerToString(g_activeSymbolCount),
               InpDashboardColor, InpDashboardFontSize);
   
   //--- Thresholds
   CreateLabel(dashPrefix + "Thresholds", startX, startY + lineHeight * line++, 
               "Entry: Long<" + DoubleToString(InpLongEntryLevel, 0) + 
               " | Short>" + DoubleToString(InpShortEntryLevel, 0), 
               InpDashboardColor, InpDashboardFontSize);
   
   CreateLabel(dashPrefix + "Exits", startX, startY + lineHeight * line++, 
               "Exit: Long>" + DoubleToString(InpLongExitLevel, 0) + 
               " | Short<" + DoubleToString(InpShortExitLevel, 0), 
               InpDashboardColor, InpDashboardFontSize);
   
   //--- Count positions
   int totalLongs = 0;
   int totalShorts = 0;
   for(int i = PositionsTotal() - 1; i >= 0; i--)
   {
      if(PositionSelectByTicket(PositionGetTicket(i)))
      {
         if(PositionGetInteger(POSITION_MAGIC) == InpMagicNumber)
         {
            if(PositionGetInteger(POSITION_TYPE) == POSITION_TYPE_BUY)
               totalLongs++;
            else
               totalShorts++;
         }
      }
   }
   
   //--- Positions
   CreateLabel(dashPrefix + "Positions", startX, startY + lineHeight * line++, 
               "Positions: Longs=" + IntegerToString(totalLongs) + 
               " | Shorts=" + IntegerToString(totalShorts), 
               InpDashboardColor, InpDashboardFontSize);
   
   //--- Total Profit
   double totalProfit = CalculateTotalProfit();
   color profitColor = totalProfit >= 0 ? clrLime : clrRed;
   CreateLabel(dashPrefix + "Profit", startX, startY + lineHeight * line++, 
               "Floating P/L: $" + DoubleToString(totalProfit, 2), 
               profitColor, InpDashboardFontSize);
   
   //--- Trade Direction
   string directionStr = "Both";
   if(InpTradeDirection == TRADE_BUY_ONLY) directionStr = "Buy Only";
   else if(InpTradeDirection == TRADE_SELL_ONLY) directionStr = "Sell Only";
   
   CreateLabel(dashPrefix + "Direction", startX, startY + lineHeight * line++, 
               "Direction: " + directionStr, 
               InpDashboardColor, InpDashboardFontSize);
   
   //--- MTF Trend Score Status
   if(InpUseMTFTrendScore)
   {
      line++;
      string keyTFStr = "";
      switch(InpKeyTimeframe)
      {
         case KEY_TF_M1:  keyTFStr = "1m"; break;
         case KEY_TF_M5:  keyTFStr = "5m"; break;
         case KEY_TF_M15: keyTFStr = "15m"; break;
         case KEY_TF_M30: keyTFStr = "30m"; break;
         case KEY_TF_H1:  keyTFStr = "1h"; break;
         case KEY_TF_H4:  keyTFStr = "4h"; break;
         case KEY_TF_D1:  keyTFStr = "1d"; break;
      }
      
      CreateLabel(dashPrefix + "MTFHeader", startX, startY + lineHeight * line++,
                  "--- MTF Trend Score (Key: " + keyTFStr + ") ---",
                  clrCyan, InpDashboardFontSize);
      
      // Show MTF scores for first active symbol (or current symbol)
      for(int i = 0; i < ArraySize(g_symbols); i++)
      {
         if(g_symbols[i].isActive && g_mtfData[i].isInitialized)
         {
            double bullScore = GetMTFBullishScore(i);
            double bearScore = GetMTFBearishScore(i);
            
            color bullColor = bullScore >= InpMinTrendConfidence ? clrLime : clrOrange;
            color bearColor = bearScore >= InpMinTrendConfidence ? clrRed : clrOrange;
            
            CreateLabel(dashPrefix + "MTFBull", startX, startY + lineHeight * line++,
                        g_symbols[i].symbol + " Bull: " + DoubleToString(bullScore, 1) + "%",
                        bullColor, InpDashboardFontSize);
            
            CreateLabel(dashPrefix + "MTFBear", startX, startY + lineHeight * line++,
                        g_symbols[i].symbol + " Bear: " + DoubleToString(bearScore, 1) + "%",
                        bearColor, InpDashboardFontSize);
            
            CreateLabel(dashPrefix + "MTFThresh", startX, startY + lineHeight * line++,
                        "Min Threshold: " + DoubleToString(InpMinTrendConfidence, 1) + "%",
                        InpDashboardColor, InpDashboardFontSize);
            
            break;  // Only show first active symbol
         }
      }
   }
   
   //--- Show active symbols with signals (limited to 10 for display)
   line++;
   CreateLabel(dashPrefix + "SymHeader", startX, startY + lineHeight * line++, 
               "--- Active Symbol RSI Values ---", 
               clrYellow, InpDashboardFontSize);
   
   int displayed = 0;
   for(int i = 0; i < ArraySize(g_symbols) && displayed < 10; i++)
   {
      if(!g_symbols[i].isActive) continue;
      
      double rsiVal = g_symbols[i].rsiBuffer[1];
      string rsiStatus = "";
      color symColor = InpDashboardColor;
      
      if(rsiVal <= InpLongEntryLevel)
      {
         rsiStatus = " [OVERSOLD]";
         symColor = clrLime;
      }
      else if(rsiVal >= InpShortEntryLevel)
      {
         rsiStatus = " [OVERBOUGHT]";
         symColor = clrRed;
      }
      
      CreateLabel(dashPrefix + "Sym" + IntegerToString(i), startX, startY + lineHeight * line++,
                  g_symbols[i].symbol + ": " + DoubleToString(rsiVal, 2) + rsiStatus,
                  symColor, InpDashboardFontSize - 1);
      displayed++;
   }
   
   if(g_activeSymbolCount > 10)
   {
      CreateLabel(dashPrefix + "MoreSyms", startX, startY + lineHeight * line++,
                  "... and " + IntegerToString(g_activeSymbolCount - 10) + " more symbols",
                  clrGray, InpDashboardFontSize - 1);
   }
}

//+------------------------------------------------------------------+
//| Create a text label on chart                                     |
//+------------------------------------------------------------------+
void CreateLabel(string name, int x, int y, string text, color clr, int fontSize)
{
   if(ObjectFind(0, name) < 0)
   {
      ObjectCreate(0, name, OBJ_LABEL, 0, 0, 0);
      ObjectSetInteger(0, name, OBJPROP_CORNER, CORNER_LEFT_UPPER);
      ObjectSetInteger(0, name, OBJPROP_ANCHOR, ANCHOR_LEFT_UPPER);
   }
   
   ObjectSetInteger(0, name, OBJPROP_XDISTANCE, x);
   ObjectSetInteger(0, name, OBJPROP_YDISTANCE, y);
   ObjectSetString(0, name, OBJPROP_TEXT, text);
   ObjectSetInteger(0, name, OBJPROP_COLOR, clr);
   ObjectSetInteger(0, name, OBJPROP_FONTSIZE, fontSize);
   ObjectSetString(0, name, OBJPROP_FONT, "Consolas");
}

//+------------------------------------------------------------------+
//| Calculate total floating profit                                  |
//+------------------------------------------------------------------+
double CalculateTotalProfit()
{
   double profit = 0;
   for(int i = PositionsTotal() - 1; i >= 0; i--)
   {
      if(PositionSelectByTicket(PositionGetTicket(i)))
      {
         if(PositionGetInteger(POSITION_MAGIC) == InpMagicNumber)
         {
            profit += PositionGetDouble(POSITION_PROFIT) + PositionGetDouble(POSITION_SWAP);
         }
      }
   }
   return profit;
}

//+------------------------------------------------------------------+
//| Check global profit/loss target and close all if reached         |
//+------------------------------------------------------------------+
bool CheckGlobalProfitTarget()
{
   //--- Get current total floating P/L
   double totalProfit = CalculateTotalProfit();
   
   //--- Check if profit target is reached
   if(InpUseGlobalProfitTarget && totalProfit >= InpGlobalProfitTarget)
   {
      Print("=== GLOBAL PROFIT TARGET REACHED ===");
      Print("Total Floating P/L: $", DoubleToString(totalProfit, 2), 
            " >= Target: $", DoubleToString(InpGlobalProfitTarget, 2));
      
      int closedCount = CloseAllTrades();
      
      Print("Closed ", closedCount, " positions. Starting fresh cycle...");
      Print("==========================================");
      
      //--- Reset last bar times to allow fresh signals on next bar
      for(int i = 0; i < ArraySize(g_symbols); i++)
      {
         g_symbols[i].lastBarTime = 0;
      }
      
      return true;
   }
   
   //--- Check if loss limit is reached
   if(InpUseGlobalLossLimit && totalProfit <= -InpGlobalLossLimit)
   {
      Print("=== GLOBAL LOSS LIMIT REACHED ===");
      Print("Total Floating P/L: $", DoubleToString(totalProfit, 2), 
            " <= Loss Limit: -$", DoubleToString(InpGlobalLossLimit, 2));
      
      int closedCount = CloseAllTrades();
      
      Print("Closed ", closedCount, " positions. Starting fresh cycle...");
      Print("==========================================");
      
      //--- Reset last bar times to allow fresh signals on next bar
      for(int i = 0; i < ArraySize(g_symbols); i++)
      {
         g_symbols[i].lastBarTime = 0;
      }
      
      return true;
   }
   
   return false;
}

//+------------------------------------------------------------------+
//| Close all trades for this EA (all symbols)                       |
//+------------------------------------------------------------------+
int CloseAllTrades()
{
   int closedCount = 0;
   
   for(int i = PositionsTotal() - 1; i >= 0; i--)
   {
      if(PositionSelectByTicket(PositionGetTicket(i)))
      {
         if(PositionGetInteger(POSITION_MAGIC) == InpMagicNumber)
         {
            ulong ticket = PositionGetTicket(i);
            string symbol = PositionGetString(POSITION_SYMBOL);
            double profit = PositionGetDouble(POSITION_PROFIT);
            
            if(trade.PositionClose(ticket))
            {
               Print("Closed position #", ticket, " on ", symbol, 
                     " | P/L: $", DoubleToString(profit, 2));
               closedCount++;
            }
            else
            {
               Print("Failed to close position #", ticket, " Error: ", GetLastError());
            }
         }
      }
   }
   
   return closedCount;
}

//+------------------------------------------------------------------+
//| Initialize MTF indicator handles for a symbol                    |
//+------------------------------------------------------------------+
bool InitializeMTFHandles(int symbolIdx)
{
   string symbol = g_symbols[symbolIdx].symbol;
   g_mtfData[symbolIdx].isInitialized = false;
   
   for(int tf = 0; tf < MTF_COUNT; tf++)
   {
      ENUM_TIMEFRAMES timeframe = g_mtfTimeframes[tf];
      
      // Initialize all handles to INVALID_HANDLE first
      g_mtfData[symbolIdx].handles[tf].adxHandle = INVALID_HANDLE;
      g_mtfData[symbolIdx].handles[tf].ema20Handle = INVALID_HANDLE;
      g_mtfData[symbolIdx].handles[tf].ema50Handle = INVALID_HANDLE;
      g_mtfData[symbolIdx].handles[tf].ema200Handle = INVALID_HANDLE;
      g_mtfData[symbolIdx].handles[tf].macdHandle = INVALID_HANDLE;
      g_mtfData[symbolIdx].handles[tf].atrHandle = INVALID_HANDLE;
      
      // Create handles
      g_mtfData[symbolIdx].handles[tf].adxHandle = iADX(symbol, timeframe, 14);
      g_mtfData[symbolIdx].handles[tf].ema20Handle = iMA(symbol, timeframe, 20, 0, MODE_EMA, PRICE_CLOSE);
      g_mtfData[symbolIdx].handles[tf].ema50Handle = iMA(symbol, timeframe, 50, 0, MODE_EMA, PRICE_CLOSE);
      g_mtfData[symbolIdx].handles[tf].ema200Handle = iMA(symbol, timeframe, 200, 0, MODE_EMA, PRICE_CLOSE);
      g_mtfData[symbolIdx].handles[tf].macdHandle = iMACD(symbol, timeframe, 12, 26, 9, PRICE_CLOSE);
      g_mtfData[symbolIdx].handles[tf].atrHandle = iATR(symbol, timeframe, 14);
      
      // Validate handles
      if(g_mtfData[symbolIdx].handles[tf].adxHandle == INVALID_HANDLE ||
         g_mtfData[symbolIdx].handles[tf].ema20Handle == INVALID_HANDLE ||
         g_mtfData[symbolIdx].handles[tf].ema50Handle == INVALID_HANDLE ||
         g_mtfData[symbolIdx].handles[tf].ema200Handle == INVALID_HANDLE ||
         g_mtfData[symbolIdx].handles[tf].macdHandle == INVALID_HANDLE ||
         g_mtfData[symbolIdx].handles[tf].atrHandle == INVALID_HANDLE)
      {
         Print("Error creating MTF handles for ", symbol, " on ", EnumToString(timeframe));
         return false;
      }
   }
   
   g_mtfData[symbolIdx].isInitialized = true;
   Print("MTF handles initialized for ", symbol);
   return true;
}

//+------------------------------------------------------------------+
//| Release MTF indicator handles for a symbol                       |
//+------------------------------------------------------------------+
void ReleaseMTFHandles(int symbolIdx)
{
   if(!g_mtfData[symbolIdx].isInitialized)
      return;
      
   for(int tf = 0; tf < MTF_COUNT; tf++)
   {
      if(g_mtfData[symbolIdx].handles[tf].adxHandle != INVALID_HANDLE)
         IndicatorRelease(g_mtfData[symbolIdx].handles[tf].adxHandle);
      if(g_mtfData[symbolIdx].handles[tf].ema20Handle != INVALID_HANDLE)
         IndicatorRelease(g_mtfData[symbolIdx].handles[tf].ema20Handle);
      if(g_mtfData[symbolIdx].handles[tf].ema50Handle != INVALID_HANDLE)
         IndicatorRelease(g_mtfData[symbolIdx].handles[tf].ema50Handle);
      if(g_mtfData[symbolIdx].handles[tf].ema200Handle != INVALID_HANDLE)
         IndicatorRelease(g_mtfData[symbolIdx].handles[tf].ema200Handle);
      if(g_mtfData[symbolIdx].handles[tf].macdHandle != INVALID_HANDLE)
         IndicatorRelease(g_mtfData[symbolIdx].handles[tf].macdHandle);
      if(g_mtfData[symbolIdx].handles[tf].atrHandle != INVALID_HANDLE)
         IndicatorRelease(g_mtfData[symbolIdx].handles[tf].atrHandle);
   }
   
   g_mtfData[symbolIdx].isInitialized = false;
}

//+------------------------------------------------------------------+
//| Calculate percentile rank of a value in an array                 |
//+------------------------------------------------------------------+
double PercentileRank(double value, double &array[], int lookback)
{
   int count = 0;
   int total = MathMin(lookback, ArraySize(array));
   
   if(total <= 0)
      return 50.0;  // Default to middle if no data
   
   for(int i = 0; i < total; i++)
   {
      if(value > array[i])
         count++;
   }
   
   return ((double)count / (double)total) * 100.0;
}

//+------------------------------------------------------------------+
//| Detect bullish trend continuation for a symbol/timeframe         |
//| Returns confidence score 0.0 to 1.0                              |
//| SIMPLIFIED PRACTICAL RULES - 3 Core Checks                       |
//+------------------------------------------------------------------+
double DetectBullishTrendContinuation(int symbolIdx, int tfIdx)
{
   if(!g_mtfData[symbolIdx].isInitialized)
      return 0.5;  // Return neutral if not initialized
      
   string symbol = g_symbols[symbolIdx].symbol;
   ENUM_TIMEFRAMES tf = g_mtfTimeframes[tfIdx];
   
   double score = 0.0;
   double maxScore = 100.0;
   
   // Buffers for indicator data
   double adxBuffer[3], adxPlusBuffer[3], adxMinusBuffer[3];
   double ema20Buffer[2], ema50Buffer[2], ema200Buffer[2];
   double macdMainBuffer[3], macdSignalBuffer[3];
   double closeBuffer[5];
   
   // Copy ADX data (buffer 0 = ADX, buffer 1 = +DI, buffer 2 = -DI)
   if(CopyBuffer(g_mtfData[symbolIdx].handles[tfIdx].adxHandle, 0, 0, 3, adxBuffer) < 3)
      return 0.5;
   if(CopyBuffer(g_mtfData[symbolIdx].handles[tfIdx].adxHandle, 1, 0, 3, adxPlusBuffer) < 3)
      return 0.5;
   if(CopyBuffer(g_mtfData[symbolIdx].handles[tfIdx].adxHandle, 2, 0, 3, adxMinusBuffer) < 3)
      return 0.5;
   if(CopyBuffer(g_mtfData[symbolIdx].handles[tfIdx].ema20Handle, 0, 0, 2, ema20Buffer) < 2)
      return 0.5;
   if(CopyBuffer(g_mtfData[symbolIdx].handles[tfIdx].ema50Handle, 0, 0, 2, ema50Buffer) < 2)
      return 0.5;
   if(CopyBuffer(g_mtfData[symbolIdx].handles[tfIdx].ema200Handle, 0, 0, 2, ema200Buffer) < 2)
      return 0.5;
   if(CopyBuffer(g_mtfData[symbolIdx].handles[tfIdx].macdHandle, 0, 0, 3, macdMainBuffer) < 3)
      return 0.5;
   if(CopyBuffer(g_mtfData[symbolIdx].handles[tfIdx].macdHandle, 1, 0, 3, macdSignalBuffer) < 3)
      return 0.5;
   if(CopyClose(symbol, tf, 0, 5, closeBuffer) < 5)
      return 0.5;
   
   // Set arrays as series (newest at index 0)
   ArraySetAsSeries(adxBuffer, true);
   ArraySetAsSeries(adxPlusBuffer, true);
   ArraySetAsSeries(adxMinusBuffer, true);
   ArraySetAsSeries(ema20Buffer, true);
   ArraySetAsSeries(ema50Buffer, true);
   ArraySetAsSeries(ema200Buffer, true);
   ArraySetAsSeries(macdMainBuffer, true);
   ArraySetAsSeries(macdSignalBuffer, true);
   ArraySetAsSeries(closeBuffer, true);
   
   double price = closeBuffer[0];
   double e200 = ema200Buffer[0];
   double e50 = ema50Buffer[0];
   double e20 = ema20Buffer[0];
   double adxVal = adxBuffer[0];
   double plusDI = adxPlusBuffer[0];
   double minusDI = adxMinusBuffer[0];
   
   //=========================================================
   // RULE 1: TREND PRESENCE (35 points max)
   // - ADX > 20 indicates a trend exists
   // - +DI > -DI indicates bullish direction
   //=========================================================
   double trendPoints = 0;
   
   // Base points if any trend exists (ADX > 15)
   if(adxVal > 15)
      trendPoints += 15;
   
   // Bonus for stronger trend
   if(adxVal > 20)
      trendPoints += 10;
   if(adxVal > 25)
      trendPoints += 5;
   
   // Direction bonus: +DI > -DI for bullish
   if(plusDI > minusDI)
      trendPoints += 5;
      
   score += trendPoints;
   
   //=========================================================
   // RULE 2: EMA ALIGNMENT (35 points max)
   // - Price position relative to EMAs
   // - More generous: any position above gives points
   //=========================================================
   double emaPoints = 0;
   
   // Price above EMA 200 (long-term bullish)
   if(price > e200)
      emaPoints += 15;
   else if(price > e200 * 0.99)  // Within 1% below
      emaPoints += 8;
   
   // Price above EMA 50 (medium-term bullish)
   if(price > e50)
      emaPoints += 10;
   else if(price > e50 * 0.995)  // Within 0.5% below
      emaPoints += 5;
   
   // Price above EMA 20 (short-term bullish)
   if(price > e20)
      emaPoints += 5;
   
   // EMA stack bonus (20 > 50 > 200)
   if(e20 > e50 && e50 > e200)
      emaPoints += 5;
   else if(e20 > e50 || e50 > e200)  // Partial stack
      emaPoints += 2;
      
   score += emaPoints;
   
   //=========================================================
   // RULE 3: MACD MOMENTUM (30 points max)
   // - MACD above signal = bullish
   // - Rising histogram = momentum increasing
   //=========================================================
   double macdPoints = 0;
   
   double macdLine = macdMainBuffer[0];
   double signalLine = macdSignalBuffer[0];
   double histogram = macdLine - signalLine;
   double prevHistogram = macdMainBuffer[1] - macdSignalBuffer[1];
   
   // MACD above signal line
   if(macdLine > signalLine)
      macdPoints += 15;
   else if(histogram > prevHistogram)  // Histogram improving even if below
      macdPoints += 8;
   
   // Histogram direction (momentum)
   if(histogram > prevHistogram)
      macdPoints += 10;
   else if(histogram > 0)  // Still positive
      macdPoints += 5;
   
   // MACD above zero line (bullish territory)
   if(macdLine > 0)
      macdPoints += 5;
      
   score += macdPoints;
   
   return score / maxScore;  // Return as 0.0 to 1.0
}

//+------------------------------------------------------------------+
//| Detect bearish trend continuation for a symbol/timeframe         |
//| Returns confidence score 0.0 to 1.0                              |
//| SIMPLIFIED PRACTICAL RULES - 3 Core Checks                       |
//+------------------------------------------------------------------+
double DetectBearishTrendContinuation(int symbolIdx, int tfIdx)
{
   if(!g_mtfData[symbolIdx].isInitialized)
      return 0.5;  // Return neutral if not initialized
      
   string symbol = g_symbols[symbolIdx].symbol;
   ENUM_TIMEFRAMES tf = g_mtfTimeframes[tfIdx];
   
   double score = 0.0;
   double maxScore = 100.0;
   
   // Buffers for indicator data
   double adxBuffer[3], adxPlusBuffer[3], adxMinusBuffer[3];
   double ema20Buffer[2], ema50Buffer[2], ema200Buffer[2];
   double macdMainBuffer[3], macdSignalBuffer[3];
   double closeBuffer[5];
   
   // Copy ADX data (buffer 0 = ADX, buffer 1 = +DI, buffer 2 = -DI)
   if(CopyBuffer(g_mtfData[symbolIdx].handles[tfIdx].adxHandle, 0, 0, 3, adxBuffer) < 3)
      return 0.5;
   if(CopyBuffer(g_mtfData[symbolIdx].handles[tfIdx].adxHandle, 1, 0, 3, adxPlusBuffer) < 3)
      return 0.5;
   if(CopyBuffer(g_mtfData[symbolIdx].handles[tfIdx].adxHandle, 2, 0, 3, adxMinusBuffer) < 3)
      return 0.5;
   if(CopyBuffer(g_mtfData[symbolIdx].handles[tfIdx].ema20Handle, 0, 0, 2, ema20Buffer) < 2)
      return 0.5;
   if(CopyBuffer(g_mtfData[symbolIdx].handles[tfIdx].ema50Handle, 0, 0, 2, ema50Buffer) < 2)
      return 0.5;
   if(CopyBuffer(g_mtfData[symbolIdx].handles[tfIdx].ema200Handle, 0, 0, 2, ema200Buffer) < 2)
      return 0.5;
   if(CopyBuffer(g_mtfData[symbolIdx].handles[tfIdx].macdHandle, 0, 0, 3, macdMainBuffer) < 3)
      return 0.5;
   if(CopyBuffer(g_mtfData[symbolIdx].handles[tfIdx].macdHandle, 1, 0, 3, macdSignalBuffer) < 3)
      return 0.5;
   if(CopyClose(symbol, tf, 0, 5, closeBuffer) < 5)
      return 0.5;
   
   // Set arrays as series
   ArraySetAsSeries(adxBuffer, true);
   ArraySetAsSeries(adxPlusBuffer, true);
   ArraySetAsSeries(adxMinusBuffer, true);
   ArraySetAsSeries(ema20Buffer, true);
   ArraySetAsSeries(ema50Buffer, true);
   ArraySetAsSeries(ema200Buffer, true);
   ArraySetAsSeries(macdMainBuffer, true);
   ArraySetAsSeries(macdSignalBuffer, true);
   ArraySetAsSeries(closeBuffer, true);
   
   double price = closeBuffer[0];
   double e200 = ema200Buffer[0];
   double e50 = ema50Buffer[0];
   double e20 = ema20Buffer[0];
   double adxVal = adxBuffer[0];
   double plusDI = adxPlusBuffer[0];
   double minusDI = adxMinusBuffer[0];
   
   //=========================================================
   // RULE 1: TREND PRESENCE (35 points max)
   // - ADX > 20 indicates a trend exists
   // - -DI > +DI indicates bearish direction
   //=========================================================
   double trendPoints = 0;
   
   // Base points if any trend exists (ADX > 15)
   if(adxVal > 15)
      trendPoints += 15;
   
   // Bonus for stronger trend
   if(adxVal > 20)
      trendPoints += 10;
   if(adxVal > 25)
      trendPoints += 5;
   
   // Direction bonus: -DI > +DI for bearish
   if(minusDI > plusDI)
      trendPoints += 5;
      
   score += trendPoints;
   
   //=========================================================
   // RULE 2: EMA ALIGNMENT (35 points max) - INVERTED
   // - Price position BELOW to EMAs = bearish
   // - More generous: any position below gives points
   //=========================================================
   double emaPoints = 0;
   
   // Price below EMA 200 (long-term bearish)
   if(price < e200)
      emaPoints += 15;
   else if(price < e200 * 1.01)  // Within 1% above
      emaPoints += 8;
   
   // Price below EMA 50 (medium-term bearish)
   if(price < e50)
      emaPoints += 10;
   else if(price < e50 * 1.005)  // Within 0.5% above
      emaPoints += 5;
   
   // Price below EMA 20 (short-term bearish)
   if(price < e20)
      emaPoints += 5;
   
   // EMA stack bonus (20 < 50 < 200) - bearish stack
   if(e20 < e50 && e50 < e200)
      emaPoints += 5;
   else if(e20 < e50 || e50 < e200)  // Partial stack
      emaPoints += 2;
      
   score += emaPoints;
   
   //=========================================================
   // RULE 3: MACD MOMENTUM (30 points max) - INVERTED
   // - MACD below signal = bearish
   // - Falling histogram = momentum increasing to downside
   //=========================================================
   double macdPoints = 0;
   
   double macdLine = macdMainBuffer[0];
   double signalLine = macdSignalBuffer[0];
   double histogram = macdLine - signalLine;
   double prevHistogram = macdMainBuffer[1] - macdSignalBuffer[1];
   
   // MACD below signal line
   if(macdLine < signalLine)
      macdPoints += 15;
   else if(histogram < prevHistogram)  // Histogram falling even if above
      macdPoints += 8;
   
   // Histogram direction (momentum) - falling = bearish
   if(histogram < prevHistogram)
      macdPoints += 10;
   else if(histogram < 0)  // Still negative
      macdPoints += 5;
   
   // MACD below zero line (bearish territory)
   if(macdLine < 0)
      macdPoints += 5;
      
   score += macdPoints;
   
   return score / maxScore;  // Return as 0.0 to 1.0
}

//+------------------------------------------------------------------+
//| Calculate combined MTF bullish score                             |
//+------------------------------------------------------------------+
double GetMTFBullishScore(int symbolIdx)
{
   if(!InpUseMTFTrendScore || !g_mtfData[symbolIdx].isInitialized)
      return 100.0;  // Return max score if feature disabled
      
   double combinedScore = 0.0;
   int keyTF = (int)InpKeyTimeframe;
   
   for(int tf = 0; tf < MTF_COUNT; tf++)
   {
      double weight = g_tfWeights[keyTF][tf];
      if(weight > 0)
      {
         double tfScore = DetectBullishTrendContinuation(symbolIdx, tf);
         combinedScore += tfScore * weight;
      }
   }
   
   return combinedScore * 100.0;  // Return as percentage
}

//+------------------------------------------------------------------+
//| Calculate combined MTF bearish score                             |
//+------------------------------------------------------------------+
double GetMTFBearishScore(int symbolIdx)
{
   if(!InpUseMTFTrendScore || !g_mtfData[symbolIdx].isInitialized)
      return 100.0;  // Return max score if feature disabled
      
   double combinedScore = 0.0;
   int keyTF = (int)InpKeyTimeframe;
   
   for(int tf = 0; tf < MTF_COUNT; tf++)
   {
      double weight = g_tfWeights[keyTF][tf];
      if(weight > 0)
      {
         double tfScore = DetectBearishTrendContinuation(symbolIdx, tf);
         combinedScore += tfScore * weight;
      }
   }
   
   return combinedScore * 100.0;  // Return as percentage
}
//+------------------------------------------------------------------+
