//+------------------------------------------------------------------+

//|                                                     EAGLE EA.mq4  |

//|                  Copyright 2023, Algofx in <algofx.in@gmail.com> |

//|                                    http://forexeaprogrammer.com/ |

//+------------------------------------------------------------------+

#property copyright "Copyright 2023, Algofx in <algofx.in@gmail.com>"

#property link      "http://forexeaprogrammer.com/"

#property version   "1.04" // Updated version

#property strict



// Account and License Information

long accNumber = 0;

string expire_date1 = "2029.01.05";    // yyyy.mm.dd



// Input Parameters

extern double   Lot                       =  0.01;

extern double   LotIncrement              =  0.01;

extern int      LotRepeater               =  2;

extern int      Candles                   =  2;    // No of Candles if Fractals lower than current price

extern double   Buffer                    =  1;    // Buffer in Pips

extern double   MinDiff                   =  1;    // Minimum Difference between BuyStop/SellStop in Pips

extern double   MaxDiff                   =  160;  // Maximum Difference between BuyStop/SellStop in Pips

extern double   TakeProfit                =  7;

extern double   StopLoss                  =  250;

extern double   TrailingStart             =  10;

extern double   TrailingStop              =  5;  

extern int      TimeZone                  =  9;    // TimeZone GMT + x

extern int      MaxOrders                 =  7;

extern double   MaxDrawDown               =  99;   // Maximum DrawDown in Percentage

extern bool     EnableMonday              = true;          // Enable Monday Trading?

extern bool     EnableTuesday             = true;          // Enable Tuesday Trading?

extern bool     EnableWednesday           = true;       // Enable Wednesday Trading?

extern bool     EnableThursday            = true;        // Enable Thursday Trading?

extern bool     EnableFriday              = true;          // Enable Friday Trading?

extern string   CloseTime                 =  "20:30";       // Deleting Pending Order Time in HH:MM Format

extern string   TradeComment              = "EAGLE";

extern int      Magic                     =  1851;



//+------------------------------------------------------------------+

//| New Input Parameter for Trade Mode Selection                     |

//+------------------------------------------------------------------+

extern int      TradeMode                 = 0; // 0 = Both, 1 = Buy Only, 2 = Sell Only



//+------------------------------------------------------------------+

//| New Input Parameter for Recovery Take Profit                    |

//+------------------------------------------------------------------+

extern double   RecoveryTakeProfit        = 14; // Take Profit in pips for recovery trades



// New SL Trailing Features

extern int      SL_TrailingStartDuration_Hours = 6;    // After how many hours to start trailing SL

extern int      SL_TrailFrequency_Hours         = 2;    // How often (in hours) to trail SL after start

extern double   SL_TrailDistance               = 20;   // SL trail distance in pips

extern double   SL_MaxTrailDifference          = 40;   // Max SL trail difference from trade position in pips



// Internal Variables

datetime sTime = 0;

datetime cycleTime = 0;

int cyclecnt = 0;



int prevday = 0;

int startflag = 0;

int oneflag = 0;

double tempLot = 0;



double buySL = 0;

double sellSL = 0;



// State Variable to Track Recovery Mode

bool isRecoveryMode = false;



//+------------------------------------------------------------------+

//| Expert initialization function                                   |

//+------------------------------------------------------------------+

int OnInit()

  {

    // Convert input parameters from pips to points

    Buffer               = Buffer * 10;

    TakeProfit           = TakeProfit * 10;

    StopLoss             = StopLoss * 10;

    TrailingStart        = TrailingStart * 10;

    TrailingStop         = TrailingStop * 10;

    MinDiff              = MinDiff * 10;

    MaxDiff              = MaxDiff * 10;

    RecoveryTakeProfit   = RecoveryTakeProfit * 10; // Convert RecoveryTakeProfit from pips to points

    

    // Initialize tempLot to initial Lot size

    tempLot = Lot;

    

    return(INIT_SUCCEEDED);

  }

//+------------------------------------------------------------------+

//| Expert deinitialization function                                 |

//+------------------------------------------------------------------+

void OnDeinit(const int reason)

  {

    // Clean up objects if necessary

    ObjectsDeleteAll(0, OBJ_HLINE, 0, 0);

  }

//+------------------------------------------------------------------+

//| Expert tick function                                             |

//+------------------------------------------------------------------+

void OnTick()

  {

    // Validate Account Number

    if(accNumber != AccountNumber() && accNumber != 0) {

      Alert ("Invalid System.!!");

      ExpertRemove(); 

      Sleep(5000000);

      return;

    }

    

    // Check for Expiration Date

    datetime e_d1 = StrToTime(expire_date1); 

    if(TimeCurrent() > e_d1 ) {

      Alert ("System Expired.!!");

      ExpertRemove(); 

      Sleep(5000000);

      return;

    }

    

    // Check for Maximum DrawDown

    double cProfit = FindCurrentProfit();

    if( cProfit <= -(AccountBalance()*MaxDrawDown*0.01) ) {

      CloseAllTrades();

      sTime = TimeCurrent() + 25;

      Print("");

      Print("All Trades Closed due Max Drawdown Reached.. Booked Loss: "+DoubleToStr(cProfit,2));

      Alert("All Trades Closed due Max Drawdown Reached..");

      return;

    }

      

    // Reset daily flags

    if(prevday != Day()) {

      startflag = 0;

      prevday = Day();

    }

    

    // Parse CloseTime

    int CloseHour = StringToInteger(StringSubstr(CloseTime,0,2));

    int CloseMinute = StringToInteger(StringSubstr(CloseTime,3,2));

    

    // Delete Pending Orders after CloseTime

    if( (TimeHour(TimeGMT()) == CloseHour && TimeMinute(TimeGMT()) >= CloseMinute) || (TimeHour(TimeGMT()) > CloseHour) ) {

      if(TradeMode == 0 || TradeMode == 1) // Both or Buy Only

         DeleteBuyStop();

      if(TradeMode == 0 || TradeMode == 2) // Both or Sell Only

         DeleteSellStop();

    }

    

    // Get Spread

    int Spread = (int)MarketInfo(Symbol(),MODE_SPREAD);



    // Display Current Time and Cycle Count

    int hour = TimeHour(TimeGMT());

    Comment("TimeGMT(): ",TimeGMT(),"\nCurrent Cycle: ",cyclecnt);

    

    // Get Current Time Structure

    MqlDateTime tm;

    TimeToStruct(TimeCurrent(),tm);

    

    // Trading Conditions Based on TimeZone and Enabled Days

    if(hour == TimeZone && startflag == 0 && 

       ((tm.day_of_week == 1 && EnableMonday) || (tm.day_of_week == 2 && EnableTuesday) || (tm.day_of_week == 3 && EnableWednesday)

       || (tm.day_of_week == 4 && EnableThursday) || (tm.day_of_week == 5 && EnableFriday))  )  {

    

      double buyFractal = 0;

      double sellFractal = 0;

      for(int i=1;i<50;i++) {

        double upfrac = iFractals(Symbol(),0,MODE_UPPER,i);

        double dnfrac = iFractals(Symbol(),0,MODE_LOWER,i);

        

        if(upfrac != 0 && upfrac != EMPTY_VALUE && buyFractal == 0) buyFractal = upfrac;

        if(dnfrac != 0 && dnfrac != EMPTY_VALUE && sellFractal == 0) sellFractal = dnfrac;

        

        if(buyFractal != 0 && sellFractal != 0)

           break;

      }

      

      double buyPrice = buyFractal + Buffer*Point;

      double sellPrice = sellFractal - Buffer*Point;

      

      if(Ask > buyPrice)

         buyPrice = iHigh(Symbol(),0,iHighest(Symbol(),0,MODE_HIGH,Candles,1)) + Buffer*Point;

      

      if(Bid < sellPrice)

         sellPrice = iLow(Symbol(),0,iLowest(Symbol(),0,MODE_LOW,Candles,1)) - Buffer*Point;

      

      double tProfit = FindTotalProfit(sTime);

      

      // Determine Trading Mode

      if(tProfit >= 0) {

        // Exit Recovery Mode if in Recovery and profit is recovered

        if(isRecoveryMode) {

          isRecoveryMode = false;

          cyclecnt = 0;

          tempLot = Lot;

          Print("Recovery complete. Switching back to normal trading.");

        }

        

        // Normal Trading

        CloseAllTrades();

        sTime = TimeCurrent();

        

        if( (buyPrice - sellPrice) >= MinDiff*Point && (buyPrice - sellPrice) <= MaxDiff*Point ) {

          cyclecnt = 1;

          cycleTime = sTime;

          

          for(int k=0;k<MaxOrders;k++) {

            double buyTP = NormalizeDouble(buyPrice + TakeProfit*Point,Digits);

            double buySL = NormalizeDouble(buyPrice - StopLoss*Point,Digits);

         

            double sellTP = NormalizeDouble(sellPrice - TakeProfit*Point,Digits);

            double sellSL = NormalizeDouble(sellPrice + StopLoss*Point,Digits);

            

            // Validate and adjust lot sizes

            double validBuyLot = GetValidLot(Lot);

            double validSellLot = GetValidLot(Lot);

            

            // Debugging statements

            Print("Normal Buy Order: Lot=", validBuyLot, " TP=", buyTP, " SL=", buySL);

            Print("Normal Sell Order: Lot=", validSellLot, " TP=", sellTP, " SL=", sellSL);

            

            if(TradeMode == 0 || TradeMode == 1) { // Both or Buy Only

              int ticket = OrderSend(Symbol(),OP_BUYSTOP,validBuyLot,NormalizeDouble(buyPrice,Digits),Spread,buySL,buyTP,TradeComment,Magic,0,clrBlue);

              if(ticket < 0) {

                Print("Failed to send BuyStop order. Error: ", GetLastError());

              }

            }

            

            if(TradeMode == 0 || TradeMode == 2) { // Both or Sell Only

              int ticket2 = OrderSend(Symbol(),OP_SELLSTOP,validSellLot,NormalizeDouble(sellPrice,Digits),Spread,sellSL,sellTP,TradeComment,Magic,0,clrRed);

              if(ticket2 < 0) {

                Print("Failed to send SellStop order. Error: ", GetLastError());

              }

            }

          }

        }

      

      } else {

        // Enter Recovery Mode

        if(!isRecoveryMode) {

          isRecoveryMode = true;

          cyclecnt = 1;

          cycleTime = TimeCurrent();

          Print("Entering Recovery Mode.");

        }

        

        // Recovery Trading

        if( (buyPrice - sellPrice) >= MinDiff*Point && (buyPrice - sellPrice) <= MaxDiff*Point ) {

          double tLot = tempLot;

          // Check if cyclecnt is a multiple of LotRepeater

          if( MathMod(cyclecnt, LotRepeater) == 0 && cyclecnt !=0 ) {

            tLot = GetNextLot(tempLot);

            tempLot = tLot;

            Print("Lot increased to ", tLot, " at cyclecnt=", cyclecnt);

          } else {

            tLot = tempLot;

          }

          

          cyclecnt++;

          cycleTime = TimeCurrent();

        

          for(int k=0;k<MaxOrders;k++) {

            // Validate and adjust lot size

            double validRecoveryLot = GetValidLot(tLot);

            

            // Debugging statements

            Print("Recovery Buy Order: Lot=", validRecoveryLot, " TP=", RecoveryTakeProfit * Point, " SL=", NormalizeDouble(buyPrice - StopLoss * Point,Digits));

            Print("Recovery Sell Order: Lot=", validRecoveryLot, " TP=", RecoveryTakeProfit * Point, " SL=", NormalizeDouble(sellPrice + StopLoss * Point,Digits));

            

            // Use RecoveryTakeProfit for all recovery trades

            double buyTP = NormalizeDouble(buyPrice + RecoveryTakeProfit * Point,Digits);

            double buySL = NormalizeDouble(buyPrice - StopLoss * Point,Digits);

         

            double sellTP = NormalizeDouble(sellPrice - RecoveryTakeProfit * Point,Digits);

            double sellSL = NormalizeDouble(sellPrice + StopLoss * Point,Digits);

            

            if(TradeMode == 0 || TradeMode == 1) { // Both or Buy Only

              int ticket = OrderSend(Symbol(),OP_BUYSTOP,validRecoveryLot,NormalizeDouble(buyPrice,Digits),Spread,buySL,buyTP,TradeComment,Magic,0,clrBlue);

              if(ticket < 0) {

                Print("Failed to send Recovery BuyStop order. Error: ", GetLastError());

              }

            }

            

            if(TradeMode == 0 || TradeMode == 2) { // Both or Sell Only

              int ticket2 = OrderSend(Symbol(),OP_SELLSTOP,validRecoveryLot,NormalizeDouble(sellPrice,Digits),Spread,sellSL,sellTP,TradeComment,Magic,0,clrRed);

              if(ticket2 < 0) {

                Print("Failed to send Recovery SellStop order. Error: ", GetLastError());

              }

            }

          }

        }

      }

      

      startflag = 1;

      oneflag = 1;

    }

    

    // Existing Trailing Stop Logic

    if(TrailingStart != 0) 

       OriginalTrailingStops();

    

    // New SL Trailing Features

    if(SL_TrailingStartDuration_Hours != 0)

       SLTrailingStops();

     

    // Close Orders Based on SL Conditions

    if(Bid <= buySL && buySL != 0) {

      CloseTodayOrder("Buy");

      buySL = 0;

      oneflag = 1;

    }

    

    if(Ask >= sellSL && sellSL != 0) {

      CloseTodayOrder("Sell");

      sellSL = 0;

      oneflag = 1;

    }   

    

    // Draw SL Lines on Chart

    DrawLine("BuySL",buySL,clrPink);

    DrawLine("SellSL",sellSL,clrPink);

    

    // Count Orders

    int buycnt = 0;

    int sellcnt = 0;

    int buystop = 0;

    int sellstop = 0;

    for (int o=0; o<OrdersTotal(); o++) {

      if (OrderSelect(o,SELECT_BY_POS,MODE_TRADES)==true) {

          if(OrderSymbol() == Symbol() && OrderMagicNumber() == Magic && OrderType() == OP_BUY && OrderOpenTime() >= cycleTime)

            buycnt++;

        

          if(OrderSymbol() == Symbol() && OrderMagicNumber() == Magic && OrderType() == OP_SELL && OrderOpenTime() >= cycleTime)

             sellcnt++;

          

          if(OrderSymbol() == Symbol() && OrderMagicNumber() == Magic && OrderType() == OP_BUYSTOP && OrderOpenTime() >= cycleTime)

             buystop++;

        

          if(OrderSymbol() == Symbol() && OrderMagicNumber() == Magic && OrderType() == OP_SELLSTOP && OrderOpenTime() >= cycleTime)

             sellstop++;

      }

    }

    

    // Manage Pending Orders Based on Active Trades

    if(buycnt != 0 && sellcnt == 0 && sellstop > 0 && (TradeMode == 0 || TradeMode == 2)) { // Check TradeMode for SellStop

      DeleteSellStop();

    }

    

    if(sellcnt != 0 && buycnt == 0 && buystop > 0 && (TradeMode == 0 || TradeMode == 1)) { // Check TradeMode for BuyStop

      DeleteBuyStop();

    }

    

    // Close Orders If No Active Trades and Profit is Negative

    if( (buycnt + sellcnt + buystop + sellstop) == 0 && oneflag == 1) {

      double tProfit = FindTotalProfit(sTime);

      if(tProfit < 0) {

        double tempProfit = FindTProfit();

        

        int j = 3;

        while(j > 0) {

          for(int cnt=0; cnt < OrdersTotal(); cnt++) { // Corrected loop condition

              if(OrderSelect(cnt, SELECT_BY_POS, MODE_TRADES)) {

                 if(OrderSymbol() == Symbol() && OrderMagicNumber() == Magic) {

                   double tPro = (OrderProfit() + OrderCommission() + OrderSwap());

                   if(MathAbs(tPro) < tempProfit) {

                      Print("");

                      Print("tPro: ",DoubleToStr(tPro,2),"    tempProfit: ",DoubleToStr(tempProfit,2));

                      if(OrderType() == OP_BUY)

                          bool chk = OrderClose(OrderTicket(),OrderLots(),Bid,Spread,clrBlue);

                       

                      if(OrderType() == OP_SELL)

                          bool chk = OrderClose(OrderTicket(),OrderLots(),Ask,Spread,clrRed);

                          

                      tempProfit = tempProfit - MathAbs(tPro);

                      oneflag = 0; 

                   }

                 }

              }

           }

           j--;

        }

      }

    }

  }

//+------------------------------------------------------------------+

//| Function to Validate and Adjust Lot Size                         |

//+------------------------------------------------------------------+

double GetValidLot(double requestedLot) {

  double minLot = MarketInfo(Symbol(), MODE_MINLOT);

  double maxLot = MarketInfo(Symbol(), MODE_MAXLOT);

  double lotStep = MarketInfo(Symbol(), MODE_LOTSTEP);

  

  // Ensure the lot is not below the minimum

  if(requestedLot < minLot)

    requestedLot = minLot;

  

  // Ensure the lot does not exceed the maximum

  if(requestedLot > maxLot)

    requestedLot = maxLot;

  

  // Adjust the lot to be a multiple of lotStep

  double validLot = NormalizeDouble(MathFloor(requestedLot / lotStep) * lotStep, 2);

  

  // Handle edge cases where rounding down might set it below minLot

  if(validLot < minLot)

    validLot = minLot;

  

  return(validLot);

}



//+------------------------------------------------------------------+

//| Function to Determine Next Lot Size for Recovery                 |

//+------------------------------------------------------------------+

double GetNextLot(double currentLot) {

  double nextLot = currentLot;

  

  // Apply Lot Increment

  if(LotIncrement > 0)

    nextLot += LotIncrement;

  

  // Validate the next lot size

  nextLot = GetValidLot(nextLot);

  

  return(nextLot);

}

//+------------------------------------------------------------------+

//| Find Total Profit                                                |

//+------------------------------------------------------------------+

double FindTotalProfit(datetime tempTime) {

  double tProfit = 0;

  int cnt;

  

  // Calculate Profit from History

  for(cnt=OrdersHistoryTotal()-1; cnt>=0; cnt--) {  // Corrected loop start

    if(OrderSelect(cnt, SELECT_BY_POS, MODE_HISTORY)) {

      if(OrderCloseTime() > sTime && OrderSymbol() == Symbol() && OrderMagicNumber() == Magic) {

        tProfit += (OrderProfit() + OrderCommission() + OrderSwap());

      }

    }

  }

  

  // Calculate Profit from Open Trades

  for(cnt=OrdersTotal()-1; cnt>=0; cnt--) {  // Corrected loop start

    if(OrderSelect(cnt, SELECT_BY_POS, MODE_TRADES)) {

      if(OrderSymbol() == Symbol() && OrderMagicNumber() == Magic) {

        tProfit += (OrderProfit() + OrderCommission() + OrderSwap());

      }

    }

  }

  

  return (tProfit);

}

//+------------------------------------------------------------------+

//| Delete All Buy Stop Orders                                       |

//+------------------------------------------------------------------+

void DeleteBuyStop() {

  for(int cnt=OrdersTotal()-1; cnt>=0; cnt--) {

    if(OrderSelect(cnt, SELECT_BY_POS, MODE_TRADES)) {

      if(OrderSymbol() == Symbol() && OrderMagicNumber() == Magic && OrderType() == OP_BUYSTOP) {

        bool chk = OrderDelete(OrderTicket());

        if(!chk) Print("Failed to delete BuyStop order #", OrderTicket(), " Error: ", GetLastError());

      }

    }

  }

}

//+------------------------------------------------------------------+

//| Delete All Sell Stop Orders                                      |

//+------------------------------------------------------------------+

void DeleteSellStop() {

  for(int cnt=OrdersTotal()-1; cnt>=0; cnt--) {

    if(OrderSelect(cnt, SELECT_BY_POS, MODE_TRADES)) {

      if(OrderSymbol() == Symbol() && OrderMagicNumber() == Magic && OrderType() == OP_SELLSTOP) {

        bool chk = OrderDelete(OrderTicket());

        if(!chk) Print("Failed to delete SellStop order #", OrderTicket(), " Error: ", GetLastError());

      }

    }

  }

}

//+------------------------------------------------------------------+

//| Close All Open Trades                                            |

//+------------------------------------------------------------------+

void CloseAllTrades() {

  for(int cnt=OrdersTotal()-1; cnt>=0; cnt--) {

    if(OrderSelect(cnt, SELECT_BY_POS, MODE_TRADES)) {

      if(OrderSymbol() == Symbol() && OrderMagicNumber() == Magic) {

        if(OrderType() == OP_BUY) {

          bool chk = OrderClose(OrderTicket(), OrderLots(), Bid, (int)MarketInfo(Symbol(), MODE_SPREAD), clrBlue);

          if(!chk) Print("Failed to close Buy order #", OrderTicket(), " Error: ", GetLastError());

        }

        if(OrderType() == OP_SELL) {

          bool chk = OrderClose(OrderTicket(), OrderLots(), Ask, (int)MarketInfo(Symbol(), MODE_SPREAD), clrRed);

          if(!chk) Print("Failed to close Sell order #", OrderTicket(), " Error: ", GetLastError());

        }

      }

    }

  }

}

//+------------------------------------------------------------------+

//| Original Trailing Stops Function (Retains Original Logic)        |

//+------------------------------------------------------------------+

void OriginalTrailingStops() {

  for (int i=0; i < OrdersTotal(); i++) {   

    if(OrderSelect(i,SELECT_BY_POS,MODE_TRADES)==true) {

      // Filter Orders by Symbol and Magic Number

      if(OrderSymbol() != Symbol() || OrderMagicNumber() != Magic)

        continue;

      

      if(OrderType() == OP_BUY) {

        double profit = (Bid - OrderOpenPrice()) / Point;

        if(profit >= TrailingStart) {

          double newSL = SlLastBar(1, Bid, 1, TrailingStop);

          if(NormalizeDouble(newSL,Digits) > NormalizeDouble(buySL,Digits) && newSL > OrderOpenPrice()) {

            buySL = NormalizeDouble(newSL,Digits);

            bool modified = OrderModify(OrderTicket(), OrderOpenPrice(), buySL, OrderTakeProfit(), 0, clrYellow);

            if(modified) {

              Print("Buy order #", OrderTicket(), " SL moved to ", DoubleToStr(buySL,Digits));

            }

            else {

              Print("Failed to modify Buy order #", OrderTicket(), " Error: ", GetLastError());

            }

          }

        }

      }

      

      if(OrderType() == OP_SELL) {

        double profit = (OrderOpenPrice() - Ask) / Point;

        if(profit >= TrailingStart) {

          double newSL = SlLastBar(-1, Ask, 1, TrailingStop);

          if( (NormalizeDouble(newSL,Digits) < NormalizeDouble(sellSL,Digits) || sellSL == 0) && newSL < OrderOpenPrice()) {

            sellSL = NormalizeDouble(newSL,Digits);

            bool modified = OrderModify(OrderTicket(), OrderOpenPrice(), sellSL, OrderTakeProfit(), 0, clrYellow);

            if(modified) {

              Print("Sell order #", OrderTicket(), " SL moved to ", DoubleToStr(sellSL,Digits));

            }

            else {

              Print("Failed to modify Sell order #", OrderTicket(), " Error: ", GetLastError());

            }

          }

        }

      }

    }   

  }

}

//+------------------------------------------------------------------+

//| Trailing Stops with New SL Trailing Features                     |

//+------------------------------------------------------------------+

void SLTrailingStops() {

  // Convert Trailing Durations from Hours to Seconds

  int trailingStartDurationSec = SL_TrailingStartDuration_Hours * 3600;

  int trailFrequencySec         = SL_TrailFrequency_Hours * 3600;

  double trailDistancePoints    = SL_TrailDistance * Point;

  double maxTrailDifferencePts  = SL_MaxTrailDifference * Point;

  

  // Iterate Through All Open Orders

  for(int i=0; i < OrdersTotal(); i++) {   

    if(OrderSelect(i,SELECT_BY_POS,MODE_TRADES)) {

      // Filter Orders by Symbol and Magic Number

      if(OrderSymbol() != Symbol() || OrderMagicNumber() != Magic)

        continue;

      

      datetime openTime = OrderOpenTime();

      double openPrice  = OrderOpenPrice();

      double currentSL  = OrderStopLoss();

      double newSL      = currentSL;

      

      // Calculate Time Elapsed Since Order Opened

      datetime timeElapsed = TimeCurrent() - openTime;

      

      // Check if Trailing Should Start

      if(timeElapsed >= trailingStartDurationSec) {

        // Calculate Number of Trailing Steps

        int number_of_trails = MathFloor( (timeElapsed - trailingStartDurationSec) / trailFrequencySec );

        

        // Ensure at least one trailing step

        if(number_of_trails < 1)

          number_of_trails = 1;

        

        // Calculate Desired SL Based on Number of Trails

        double desired_SL;

        double original_SL;

        double min_SL;

        

        if(OrderType() == OP_BUY) {

          original_SL = openPrice - StopLoss * Point;

          min_SL      = openPrice - SL_MaxTrailDifference * Point;

          

          desired_SL = original_SL + (number_of_trails * trailDistancePoints);

          desired_SL = MathMin(desired_SL, openPrice - maxTrailDifferencePts);

          

          // Update SL Only If Desired SL is Higher Than Current SL

          if(desired_SL > currentSL && desired_SL <= openPrice - maxTrailDifferencePts) {

            newSL = desired_SL;

          }

        }

        

        if(OrderType() == OP_SELL) {

          original_SL = openPrice + StopLoss * Point;

          min_SL      = openPrice + SL_MaxTrailDifference * Point;

          

          desired_SL = original_SL - (number_of_trails * trailDistancePoints);

          desired_SL = MathMax(desired_SL, openPrice + maxTrailDifferencePts);

          

          // Update SL Only If Desired SL is Lower Than Current SL

          if(desired_SL < currentSL && desired_SL >= openPrice + maxTrailDifferencePts) {

            newSL = desired_SL;

          }

        }

        

        // Modify Order If SL Needs to be Trailed

        if(newSL != currentSL) {

          bool modified = OrderModify(OrderTicket(), OrderOpenPrice(), newSL, OrderTakeProfit(), 0, clrYellow);

          if(modified) {

            if(OrderType() == OP_BUY)

              Print("Buy order #", OrderTicket(), " SL moved to ", DoubleToStr(newSL,Digits));

            if(OrderType() == OP_SELL)

              Print("Sell order #", OrderTicket(), " SL moved to ", DoubleToStr(newSL,Digits));

          }

          else {

            Print("Failed to modify order #", OrderTicket(), " Error: ", GetLastError());

          }

        }

      }

    }

  }

}

//+------------------------------------------------------------------+

//| Calculate Stop Loss for Last Bar                                 |

//+------------------------------------------------------------------+

double SlLastBar(int tip,double price, int tipFr, double tral)

{

  double fr = 0;

  if (tral!=0)

  {

    if (tip==1) fr = Bid - tral*Point;  

    else fr = Ask + tral*Point;  

  }

  

  return(fr);

}

//+------------------------------------------------------------------+

//| Find Total Profit from History and Open Trades                   |

//+------------------------------------------------------------------+

double FindTProfit() {

  double tProfit = 0;

  int cnt;

  

  // Calculate Profit from History

  for(cnt=OrdersHistoryTotal()-1; cnt>=0; cnt--) {  // Corrected loop start

    if(OrderSelect(cnt, SELECT_BY_POS, MODE_HISTORY)) {

      if(OrderCloseTime() > sTime && OrderSymbol() == Symbol() && OrderMagicNumber() == Magic) {

        tProfit += (OrderProfit() + OrderCommission() + OrderSwap());

      }

    }

  }

  return (tProfit);

}

//+------------------------------------------------------------------+

//| Close Orders Based on Type and SL Conditions                     |

//+------------------------------------------------------------------+

void CLOSEORDER(string ord) {

  for (int k=0; k<OrdersTotal(); k++) {                                               

    if (OrderSelect(k,SELECT_BY_POS,MODE_TRADES)==true) {

      RefreshRates();

      if (OrderType()==OP_BUY && OrderSymbol()==Symbol() && ord=="Buy" && Magic ==OrderMagicNumber() )

        bool chk = OrderClose(OrderTicket(),OrderLots(),Bid,(int)MarketInfo(Symbol(),MODE_SPREAD),clrBlue);

      if (OrderType()==OP_SELL && OrderSymbol()==Symbol() && ord=="Sell" && Magic ==OrderMagicNumber() )

        bool chk = OrderClose(OrderTicket(),OrderLots(),Ask,(int)MarketInfo(Symbol(),MODE_SPREAD),clrRed);

    }   

  }

}

//+------------------------------------------------------------------+

//| Find Current Profit of Open Trades                               |

//+------------------------------------------------------------------+

double FindCurrentProfit() {

  double tProfit = 0;

  for (int k=0; k<OrdersTotal(); k++) {                                               

    if (OrderSelect(k, SELECT_BY_POS, MODE_TRADES)==true) {

      RefreshRates();

      if (OrderSymbol()==Symbol() && Magic ==OrderMagicNumber() )

        tProfit += OrderProfit() + OrderCommission() + OrderSwap();

    }   

  }

  return (tProfit);

}

//+------------------------------------------------------------------+

//| Close Today's Orders Based on Trailing Conditions                |

//+------------------------------------------------------------------+

void CloseTodayOrder(string type) {

  for (int o=0; o<OrdersTotal(); o++) {

    if (OrderSelect(o,SELECT_BY_POS,MODE_TRADES)==true) {

      if(OrderSymbol() == Symbol() && OrderMagicNumber() == Magic && OrderType() == OP_BUY 

         && Bid >= (OrderOpenPrice()+(TrailingStart-TrailingStop)*Point) && type == "Buy")

        bool chk = OrderClose(OrderTicket(),OrderLots(),Bid,(int)MarketInfo(Symbol(),MODE_SPREAD),clrBlue);

      

      if(OrderSymbol() == Symbol() && OrderMagicNumber() == Magic && OrderType() == OP_SELL 

         && Ask <= (OrderOpenPrice()-(TrailingStart-TrailingStop)*Point) && type == "Sell")

        bool chk = OrderClose(OrderTicket(),OrderLots(),Ask,(int)MarketInfo(Symbol(),MODE_SPREAD),clrRed);

    }

  }

}

//+------------------------------------------------------------------+

//| Draw Horizontal Line on Chart                                    |

//+------------------------------------------------------------------+

void DrawLine(string name,double price,color clr)

  {

    if(ObjectFind(0,name) < 0)

       ObjectCreate(0,name, OBJ_HLINE, 0, TimeCurrent(), price);

    

    ObjectSet(name,OBJPROP_PRICE1,price);

    ObjectSet(name,OBJPROP_COLOR,clr);

    ObjectSet(name,OBJPROP_WIDTH,2);

    ObjectSet(name,OBJPROP_STYLE,STYLE_SOLID);

  }

//+------------------------------------------------------------------+

