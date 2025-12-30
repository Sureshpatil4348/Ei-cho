#property strict

// MT4 EA — RSI Suresh Bhai strategy
// Laddered RSI entries with martingale sizing, basket TP in pips before max trades,
// and dollar TP after reaching max trades. Mirrored for buys and sells.

input ENUM_TIMEFRAMES InpRSITimeframe            = PERIOD_M5;      // Lower timeframe (LTF) to check
input ENUM_TIMEFRAMES InpRSIHigherTimeframe      = PERIOD_H1;      // Higher timeframe (HTF) filter
input int            InpRSIPeriod                = 14;             // RSI period
input double         InpRSIOverbought            = 70;             // Overbought threshold (>= triggers sell first entry)
input double         InpRSIOversold              = 30;             // Oversold threshold (<= triggers buy first entry)
input bool           InpFirstEntryWaitClose      = true;           // First entry waits for candle close
input double         InpInitialLot               = 0.01;           // Initial lot size
input double         InpMartingaleMultiplier     = 2.0;            // Multiplier for each subsequent lot
input int            InpMaxTrades                = 8;              // Max trades per side
input bool           InpAllowBothSides           = true;           // Allow both buy and sell baskets simultaneously
input int            InpMagicNumber              = 123456;         // Magic number for this EA
input string         InpOrderComment             = "RSI_Ladder_Martingale"; // Order comment tag

// ATR-based grid and TP configuration
input ENUM_TIMEFRAMES InpATRGridTimeframe        = PERIOD_H4;      // ATR timeframe for grid spacing
input int            InpATRGridPeriod            = 14;             // ATR period for grid spacing
input ENUM_TIMEFRAMES InpATRTPTimeframe          = PERIOD_H1;      // ATR timeframe for pre-max TP
input int            InpATRTPPeriod              = 14;             // ATR period for pre-max TP

// Internal state for scaling logic per side
double g_firstRSI_Buy  = 0.0;
double g_firstRSI_Sell = 0.0;
int    g_nextIdx_Buy   = 1;
int    g_nextIdx_Sell  = 1;
bool   g_zoneUsed_Buy  = false;   // true = one sequence already taken for current oversold stay
bool   g_zoneUsed_Sell = false;   // true = one sequence already taken for current overbought stay

// Constants
int    C_Slippage      = 3; // slippage in points

// ATR-based sequencing state per side (frozen at first entry)
double g_firstPrice_Buy  = 0.0;
double g_firstPrice_Sell = 0.0;
double g_atrGrid_Buy     = 0.0;
double g_atrGrid_Sell    = 0.0;
double g_atrTP_Buy       = 0.0;
double g_atrTP_Sell      = 0.0;

// Utility: pip size in price units
double PipSize()
{
	if(Digits == 3 || Digits == 5) return(10 * Point);
	return(Point);
}

// Normalize lots to broker constraints
double NormalizeLots(double lots)
{
	double minLot  = MarketInfo(Symbol(), MODE_MINLOT);
	double maxLot  = MarketInfo(Symbol(), MODE_MAXLOT);
	double lotStep = MarketInfo(Symbol(), MODE_LOTSTEP);
	if(lotStep <= 0.0) lotStep = 0.01;
	// clamp
	double clamped = MathMax(minLot, MathMin(lots, maxLot));
	// step align
	double steps = MathRound(clamped / lotStep);
	double aligned = steps * lotStep;
	// final clamp to avoid 0
	if(aligned < minLot) aligned = minLot;
	return(NormalizeDouble(aligned, 2));
}

bool IsOurOrder(int index)
{
	if(!OrderSelect(index, SELECT_BY_POS, MODE_TRADES)) return(false);
	if(OrderSymbol() != Symbol()) return(false);
	if(OrderMagicNumber() != InpMagicNumber) return(false);
	if(OrderComment() != InpOrderComment) return(false);
	return(true);
}

int CountSideOrders(bool isBuy)
{
	int count = 0;
	for(int i=OrdersTotal()-1; i>=0; i--)
	{
		if(!IsOurOrder(i)) continue;
		int type = OrderType();
		if(isBuy && type == OP_BUY) count++;
		if(!isBuy && type == OP_SELL) count++;
	}
	return(count);
}

double SideVWAP(bool isBuy, double &totalLots)
{
	double volSum = 0.0;
	double volPriceSum = 0.0;
	for(int i=OrdersTotal()-1; i>=0; i--)
	{
		if(!IsOurOrder(i)) continue;
		int type = OrderType();
		if(isBuy && type != OP_BUY) continue;
		if(!isBuy && type != OP_SELL) continue;
		double v = OrderLots();
		double p = OrderOpenPrice();
		volSum += v;
		volPriceSum += v * p;
	}
	totalLots = volSum;
	if(volSum <= 0.0) return(0.0);
	return(volPriceSum / volSum);
}

double SideTotalProfit(bool isBuy)
{
	double profit = 0.0;
	for(int i=OrdersTotal()-1; i>=0; i--)
	{
		if(!IsOurOrder(i)) continue;
		int type = OrderType();
		if(isBuy && type != OP_BUY) continue;
		if(!isBuy && type != OP_SELL) continue;
		profit += OrderProfit() + OrderSwap() + OrderCommission();
	}
	return(profit);
}

double GetLastLotForSide(bool isBuy)
{
	datetime latest = 0;
	double lastLots = 0.0;
	for(int i=OrdersTotal()-1; i>=0; i--)
	{
		if(!IsOurOrder(i)) continue;
		int type = OrderType();
		if(isBuy && type != OP_BUY) continue;
		if(!isBuy && type != OP_SELL) continue;
		if(OrderOpenTime() >= latest)
		{
			latest = OrderOpenTime();
			lastLots = OrderLots();
		}
	}
	return(lastLots);
}

double GetLatestTradePriceForSide(bool isBuy)
{
	datetime latest = 0;
	double lastPrice = 0.0;
	for(int i=OrdersTotal()-1; i>=0; i--)
	{
		if(!IsOurOrder(i)) continue;
		int type = OrderType();
		if(isBuy && type != OP_BUY) continue;
		if(!isBuy && type != OP_SELL) continue;
		if(OrderOpenTime() >= latest)
		{
			latest = OrderOpenTime();
			lastPrice = OrderOpenPrice();
		}
	}
	return(lastPrice);
}

double GetFirstTradePriceForSide(bool isBuy)
{
	datetime earliest = 0;
	double firstPrice = 0.0;
	for(int i=OrdersTotal()-1; i>=0; i--)
	{
		if(!IsOurOrder(i)) continue;
		int type = OrderType();
		if(isBuy && type != OP_BUY) continue;
		if(!isBuy && type != OP_SELL) continue;
		if(earliest == 0 || OrderOpenTime() <= earliest)
		{
			earliest = OrderOpenTime();
			firstPrice = OrderOpenPrice();
		}
	}
	return(firstPrice);
}

bool AllowedToOpenOnSide(bool isBuy)
{
	if(InpAllowBothSides) return(true);
	int otherCount = CountSideOrders(!isBuy);
	return(otherCount == 0);
}

bool PlaceMarketOrder(bool isBuy, double lots)
{
	RefreshRates();
	double price = isBuy ? Ask : Bid;
	int type = isBuy ? OP_BUY : OP_SELL;
	double lotsNorm = NormalizeLots(lots);
	int ticket = OrderSend(Symbol(), type, lotsNorm, price, C_Slippage, 0, 0, InpOrderComment, InpMagicNumber, 0, isBuy ? clrBlue : clrRed);
	return(ticket > 0);
}

bool CloseAllSideOrders(bool isBuy)
{
	bool allOk = true;
	for(int pass=0; pass<3; pass++)
	{
		RefreshRates();
		bool any = false;
		for(int i=OrdersTotal()-1; i>=0; i--)
		{
			if(!IsOurOrder(i)) continue;
			int type = OrderType();
			if(isBuy && type != OP_BUY) continue;
			if(!isBuy && type != OP_SELL) continue;
			any = true;
			double price = (type == OP_BUY) ? Bid : Ask;
			if(!OrderClose(OrderTicket(), OrderLots(), price, C_Slippage, clrWhite))
				allOk = false;
		}
		if(!any) break;
	}
	return(allOk);
}

void ResetSideState(bool isBuy)
{
	if(isBuy)
	{
		g_firstRSI_Buy = 0.0;
		g_nextIdx_Buy = 1;
		g_firstPrice_Buy = 0.0;
		g_atrGrid_Buy = 0.0;
		g_atrTP_Buy = 0.0;
	}
	else
	{
		g_firstRSI_Sell = 0.0;
		g_nextIdx_Sell = 1;
		g_firstPrice_Sell = 0.0;
		g_atrGrid_Sell = 0.0;
		g_atrTP_Sell = 0.0;
	}
}

int OnInit()
{
	ResetSideState(true);
	ResetSideState(false);
	g_zoneUsed_Buy = false;
	g_zoneUsed_Sell = false;
	return(INIT_SUCCEEDED);
}

void OnDeinit(const int reason)
{
    // Clean up chart objects
    ObjectDelete(0, "RSI_SB_BUY_TP");
    ObjectDelete(0, "RSI_SB_SELL_TP");
}

void TryOpenFirstEntries(
    double rsiClosedLTF,
    double rsiCurrentLTF,
    double rsiClosedHTF,
    double rsiCurrentHTF)
{
	// SELL first entry: require both HTF and LTF to be overbought
	if(CountSideOrders(false) == 0 && AllowedToOpenOnSide(false) && !g_zoneUsed_Sell)
	{
		bool condSellLTF = InpFirstEntryWaitClose ? (rsiClosedLTF >= InpRSIOverbought) : (rsiCurrentLTF >= InpRSIOverbought);
		bool condSellHTF = InpFirstEntryWaitClose ? (rsiClosedHTF >= InpRSIOverbought) : (rsiCurrentHTF >= InpRSIOverbought);
		if(condSellLTF && condSellHTF)
		{
			if(PlaceMarketOrder(false, InpInitialLot))
			{
				g_firstRSI_Sell = InpFirstEntryWaitClose ? rsiClosedLTF : rsiCurrentLTF;
				g_nextIdx_Sell  = 1;
				g_zoneUsed_Sell = true;
				// Freeze ATRs and anchor price at first SELL entry
				g_firstPrice_Sell = GetFirstTradePriceForSide(false);
				g_atrGrid_Sell   = iATR(Symbol(), InpATRGridTimeframe, InpATRGridPeriod, 1);
				g_atrTP_Sell     = iATR(Symbol(), InpATRTPTimeframe,  InpATRTPPeriod, 1);
			}
		}
	}

	// BUY first entry: require both HTF and LTF to be oversold
	if(CountSideOrders(true) == 0 && AllowedToOpenOnSide(true) && !g_zoneUsed_Buy)
	{
		bool condBuyLTF = InpFirstEntryWaitClose ? (rsiClosedLTF <= InpRSIOversold) : (rsiCurrentLTF <= InpRSIOversold);
		bool condBuyHTF = InpFirstEntryWaitClose ? (rsiClosedHTF <= InpRSIOversold) : (rsiCurrentHTF <= InpRSIOversold);
		if(condBuyLTF && condBuyHTF)
		{
			if(PlaceMarketOrder(true, InpInitialLot))
			{
				g_firstRSI_Buy = InpFirstEntryWaitClose ? rsiClosedLTF : rsiCurrentLTF;
				g_nextIdx_Buy  = 1;
				g_zoneUsed_Buy = true;
				// Freeze ATRs and anchor price at first BUY entry
				g_firstPrice_Buy = GetFirstTradePriceForSide(true);
				g_atrGrid_Buy   = iATR(Symbol(), InpATRGridTimeframe, InpATRGridPeriod, 1);
				g_atrTP_Buy     = iATR(Symbol(), InpATRTPTimeframe,  InpATRTPPeriod, 1);
			}
		}
	}
}

void TryScaleEntries(double rsiCurrent)
{
	// SELL scaling using ATR-based grid spacing from first entry price
	int sellCount = CountSideOrders(false);
	if(sellCount > 0 && sellCount < InpMaxTrades && g_firstPrice_Sell > 0.0 && g_atrGrid_Sell > 0.0)
	{
		double latestTradePrice = GetLatestTradePriceForSide(false);
		RefreshRates();
		double currentAsk = Ask;
		double targetPrice = g_firstPrice_Sell + g_nextIdx_Sell * g_atrGrid_Sell;

		while(currentAsk >= targetPrice && currentAsk > latestTradePrice && sellCount < InpMaxTrades)
		{
			double lastLot = GetLastLotForSide(false);
			double nextLot = (lastLot > 0.0 ? lastLot : InpInitialLot) * InpMartingaleMultiplier;
			if(PlaceMarketOrder(false, nextLot))
			{
				sellCount++;
				g_nextIdx_Sell++;
				targetPrice = g_firstPrice_Sell + g_nextIdx_Sell * g_atrGrid_Sell;
				latestTradePrice = GetLatestTradePriceForSide(false);
			}
			else
			{
				break;
			}
		}
	}

	// BUY scaling using ATR-based grid spacing from first entry price
	int buyCount = CountSideOrders(true);
	if(buyCount > 0 && buyCount < InpMaxTrades && g_firstPrice_Buy > 0.0 && g_atrGrid_Buy > 0.0)
	{
		double latestTradePrice = GetLatestTradePriceForSide(true);
		RefreshRates();
		double currentBid = Bid;
		double targetPriceB = g_firstPrice_Buy - g_nextIdx_Buy * g_atrGrid_Buy;

		while(currentBid <= targetPriceB && currentBid < latestTradePrice && buyCount < InpMaxTrades)
		{
			double lastLot = GetLastLotForSide(true);
			double nextLot = (lastLot > 0.0 ? lastLot : InpInitialLot) * InpMartingaleMultiplier;
			if(PlaceMarketOrder(true, nextLot))
			{
				buyCount++;
				g_nextIdx_Buy++;
				targetPriceB = g_firstPrice_Buy - g_nextIdx_Buy * g_atrGrid_Buy;
				latestTradePrice = GetLatestTradePriceForSide(true);
			}
			else
			{
				break;
			}
		}
	}
}

void TryCloseByTargets()
{
	// SELL basket
	int sellCount = CountSideOrders(false);
	if(sellCount > 0)
	{
		// Retrace-to-first-order exit: after grid started (>1 trades), if price crosses below first sell price
		if(sellCount > 1)
		{
			double firstSellPrice = GetFirstTradePriceForSide(false);
			if(firstSellPrice > 0.0)
			{
				RefreshRates();
				double askNow = Ask;
				if(askNow <= firstSellPrice)
				{
					CloseAllSideOrders(false);
					ResetSideState(false);
					// Set count to 0 locally to skip remaining sell-side checks this tick
					sellCount = 0;
				}
			}
		}

        if(sellCount < InpMaxTrades)
		{
			double totLots = 0.0;
			double vwap = SideVWAP(false, totLots);
			if(vwap > 0.0)
			{
				RefreshRates();
				double ask = Ask;
                // Use ATR-only TP distance when available; if not available, skip closing
                if(g_atrTP_Sell > 0.0)
                {
                    double targetMove = g_atrTP_Sell;
                    if((vwap - ask) >= targetMove)
                    {
                        CloseAllSideOrders(false);
                        ResetSideState(false);
                    }
                }
			}
		}
	}

	// BUY basket
	int buyCount = CountSideOrders(true);
	if(buyCount > 0)
	{
		// Retrace-to-first-order exit: after grid started (>1 trades), if price crosses above first buy price
		if(buyCount > 1)
		{
			double firstBuyPrice = GetFirstTradePriceForSide(true);
			if(firstBuyPrice > 0.0)
			{
				RefreshRates();
				double bidNow = Bid;
				if(bidNow >= firstBuyPrice)
				{
					CloseAllSideOrders(true);
					ResetSideState(true);
					// Set count to 0 locally to skip remaining buy-side checks this tick
					buyCount = 0;
				}
			}
		}

        if(buyCount < InpMaxTrades)
		{
			double totLotsB = 0.0;
			double vwapB = SideVWAP(true, totLotsB);
			if(vwapB > 0.0)
			{
				RefreshRates();
				double bid = Bid;
                // Use ATR-only TP distance when available; if not available, skip closing
                if(g_atrTP_Buy > 0.0)
                {
                    double targetMove = g_atrTP_Buy;
                    if((bid - vwapB) >= targetMove)
                    {
                        CloseAllSideOrders(true);
                        ResetSideState(true);
                    }
                }
			}
		}
	}
}

void OnTick()
{
	// Exit checks first
	TryCloseByTargets();

	// Read RSI values (Lower TF and Higher TF)
	double rsiClosedLTF   = iRSI(Symbol(), InpRSITimeframe,       InpRSIPeriod, PRICE_CLOSE, 1);
	double rsiCurrentLTF  = iRSI(Symbol(), InpRSITimeframe,       InpRSIPeriod, PRICE_CLOSE, 0);
	double rsiClosedHTF   = iRSI(Symbol(), InpRSIHigherTimeframe, InpRSIPeriod, PRICE_CLOSE, 1);
	double rsiCurrentHTF  = iRSI(Symbol(), InpRSIHigherTimeframe, InpRSIPeriod, PRICE_CLOSE, 0);

	// Zone exit detection to re-arm sequence permission per side
	// Re-arming is based on LTF zone exits only; HTF may remain extended
	bool inSellZone = InpFirstEntryWaitClose ? (rsiClosedLTF  >= InpRSIOverbought) : (rsiCurrentLTF >= InpRSIOverbought);
	bool inBuyZone  = InpFirstEntryWaitClose ? (rsiClosedLTF  <= InpRSIOversold)   : (rsiCurrentLTF <= InpRSIOversold);
	if(!inSellZone) g_zoneUsed_Sell = false; // exited overbought -> allow new sell sequence on next re-entry
	if(!inBuyZone)  g_zoneUsed_Buy  = false; // exited oversold   -> allow new buy sequence on next re-entry

	// First entries (per side) when no open trades on that side
	TryOpenFirstEntries(rsiClosedLTF, rsiCurrentLTF, rsiClosedHTF, rsiCurrentHTF);

	// Scaling (tick-based) while basket active and below max trades, based on LTF RSI
	TryScaleEntries(rsiCurrentLTF);

	// Draw/update TP lines for visibility on chart
	DrawTPLines();
}


void EnsureHLine(const string name, const double price, const color clr)
{
	if(ObjectFind(0, name) < 0)
	{
		ObjectCreate(0, name, OBJ_HLINE, 0, 0, price);
		ObjectSetInteger(0, name, OBJPROP_COLOR, clr);
		ObjectSetInteger(0, name, OBJPROP_STYLE, STYLE_DASH);
		ObjectSetInteger(0, name, OBJPROP_WIDTH, 1);
	}
	ObjectSetDouble(0, name, OBJPROP_PRICE, price);
}

void DeleteIfExists(const string name)
{
	if(ObjectFind(0, name) >= 0) ObjectDelete(0, name);
}

void DrawTPLines()
{
	// SELL basket TP line
	int sellCount = CountSideOrders(false);
	if(sellCount > 0)
	{
        double lots=0.0; double vwap = SideVWAP(false, lots);
        if(vwap>0.0 && g_atrTP_Sell > 0.0)
        {
            double dist = g_atrTP_Sell;
            double priceTarget = vwap - dist;
            EnsureHLine("RSI_SB_SELL_TP", priceTarget, clrRed);
        }
        else
        {
            DeleteIfExists("RSI_SB_SELL_TP");
        }
	}
	else
	{
		DeleteIfExists("RSI_SB_SELL_TP");
	}

	// BUY basket TP line
	int buyCount = CountSideOrders(true);
	if(buyCount > 0)
	{
        double lotsB=0.0; double vwapB = SideVWAP(true, lotsB);
        if(vwapB>0.0 && g_atrTP_Buy > 0.0)
        {
            double distB = g_atrTP_Buy;
            double priceTarget = vwapB + distB;
            EnsureHLine("RSI_SB_BUY_TP", priceTarget, clrBlue);
        }
        else
        {
            DeleteIfExists("RSI_SB_BUY_TP");
        }
	}
	else
	{
		DeleteIfExists("RSI_SB_BUY_TP");
	}
}


