from typing import List, Optional, Dict
from datetime import datetime
from decimal import Decimal
from sqlalchemy.orm import Session
from app.models.trading_models import MarketData, TradeSignal, TradeDirection, SetupState
from app.models.strategy_models import GoldBuyDipConfig, GoldBuyDipState
from app.indicators.zscore import calculate_zscore
from app.indicators.atr import calculate_atr
from app.utilities.forex_logger import forex_logger
from app.services.strategy_performance_tracker import StrategyPerformanceTracker
from app.services.base_strategy import BaseStrategy
from app.services.mt5_margin_validator import MT5MarginValidator

logger = forex_logger.get_logger(__name__)

class GoldBuyDipStrategy(BaseStrategy):
    def __init__(self, config: GoldBuyDipConfig, pair: str, timeframe: str = "15M", db: Session = None):
        super().__init__(pair, timeframe, "gold_buy_dip", db)
        self.config = config
        self.state = GoldBuyDipState()
        self.candles: List[MarketData] = []
        self.performance_tracker = StrategyPerformanceTracker(timeframe)
        self.margin_validator = MT5MarginValidator()
        self.is_gold = "XAU" in pair
    
    def add_candle(self, candle: MarketData):
        self.candles.append(candle)
        max_needed = max(self.config.lookback_candles, self.config.zscore_period, self.config.atr_period) + 10
        if len(self.candles) > max_needed:
            self.candles = self.candles[-max_needed:]
    
    def check_percentage_trigger(self) -> Optional[TradeDirection]:
        # Use previous candle and exclude current from lookback range
        # Ensure enough candles for lookback calculation
        min_candles_required = self.config.lookback_candles + 2
        if len(self.candles) < min_candles_required:
            return None

        lookback_count = self.config.lookback_candles
        # Use previous candle as current price, exclude last 2 candles from range
        current_price = self.candles[-2].close
        recent_candles = self.candles[-(lookback_count + 2):-2]

        # Calculate highest_high and lowest_low in a single loop for efficiency
        highest_high = float('-inf')
        lowest_low = float('inf')
        for c in recent_candles:
            close = c.close
            if close > highest_high:
                highest_high = close
            if close < lowest_low:
                lowest_low = close
        
        # Prevent division by zero
        if lowest_low <= 0:
            logger.error(f"Invalid lowest_low value: {lowest_low}. Cannot calculate percentage trigger.")
            return None
        if highest_high <= 0:
            logger.error(f"Invalid highest_high value: {highest_high}. Cannot calculate percentage trigger.")
            return None
        
        pct_from_low = ((current_price - lowest_low) / lowest_low) * 100
        if pct_from_low >= self.config.percentage_threshold:
            return TradeDirection.SELL
        
        # Use negative threshold check for buy signals
        pct_from_high = ((current_price - highest_high) / highest_high) * 100
        if pct_from_high <= -self.config.percentage_threshold:
            return TradeDirection.BUY
        
        return None
    
    def check_zscore_confirmation(self) -> bool:
        if len(self.candles) < self.config.zscore_period + 1:
            return False
        
        closes = [c.close for c in self.candles[-(self.config.zscore_period + 1):]]
        zscore = calculate_zscore(closes, self.config.zscore_period)
        
        if self.state.trigger_direction == TradeDirection.SELL:
            return zscore >= self.config.zscore_threshold_sell
        elif self.state.trigger_direction == TradeDirection.BUY:
            return zscore <= self.config.zscore_threshold_buy
        
        return False
    
    def calculate_grid_spacing(self) -> float:
        """Calculate grid spacing based on user configuration"""
        if self.config.use_grid_percent:
            # No fallback to market price, fail safely
            if not self.state.grid_trades:
                logger.error("Cannot calculate spacing without grid trades")
                return 0.0
            reference_price = self.state.grid_trades[-1]["price"]
            spacing = reference_price * (self.config.grid_percent / 100)
            return spacing
        else:
            atr = calculate_atr(self.candles, self.config.atr_period)
            # Return minimum ATR safety value if insufficient data (matches MT4)
            if atr <= 0:
                return 0.001
            return atr * self.config.grid_atr_multiplier
    
    def calculate_grid_lot_size(self, grid_level: int) -> float:
        """Calculate lot size for grid trade based on user configuration"""
        if self.config.use_progressive_lots:
            # Cache progressive lot sizes for efficiency
            if not hasattr(self, '_progressive_lot_cache'):
                self._progressive_lot_cache = {}
            if grid_level not in self._progressive_lot_cache:
                self._progressive_lot_cache[grid_level] = self.config.lot_size * (self.config.lot_progression_factor ** grid_level)
            return self._progressive_lot_cache[grid_level]
        else:
            return self.config.lot_size * self.config.grid_lot_multiplier
    
    def calculate_volume_weighted_take_profit(self) -> Optional[float]:
        """Calculate take profit based on volume-weighted average price of all grid trades"""
        if not self.state.grid_trades:
            return None
        
        total_lots = sum(trade["lot_size"] for trade in self.state.grid_trades)
        if total_lots == 0:
            return None
        
        # Volume-weighted average price calculation
        weighted_price = sum(trade["price"] * trade["lot_size"] for trade in self.state.grid_trades)
        vwap = weighted_price / total_lots
        
        first_trade = self.state.grid_trades[0]

        # Use Point value like MT4 (0.01 for gold, 0.0001 for forex)
        point = 0.01 if self.is_gold else 0.0001

        if first_trade["direction"] == "BUY":
            if self.config.use_take_profit_percent:
                return vwap * (1 + self.config.take_profit_percent / 100)
            else:
                # Convert points to price using Point value (matches MT4)
                return vwap + (self.config.take_profit * point)
        else:
            if self.config.use_take_profit_percent:
                return vwap * (1 - self.config.take_profit_percent / 100)
            else:
                # Convert points to price using Point value (matches MT4)
                return vwap - (self.config.take_profit * point)
    
    def check_grid_exit_conditions(self, current_price: float) -> bool:
        """Check if grid should be closed based on profit target only."""
        if not self.state.grid_trades:
            return False
        
        # Check volume-weighted take profit
        avg_tp = self.calculate_volume_weighted_take_profit()
        if avg_tp is not None:
            first_trade = self.state.grid_trades[0]
            if first_trade["direction"] == "BUY" and current_price >= avg_tp:
                logger.info(f"BUY grid profit target reached: {current_price:.2f} >= {avg_tp:.2f}")
                return True
            elif first_trade["direction"] == "SELL" and current_price <= avg_tp:
                logger.info(f"SELL grid profit target reached: {current_price:.2f} <= {avg_tp:.2f}")
                return True
        
        return False
    
    def _check_strategy_drawdown(self, current_equity: float) -> bool:
        """Check if strategy-specific maximum drawdown is exceeded"""
        if self.state.initial_balance <= 0:
            logger.error("Initial balance not set - cannot check drawdown")
            return False
        
        # Calculate drawdown based on account equity change
        equity_drawdown_pct = ((self.state.initial_balance - current_equity) / self.state.initial_balance) * 100
        
        # Only trigger on losses (positive drawdown), not profits (negative drawdown)
        if equity_drawdown_pct >= self.config.max_drawdown_percent:
            logger.critical(f"STRATEGY DRAWDOWN LIMIT EXCEEDED: {equity_drawdown_pct:.2f}% >= {self.config.max_drawdown_percent}%")
            return True
        
        return False
    
    def _process_market_data(self, candle: MarketData, current_equity: float = None) -> Optional[TradeSignal]:
        """Strategy-specific market data processing."""
        self.add_candle(candle)
        
        # Debug logging
        logger.info(f"Gold Buy Dip: Candles={len(self.candles)}, State={self.state.setup_state}, Price={candle.close:.2f}")
        
        # Calculate indicators for logging
        zscore = 0
        atr = 0
        price_movement_score = 0
        
        if len(self.candles) >= self.config.zscore_period + 1:
            closes = [c.close for c in self.candles[-(self.config.zscore_period + 1):]]
            zscore = calculate_zscore(closes, self.config.zscore_period)
        
        if len(self.candles) >= self.config.atr_period:
            atr = calculate_atr(self.candles, self.config.atr_period)
        
        # Use configured candles for price movement calculation
        if len(self.candles) >= self.config.lookback_candles + 2:
            # Use same logic as check_percentage_trigger for consistency
            current_price = self.candles[-2].close
            recent_candles = self.candles[-(self.config.lookback_candles + 2):-2]
            # Performance: Calculate highest_high and lowest_low in a single loop
            highest_high = float('-inf')
            lowest_low = float('inf')
            for c in recent_candles:
                close = c.close
                if close > highest_high:
                    highest_high = close
                if close < lowest_low:
                    lowest_low = close
            price_range = highest_high - lowest_low
            if price_range > 0:
                price_movement_score = ((current_price - lowest_low) / price_range) * 100
            
            # Debug percentage trigger calculation
            # Prevent division by zero
            if lowest_low > 0 and highest_high > 0:
                pct_from_low = ((current_price - lowest_low) / lowest_low) * 100
                pct_from_high = ((current_price - highest_high) / highest_high) * 100
            else:
                pct_from_low = 0
                pct_from_high = 0
                logger.warning(f"Invalid price values for percentage calculation: lowest_low={lowest_low}, highest_high={highest_high}")
            
            # Debug take profit monitoring for active positions (only for grid trading)
            tp_info = ""
            if self.config.use_grid_trading and self.state.grid_trades:
                avg_tp = self.calculate_volume_weighted_take_profit()
                if avg_tp:
                    tp_info = f", TP={avg_tp:.2f}"
            
            logger.info(f"Gold Buy Dip: PctFromLow={pct_from_low:.2f}%, PctFromHigh={pct_from_high:.2f}% , Threshold=±{self.config.percentage_threshold}%{tp_info}")
        
        # Check for maximum drawdown only if positions exist
        signal = None
        if current_equity and self.state.grid_trades and self._check_strategy_drawdown(current_equity):
            self.state.setup_state = SetupState.WAITING_FOR_TRIGGER
            self.state.grid_trades.clear()
            signal = TradeSignal(
                action="CLOSE_ALL",
                lot_size=0,
                reason="Strategy maximum drawdown exceeded"
            )
            # Remove pairs trading attribute for Gold Buy Dip
            signal.__dict__.pop('symbol2_trade', None)
        
        # Check take profit in real-time for active positions
        if not signal and self.state.grid_trades:
            should_exit = False
            exit_reason = ""
            
            # Apply same exit logic for both grid and single trades
            should_exit = self.check_grid_exit_conditions(candle.close)
            if should_exit:
                # Exit reason reflects TP-only logic in check_grid_exit_conditions
                if self.config.use_grid_trading:
                    exit_reason = "Grid profit target reached"
                else:
                    exit_reason = "Single trade profit target reached"
            
            if should_exit:
                # Close all trades
                self.state.setup_state = SetupState.WAITING_FOR_TRIGGER
                total_trades = len(self.state.grid_trades)
                self.state.grid_trades.clear()
                
                logger.info(f"Position closed at {candle.close:.2f}: {exit_reason}")
                
                signal = TradeSignal(
                    action="CLOSE_ALL",
                    lot_size=0,
                    reason=exit_reason
                )
                # Remove pairs trading attribute for Gold Buy Dip
                if hasattr(signal, 'symbol2_trade'):
                    delattr(signal, 'symbol2_trade')
        
        # Log only when signals are generated
        if signal is not None:
            drawdown_pct = ((self.state.initial_balance - current_equity) / self.state.initial_balance) * 100 if current_equity and self.state.initial_balance > 0 else 0
            self._log_market_data_with_signals(signal, zscore, atr, price_movement_score, drawdown_pct)
            return signal
        
        if self.state.setup_state == SetupState.WAITING_FOR_TRIGGER:
            trigger = self.check_percentage_trigger()
            if trigger:
                logger.info(f"Gold Buy Dip: Percentage trigger detected - {trigger}")
                self.state.setup_state = SetupState.WAITING_FOR_ZSCORE
                self.state.trigger_direction = trigger
                self.state.trigger_candle = len(self.candles) - 1
                self.state.wait_candles_count = 0
        
        elif self.state.setup_state == SetupState.WAITING_FOR_ZSCORE:
            self.state.wait_candles_count += 1
            logger.info(f"Gold Buy Dip: Waiting for Z-score confirmation, wait_count={self.state.wait_candles_count}, zscore={zscore:.2f}")
            
            if self.check_zscore_confirmation():
                logger.info(f"Gold Buy Dip: Z-score confirmation received - {zscore:.2f}")
                # Capital allocation disabled - trading with MT5 balance directly
                
                # Execute initial trade
                self.state.setup_state = SetupState.TRADE_EXECUTED
                
                # Use configured lot size directly - no calculations or fallbacks
                position_size = self.config.lot_size
                
                # No individual take profit
                take_profit = None
                
                # Add ticket and open_time fields for trade tracking
                self.state.grid_trades.append({
                    "price": candle.close,
                    "direction": self.state.trigger_direction.value,
                    "lot_size": self.config.lot_size,
                    "grid_level": 0,
                    "ticket": None,
                    "open_time": candle.timestamp
                })
                
                logger.info(f"Initial grid trade: Price: {candle.close:.2f}, Direction: {self.state.trigger_direction.value}, Lot: {position_size}")
                
                signal = TradeSignal(
                    action=self.state.trigger_direction.value,
                    lot_size=self.config.lot_size,
                    take_profit=take_profit,
                    reason=f"Initial trade - Z-score confirmed (Magic: {self.config.magic_number})"
                )
                # Remove pairs trading attribute for Gold Buy Dip
                if hasattr(signal, 'symbol2_trade'):
                    delattr(signal, 'symbol2_trade')
                
                # Log the signal before returning
                self._log_market_data_with_signals(signal, zscore, atr, price_movement_score, 0)
                return signal
            
            elif self.state.wait_candles_count > self.config.zscore_wait_candles:
                # Z-score confirmation timeout - reset to waiting for trigger
                logger.info(f"Z-score confirmation timeout after {self.config.zscore_wait_candles} candles")
                self.state.setup_state = SetupState.WAITING_FOR_TRIGGER
                self.state.trigger_direction = None
        
        elif self.state.setup_state == SetupState.TRADE_EXECUTED:
            
            # Handle adding new grid trades ONLY if grid trading is enabled
            if (
                self.config.use_grid_trading and
                self.state.grid_trades and
                len(self.state.grid_trades) < self.config.max_grid_trades
            ):
                
                last_trade = self.state.grid_trades[-1]
                grid_spacing = self.calculate_grid_spacing()
                
                # Check if price moved against position by grid spacing amount
                should_add_grid = False
                price_diff = 0
                
                if last_trade["direction"] == "BUY":
                    price_diff = last_trade["price"] - candle.close
                    should_add_grid = price_diff >= grid_spacing
                elif last_trade["direction"] == "SELL":
                    price_diff = candle.close - last_trade["price"]
                    should_add_grid = price_diff >= grid_spacing
                
                if should_add_grid:
                    grid_level = len(self.state.grid_trades)
                    lot_size = self.calculate_grid_lot_size(grid_level)
                    
                    # Add ticket and open_time fields for grid trade tracking
                    self.state.grid_trades.append({
                        "price": candle.close,
                        "direction": last_trade["direction"],
                        "lot_size": lot_size,
                        "grid_level": grid_level,
                        "ticket": None,
                        "open_time": candle.timestamp
                    })
                    
                    logger.info(f"Grid trade added: Level {grid_level + 1}/{self.config.max_grid_trades}, Price: {candle.close:.2f}, Lot: {lot_size:.2f}, Spacing: {grid_spacing:.2f}, PriceDiff: {price_diff:.2f}")
                    
                    # Return signal for the grid trade that was just added
                    signal = TradeSignal(
                        action=last_trade["direction"],
                        lot_size=lot_size,
                        take_profit=None,  # No individual TP for grid trades
                        reason=f"Grid trade level {grid_level + 1}/{self.config.max_grid_trades} (Magic: {self.config.magic_number})"
                    )
                    # Remove pairs trading attribute for Gold Buy Dip
                    if hasattr(signal, 'symbol2_trade'):
                        delattr(signal, 'symbol2_trade')
                    return signal
                else:
                    # Debug why grid trade wasn't added
                    logger.debug(f"Grid trade not added: PriceDiff={price_diff:.2f} < RequiredSpacing={grid_spacing:.2f}, Direction={last_trade['direction']}, LastPrice={last_trade['price']:.2f}, CurrentPrice={candle.close:.2f}")
        
        return None
    
    def get_grid_status(self) -> dict:
        """Get current grid trading status."""
        if not self.state.grid_trades:
            return {
                "active": False
            }
        else:
            # Return active grid information
            total_lots = sum(t["lot_size"] for t in self.state.grid_trades)
            # Prevent division by zero in average price calculation
            if total_lots > 0:
                average_price = sum(t["price"] * t["lot_size"] for t in self.state.grid_trades) / total_lots
            else:
                average_price = 0
                logger.warning("Total lots is zero in get_grid_status, cannot calculate average price")
            
            return {
                "active": True,
                "grid_level": len(self.state.grid_trades),
                "grid_direction": self.state.grid_trades[0]["direction"],
                "last_grid_price": self.state.grid_trades[-1]["price"],
                "average_price": average_price,
                "total_lots": total_lots
            }
    
    def _log_market_data_with_signals(self, signal: Optional[TradeSignal], zscore: float, atr: float, 
                                    price_movement_score: float, drawdown_pct: float):
        """Log market data with signals - trade from available data."""
        # Use available candles (minimum period for Z-score)
        if len(self.candles) < self.config.zscore_period + 1:
            return
            
        # Log 50 candles with signal (show all candles like before)
        from itertools import islice
        candle_count = len(self.candles)
        if candle_count >= 50:
            # Use islice for efficient access to the last 50 elements
            recent_candles = list(islice(self.candles, candle_count - 50, candle_count))
        else:
            recent_candles = self.candles
        
        # Use simple CSV logger for signal logging
        candle_data = {
            'timestamp': recent_candles[-1].timestamp,
            'open': recent_candles[-1].open,
            'high': recent_candles[-1].high,
            'low': recent_candles[-1].low,
            'close': recent_candles[-1].close
        }
        
        indicators = {
            'zscore': zscore,
            'atr': atr,
            'price_movement_score': price_movement_score,
            'strategy_state': f"Grid:{len(self.state.grid_trades)} Drawdown:{drawdown_pct:.1f}%"
        }
        
        forex_logger.log_signal(self.timeframe, candle_data, signal, indicators)

    
    def set_initial_balance(self, balance: float):
        """Set initial balance for MaxDrawdownPercent calculation"""
        self.state.initial_balance = balance
        logger.info(f"Initial balance set: ${balance}")
    
    def get_magic_number(self) -> int:
        """Get magic number for trade identification"""
        return self.config.magic_number
    
    def reset_strategy(self):
        """Reset strategy state."""
        logger.info("Resetting strategy state")
        self.state = GoldBuyDipState()
        self.candles.clear()
        logger.info("Strategy reset complete")
    
    def update_trade_ticket(self, grid_level: int, ticket: str):
        """Update trade ticket after successful execution"""
        if not isinstance(grid_level, int) or grid_level < 0 or grid_level >= len(self.state.grid_trades):
            logger.error(f"Invalid grid_level {grid_level} for updating ticket. No trade updated.")
            return
        try:
            self.state.grid_trades[grid_level]["ticket"] = ticket
            logger.info(f"Updated grid trade {grid_level} with ticket {ticket}")
        except Exception as e:
            logger.error(f"Failed to update ticket for grid trade {grid_level}: {e}")
    
    def remove_failed_trade(self, grid_level: int):
        """Remove failed trade from grid"""
        # Check grid_level bounds before attempting removal
        if 0 <= grid_level < len(self.state.grid_trades):
            removed_trade = self.state.grid_trades.pop(grid_level)
            logger.warning(f"Removed failed trade at level {grid_level}: {removed_trade}")
        else:
            logger.error(f"Attempted to remove trade at invalid grid level {grid_level}. No trade removed.")
