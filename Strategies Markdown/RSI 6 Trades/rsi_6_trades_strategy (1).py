"""
Python port of the MT4 expert advisor `RSI  6 Trades.mq4`.

This module mirrors the laddered RSI martingale behaviour in Python so it can
be executed against MetaTrader 5 (MT5) via the official `MetaTrader5` package.
The strategy manages independent buy and sell baskets, scales in using ATR
spacing, and exits when price retraces to the first entry or reaches an
ATR-derived target.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

try:
    import MetaTrader5 as mt5
except ImportError as exc:
    mt5 = None
    _import_error = exc
else:
    _import_error = None

from app.services.base_strategy import BaseStrategy
from app.models.strategy_models import RSI6TradesConfig, RSI6TradesState
from app.models.trading_models import MarketData, TradeSignal, TradeDirection
from app.utilities.forex_logger import forex_logger

logger = forex_logger.get_logger(__name__)

_TIMEFRAME_ALIASES: Dict[str, str] = {
    "M1": "TIMEFRAME_M1", "M2": "TIMEFRAME_M2", "M3": "TIMEFRAME_M3", "M4": "TIMEFRAME_M4",
    "M5": "TIMEFRAME_M5", "M6": "TIMEFRAME_M6", "M10": "TIMEFRAME_M10", "M12": "TIMEFRAME_M12",
    "M15": "TIMEFRAME_M15", "M20": "TIMEFRAME_M20", "M30": "TIMEFRAME_M30",
    "H1": "TIMEFRAME_H1", "H2": "TIMEFRAME_H2", "H3": "TIMEFRAME_H3", "H4": "TIMEFRAME_H4",
    "H6": "TIMEFRAME_H6", "H8": "TIMEFRAME_H8", "H12": "TIMEFRAME_H12",
    "D1": "TIMEFRAME_D1", "W1": "TIMEFRAME_W1", "MN1": "TIMEFRAME_MN1",
}

def _resolve_timeframe(value: str | int) -> Tuple[int, str]:
    if isinstance(value, int):
        return value, str(value)
    tf = str(value).upper().strip()
    if tf.startswith("PERIOD_"):
        tf = tf.replace("PERIOD_", "")
    if mt5 is None:
        raise RuntimeError("MetaTrader5 module required") from _import_error
    attr_name = _TIMEFRAME_ALIASES.get(tf)
    if attr_name and hasattr(mt5, attr_name):
        return getattr(mt5, attr_name), tf
    raise ValueError(f"Unsupported timeframe: {value}")

def calculate_rsi_series(closes: List[float], period: int) -> List[Optional[float]]:
    length = len(closes)
    if length < period + 1:
        return []
    rsis = [None] * length
    gains = [0.0] * length
    losses = [0.0] * length
    
    for i in range(1, length):
        change = closes[i] - closes[i - 1]
        gains[i] = max(change, 0.0)
        losses[i] = max(-change, 0.0)
    
    avg_gain = sum(gains[1:period + 1]) / period
    avg_loss = sum(losses[1:period + 1]) / period
    
    def _compute_rsi(ag, al):
        if al == 0 and ag == 0: return 50.0
        if al == 0: return 100.0
        if ag == 0: return 0.0
        return 100.0 - (100.0 / (1.0 + ag / al))
    
    rsis[period] = _compute_rsi(avg_gain, avg_loss)
    
    for i in range(period + 1, length):
        avg_gain = ((avg_gain * (period - 1)) + gains[i]) / period
        avg_loss = ((avg_loss * (period - 1)) + losses[i]) / period
        rsis[i] = _compute_rsi(avg_gain, avg_loss)
    
    return rsis

def calculate_atr_series(highs: List[float], lows: List[float], closes: List[float], period: int) -> List[Optional[float]]:
    length = len(closes)
    if length < period + 1:
        return []
    atrs = [None] * length
    true_ranges = [0.0] * length
    
    for i in range(1, length):
        hl = highs[i] - lows[i]
        hc = abs(highs[i] - closes[i - 1])
        lc = abs(lows[i] - closes[i - 1])
        true_ranges[i] = max(hl, hc, lc)
    
    first_atr = sum(true_ranges[1:period + 1]) / period
    atrs[period] = first_atr
    
    for i in range(period + 1, length):
        atrs[i] = ((atrs[i - 1] or 0.0) * (period - 1) + true_ranges[i]) / period
    
    return atrs

@dataclass
class SideState:
    first_rsi: float = 0.0
    next_index: int = 1
    zone_used: bool = False
    first_price: float = 0.0
    atr_grid: float = 0.0
    atr_tp: float = 0.0
    expected_trades: int = 0

class RSI6TradesStrategy(BaseStrategy):
    """
    Python implementation of the RSI 6 Trades martingale EA.
    
    Matches the exact behavior described in RSI_6_TRADES_STRATEGY.md:
    - Dual timeframe RSI confirmation (LTF + HTF)
    - Independent buy and sell baskets
    - ATR-based grid spacing and take profit
    - Retrace-to-first exit mechanism
    - Zone permission system
    - Martingale position sizing
    """
    
    def __init__(self, config: dict, pair: str, timeframe: str, db_session=None):
        super().__init__(pair, timeframe, "RSI6Trades", db_session)
        
        # Convert dict config to RSI6TradesConfig
        self.config = RSI6TradesConfig(**config)
        self.state = RSI6TradesState()
        
        # Independent side states as per documentation
        self.buy_state = SideState()
        self.sell_state = SideState()
        
        # Strategy data storage
        self.candle_data: List[MarketData] = []
        self.max_candles = max(self.config.rsi_period, self.config.atr_grid_period, self.config.atr_tp_period) + 50
        
        # Current positions tracking (simulated for BaseStrategy interface)
        self.current_positions = []
        self._broker_positions_available = False
        
        logger.info(f"RSI 6 Trades Strategy initialized for {pair} on {timeframe}")
    
    def _process_market_data(self, candle: MarketData, current_equity: float = None) -> Optional[TradeSignal]:
        """
        Process market data following the exact flow from documentation:
        1. Check exit conditions (priority)
        2. Calculate RSI (LTF & HTF) 
        3. Update zone permissions
        4. Check first entry conditions
        5. Check scaling conditions
        """
        # Sync live positions before making decisions so we do not rely on stale local state
        self._refresh_positions_cache()
        
        # Add candle to data storage
        self.candle_data.append(candle)
        if len(self.candle_data) > self.max_candles:
            self.candle_data = self.candle_data[-self.max_candles:]
        
        # Need enough data for indicators
        min_candles = max(self.config.rsi_period, self.config.atr_grid_period, self.config.atr_tp_period) + 10
        if len(self.candle_data) < min_candles:
            return None
        
        # Calculate indicators for both timeframes
        indicators = self._collect_indicators()
        if not indicators:
            return None
        
        # 1. Check exit conditions FIRST (highest priority)
        exit_signal = self._try_close_by_targets(indicators, candle)
        if exit_signal:
            return exit_signal
        
        # 2. Update zone permissions
        self._update_zone_permissions(indicators["rsi_closed_ltf"], indicators["rsi_current_ltf"])
        
        # 3. Check first entry conditions
        entry_signal = self._try_open_first_entries(indicators, candle)
        if entry_signal:
            return entry_signal
        
        # 4. Check scaling conditions
        scale_signal = self._try_scale_entries(indicators, candle)
        if scale_signal:
            return scale_signal
        
        self.state.last_tick_time = datetime.now()
        return None
    
    def _collect_indicators(self) -> Optional[Dict[str, float]]:
        """Calculate RSI and ATR indicators for both timeframes"""
        try:
            timeframe_cache: Dict[str, Tuple[List[float], List[float], List[float]]] = {}
            
            def _get_series(tf_value: Any, required: int) -> Optional[Tuple[List[float], List[float], List[float]]]:
                key = self._normalize_timeframe_key(tf_value)
                cached = timeframe_cache.get(key)
                if cached:
                    closes = cached[0]
                    if len(closes) >= required:
                        return cached
                data = self._load_timeframe_data(tf_value, required)
                if data:
                    timeframe_cache[key] = data
                return data
            
            # Lower timeframe RSI series
            ltf_value = getattr(self.config, "rsi_timeframe", self.timeframe)
            ltf_required = max(self.config.rsi_period + 2, 20)
            ltf_series = _get_series(ltf_value, ltf_required)
            if not ltf_series:
                return None
            closes_ltf, highs_ltf, lows_ltf = ltf_series
            if len(closes_ltf) < self.config.rsi_period + 1:
                return None
            rsi_series_ltf = calculate_rsi_series(closes_ltf, self.config.rsi_period)
            if len(rsi_series_ltf) < 2 or rsi_series_ltf[-1] is None or rsi_series_ltf[-2] is None:
                return None
            
            # Higher timeframe RSI series
            htf_value = getattr(self.config, "higher_timeframe", ltf_value)
            if self._normalize_timeframe_key(htf_value) == self._normalize_timeframe_key(ltf_value):
                closes_htf = closes_ltf
            else:
                htf_required = max(self.config.rsi_period + 2, 20)
                htf_series = _get_series(htf_value, htf_required)
                if not htf_series:
                    return None
                closes_htf = htf_series[0]
            if len(closes_htf) < self.config.rsi_period + 1:
                return None
            rsi_series_htf = calculate_rsi_series(closes_htf, self.config.rsi_period)
            if len(rsi_series_htf) < 2 or rsi_series_htf[-1] is None or rsi_series_htf[-2] is None:
                return None
            
            # ATR grid series using configured timeframe
            atr_grid_value = getattr(self.config, "atr_grid_timeframe", ltf_value)
            atr_grid_required = max(self.config.atr_grid_period + 2, 20)
            if self._normalize_timeframe_key(atr_grid_value) == self._normalize_timeframe_key(ltf_value):
                atr_grid_closes, atr_grid_highs, atr_grid_lows = closes_ltf, highs_ltf, lows_ltf
            else:
                atr_grid_series = _get_series(atr_grid_value, atr_grid_required)
                if atr_grid_series:
                    atr_grid_closes, atr_grid_highs, atr_grid_lows = atr_grid_series
                else:
                    atr_grid_closes = atr_grid_highs = atr_grid_lows = []
            atr_grid = 0.0
            if atr_grid_closes and len(atr_grid_closes) >= self.config.atr_grid_period + 1:
                atr_grid_series_vals = calculate_atr_series(atr_grid_highs, atr_grid_lows, atr_grid_closes, self.config.atr_grid_period)
                if len(atr_grid_series_vals) >= 2 and atr_grid_series_vals[-2] is not None:
                    atr_grid = float(atr_grid_series_vals[-2])
            
            # ATR TP series using configured timeframe
            atr_tp_value = getattr(self.config, "atr_tp_timeframe", ltf_value)
            atr_tp_required = max(self.config.atr_tp_period + 2, 20)
            if self._normalize_timeframe_key(atr_tp_value) == self._normalize_timeframe_key(ltf_value):
                atr_tp_closes, atr_tp_highs, atr_tp_lows = closes_ltf, highs_ltf, lows_ltf
            else:
                atr_tp_series = _get_series(atr_tp_value, atr_tp_required)
                if atr_tp_series:
                    atr_tp_closes, atr_tp_highs, atr_tp_lows = atr_tp_series
                else:
                    atr_tp_closes = atr_tp_highs = atr_tp_lows = []
            atr_tp = 0.0
            if atr_tp_closes and len(atr_tp_closes) >= self.config.atr_tp_period + 1:
                atr_tp_series_vals = calculate_atr_series(atr_tp_highs, atr_tp_lows, atr_tp_closes, self.config.atr_tp_period)
                if len(atr_tp_series_vals) >= 2 and atr_tp_series_vals[-2] is not None:
                    atr_tp = float(atr_tp_series_vals[-2])
            
            return {
                "rsi_closed_ltf": float(rsi_series_ltf[-2]),
                "rsi_current_ltf": float(rsi_series_ltf[-1]),
                "rsi_closed_htf": float(rsi_series_htf[-2]),
                "rsi_current_htf": float(rsi_series_htf[-1]),
                "atr_grid": atr_grid,
                "atr_tp": atr_tp,
            }
            
        except Exception as e:
            logger.error(f"Error calculating indicators: {e}")
            return None

    def _normalize_timeframe_key(self, timeframe: Any) -> str:
        if timeframe is None:
            return str(self.timeframe).upper()
        if isinstance(timeframe, str):
            tf = timeframe.upper().strip()
            if tf.startswith("PERIOD_"):
                tf = tf.replace("PERIOD_", "")
            return tf
        return str(timeframe)

    def _load_timeframe_data(self, timeframe_value: Any, minimum_bars: int) -> Optional[Tuple[List[float], List[float], List[float]]]:
        """
        Fetch candle data for a given timeframe. Uses the locally buffered candles when
        the timeframe matches the strategy timeframe; otherwise attempts to pull from MT5.
        Returns (closes, highs, lows) or None when insufficient data is available.
        """
        required = max(minimum_bars, 2)
        tf_key = self._normalize_timeframe_key(timeframe_value)
        strategy_key = self._normalize_timeframe_key(self.timeframe)
        
        if tf_key == strategy_key:
            if not self.candle_data:
                return None
            slice_length = min(len(self.candle_data), required + 5)
            window = self.candle_data[-slice_length:]
            if len(window) < required:
                return None
            closes = [float(getattr(c, "close", 0.0)) for c in window]
            highs = [float(getattr(c, "high", 0.0)) for c in window]
            lows = [float(getattr(c, "low", 0.0)) for c in window]
            return closes, highs, lows
        
        if mt5 is None:
            logger.debug("MetaTrader5 module unavailable; cannot load timeframe %s", tf_key)
            return None
        
        symbol = getattr(self.config, "symbol", None) or self.pair
        try:
            timeframe_id, _ = _resolve_timeframe(timeframe_value)
        except Exception as exc:
            logger.error(f"Failed to resolve timeframe {timeframe_value}: {exc}")
            return None
        
        bars_to_request = required + 5
        try:
            rates = mt5.copy_rates_from_pos(symbol, timeframe_id, 0, bars_to_request)
        except Exception as exc:
            logger.error(f"Error loading rates for {symbol} ({tf_key}): {exc}")
            return None
        
        if rates is None or len(rates) < required:
            last_error = None
            if hasattr(mt5, "last_error"):
                try:
                    last_error = mt5.last_error()
                except Exception:
                    last_error = None
            logger.error(f"Insufficient MT5 data for {symbol} ({tf_key}); last_error={last_error}")
            return None
        
        rates_slice = rates[-required:]
        closes = [float(entry["close"]) for entry in rates_slice]
        highs = [float(entry["high"]) for entry in rates_slice]
        lows = [float(entry["low"]) for entry in rates_slice]
        return closes, highs, lows

    def _refresh_positions_cache(self) -> None:
        """
        Synchronise the in-memory position cache with live broker data when available.
        Falls back to the local cache if MT5 is unavailable or returns an error.
        """
        broker_positions = self._get_broker_positions()
        if broker_positions is None:
            self._broker_positions_available = False
            return
        
        self._broker_positions_available = True
        self.current_positions = broker_positions
        self._sync_side_state_from_positions(is_buy=True)
        self._sync_side_state_from_positions(is_buy=False)

    def _get_broker_positions(self) -> Optional[List[Dict[str, Any]]]:
        """Fetch open positions from MT5 filtered by symbol, comment, and magic number."""
        if mt5 is None:
            return None
        
        symbol = getattr(self.config, "symbol", None) or self.pair
        order_comment = getattr(self.config, "order_comment", None)
        magic_number = getattr(self.config, "magic_number", None)
        
        try:
            positions = mt5.positions_get(symbol=symbol)
        except Exception as exc:
            logger.error(f"Error retrieving MT5 positions for {symbol}: {exc}")
            return None
        
        if positions is None:
            last_error = None
            if hasattr(mt5, "last_error"):
                try:
                    last_error = mt5.last_error()
                except Exception:
                    last_error = None
            logger.error(f"MT5 positions_get returned None for {symbol}; last_error={last_error}")
            return None
        
        filtered: List[Dict[str, Any]] = []
        position_type_buy = getattr(mt5, "POSITION_TYPE_BUY", 0)
        for pos in positions:
            if order_comment and getattr(pos, "comment", None) != order_comment:
                continue
            if magic_number is not None and getattr(pos, "magic", None) != magic_number:
                continue
            pos_type = "BUY" if getattr(pos, "type", position_type_buy) == position_type_buy else "SELL"
            filtered.append({
                "ticket": getattr(pos, "ticket", None),
                "type": pos_type,
                "volume": float(getattr(pos, "volume", 0.0)),
                "price": float(getattr(pos, "price_open", 0.0)),
                "time": getattr(pos, "time_msc", None) or getattr(pos, "time", 0),
            })
        return filtered

    def _sync_side_state_from_positions(self, is_buy: bool) -> None:
        """Align side state with live broker positions, clearing expectations when fills fail."""
        side = self.buy_state if is_buy else self.sell_state
        target_type = "BUY" if is_buy else "SELL"
        relevant_positions = [pos for pos in self.current_positions if pos["type"] == target_type]
        count = len(relevant_positions)
        
        if count == 0:
            if side.expected_trades > 0 or side.first_price != 0.0:
                logger.debug("Broker shows no %s positions; resetting local side state", target_type)
                self._reset_side_state(is_buy)
            return
        
        side.zone_used = True
        if side.expected_trades != count:
            if side.expected_trades > count:
                logger.debug("Expected %s trades but broker shows %s on %s side", side.expected_trades, count, target_type)
            side.expected_trades = count
        side.next_index = max(count, 1)
        earliest = min(relevant_positions, key=lambda p: p.get("time") or 0)
        side.first_price = earliest["price"]
    
    def _update_zone_permissions(self, rsi_closed: float, rsi_current: float) -> None:
        """
        Update zone permissions as per documentation:
        - SELL zone resets when RSI exits overbought (< 70)
        - BUY zone resets when RSI exits oversold (> 30)
        """
        if self.config.wait_for_candle_close:
            in_sell_zone = rsi_closed >= self.config.rsi_overbought
            in_buy_zone = rsi_closed <= self.config.rsi_oversold
        else:
            in_sell_zone = rsi_current >= self.config.rsi_overbought
            in_buy_zone = rsi_current <= self.config.rsi_oversold
        
        if not in_sell_zone:
            self.sell_state.zone_used = False
        if not in_buy_zone:
            self.buy_state.zone_used = False
    
    def _try_open_first_entries(self, indicators: Dict[str, float], candle: MarketData) -> Optional[TradeSignal]:
        """
        Check first entry conditions as per documentation:
        - No existing positions on side
        - Both LTF & HTF RSI extreme
        - Zone not used
        - Direction permissions (if allow_both_directions = false)
        """
        
        # Check SELL first entry
        sell_signal = self._check_first_entry_side(indicators, candle, is_buy=False)
        if sell_signal:
            return sell_signal
        
        # Check BUY first entry
        buy_signal = self._check_first_entry_side(indicators, candle, is_buy=True)
        if buy_signal:
            return buy_signal
        
        return None
    
    def _check_first_entry_side(self, indicators: Dict[str, float], candle: MarketData, is_buy: bool) -> Optional[TradeSignal]:
        """Check first entry conditions for one side"""
        side_state = self.buy_state if is_buy else self.sell_state
        count = self._count_side_orders(is_buy)
        
        if count == 0 and not side_state.zone_used and self._allowed_to_open(is_buy):
            if self._check_rsi_conditions(indicators, is_buy):
                return self._create_first_entry_signal(indicators, candle, is_buy)
        return None
    
    def _check_rsi_conditions(self, indicators: Dict[str, float], is_buy: bool) -> bool:
        """Check RSI conditions for entry"""
        threshold = self.config.rsi_oversold if is_buy else self.config.rsi_overbought
        rsi_key = "rsi_closed_ltf" if self.config.wait_for_candle_close else "rsi_current_ltf"
        htf_key = "rsi_closed_htf" if self.config.wait_for_candle_close else "rsi_current_htf"
        
        if is_buy:
            return indicators[rsi_key] <= threshold and indicators[htf_key] <= threshold
        else:
            return indicators[rsi_key] >= threshold and indicators[htf_key] >= threshold
    
    def _create_first_entry_signal(self, indicators: Dict[str, float], candle: MarketData, is_buy: bool) -> TradeSignal:
        """Create first entry signal and initialize state"""
        side_state = self.buy_state if is_buy else self.sell_state
        action = TradeDirection.BUY if is_buy else TradeDirection.SELL
        side_name = "BUY" if is_buy else "SELL"
        
        # Initialize sequence state
        side_state.zone_used = True
        side_state.next_index = 1
        side_state.first_price = candle.close
        side_state.atr_grid = indicators["atr_grid"]
        side_state.atr_tp = indicators["atr_tp"]
        side_state.first_rsi = (indicators["rsi_closed_ltf"] if self.config.wait_for_candle_close 
                               else indicators["rsi_current_ltf"])
        side_state.expected_trades = self._count_side_orders(is_buy) + 1
        
        # Add to positions tracking
        self._add_position(side_name, self.config.initial_lot, candle.close)
        
        logger.info(f"Opened first {side_name} @ {candle.close:.5f} (ATR grid {indicators['atr_grid']:.5f}, ATR TP {indicators['atr_tp']:.5f})")
        
        confidence = abs(side_state.first_rsi - 50) / 50
        return TradeSignal(
            action=action,
            entry_price=candle.close,
            lot_size=self.config.initial_lot,
            take_profit=None,
            stop_loss=None,
            timestamp=candle.timestamp,
            confidence=confidence,
            strategy_name=self.strategy_name,
            signal_strength=confidence
        )
    
    def _try_scale_entries(self, indicators: Dict[str, float], candle: MarketData) -> Optional[TradeSignal]:
        """
        Check scaling conditions as per documentation:
        - At least 1 position exists
        - Total positions < max_trades
        - Price reached next grid level
        - Price > latest trade entry (for sells) or < latest trade entry (for buys)
        """
        
        # Check SELL scaling
        sell_count = self._count_side_orders(is_buy=False)
        if (sell_count > 0 and sell_count < self.config.max_trades and 
            self.sell_state.first_price > 0.0 and self.sell_state.atr_grid > 0.0):
            
            latest_price = self._get_latest_trade_price(is_buy=False)
            current_price = candle.close  # Using close as ASK approximation
            target_price = self.sell_state.first_price + self.sell_state.next_index * self.sell_state.atr_grid
            
            if current_price >= target_price and current_price > latest_price:
                last_volume = self._get_last_volume(is_buy=False)
                next_volume = (last_volume if last_volume > 0 else self.config.initial_lot) * self.config.martingale_multiplier
                
                # Add to positions tracking
                self._add_position("SELL", next_volume, candle.close)
                self.sell_state.next_index = sell_count + 1
                self.sell_state.expected_trades = sell_count + 1
                
                logger.info(f"Scaled SELL #{sell_count + 1} @ {candle.close:.5f} (lot {next_volume:.2f})")
                
                return TradeSignal(
                    action=TradeDirection.SELL,
                    entry_price=candle.close,
                    lot_size=next_volume,
                    take_profit=None,
                    stop_loss=None,
                    timestamp=candle.timestamp,
                    confidence=0.8,
                    strategy_name=self.strategy_name,
                    signal_strength=0.8
                )
        
        # Check BUY scaling
        buy_count = self._count_side_orders(is_buy=True)
        if (buy_count > 0 and buy_count < self.config.max_trades and 
            self.buy_state.first_price > 0.0 and self.buy_state.atr_grid > 0.0):
            
            latest_price = self._get_latest_trade_price(is_buy=True)
            current_price = candle.close  # Using close as BID approximation
            target_price = self.buy_state.first_price - self.buy_state.next_index * self.buy_state.atr_grid
            
            if current_price <= target_price and current_price < latest_price:
                last_volume = self._get_last_volume(is_buy=True)
                next_volume = (last_volume if last_volume > 0 else self.config.initial_lot) * self.config.martingale_multiplier
                
                # Add to positions tracking
                self._add_position("BUY", next_volume, candle.close)
                self.buy_state.next_index = buy_count + 1
                self.buy_state.expected_trades = buy_count + 1
                
                logger.info(f"Scaled BUY #{buy_count + 1} @ {candle.close:.5f} (lot {next_volume:.2f})")
                
                return TradeSignal(
                    action=TradeDirection.BUY,
                    entry_price=candle.close,
                    lot_size=next_volume,
                    take_profit=None,
                    stop_loss=None,
                    timestamp=candle.timestamp,
                    confidence=0.8,
                    strategy_name=self.strategy_name,
                    signal_strength=0.8
                )
        
        return None
    
    def _try_close_by_targets(self, indicators: Dict[str, float], candle: MarketData) -> Optional[TradeSignal]:
        """
        Check exit conditions as per documentation:
        1. Retrace to first entry (if positions > 1)
        2. ATR-based take profit using VWAP
        """
        
        # Check SELL positions exit
        sell_count = self._count_side_orders(is_buy=False)
        if sell_count > 0:
            current_price = candle.close  # Using close as ASK approximation
            
            # Exit condition 1: Retrace to first entry (priority exit)
            if sell_count > 1 and self.sell_state.first_price > 0.0:
                if current_price <= self.sell_state.first_price:
                    self._close_all_side_orders(is_buy=False)
                    self._reset_side_state(is_buy=False)
                    
                    logger.info(f"Closed SELL basket - retrace to first entry @ {current_price:.5f}")
                    
                    return TradeSignal(
                        action=TradeDirection.CLOSE_SELL,
                        entry_price=current_price,
                        lot_size=0.0,  # Close all
                        take_profit=None,
                        stop_loss=None,
                        timestamp=candle.timestamp,
                        confidence=1.0,
                        strategy_name=self.strategy_name,
                        signal_strength=1.0
                    )
            
            # Exit condition 2: ATR take profit using VWAP
            if sell_count < self.config.max_trades and self.sell_state.atr_tp > 0.0:
                vwap, total_volume = self._side_vwap(is_buy=False)
                if total_volume > 0.0 and (vwap - current_price) >= self.sell_state.atr_tp:
                    self._close_all_side_orders(is_buy=False)
                    self._reset_side_state(is_buy=False)
                    
                    logger.info(f"Closed SELL basket - ATR TP @ {current_price:.5f} (VWAP {vwap:.5f})")
                    
                    return TradeSignal(
                        action=TradeDirection.CLOSE_SELL,
                        entry_price=current_price,
                        lot_size=0.0,  # Close all
                        take_profit=None,
                        stop_loss=None,
                        timestamp=candle.timestamp,
                        confidence=1.0,
                        strategy_name=self.strategy_name,
                        signal_strength=1.0
                    )
        
        # Check BUY positions exit
        buy_count = self._count_side_orders(is_buy=True)
        if buy_count > 0:
            current_price = candle.close  # Using close as BID approximation
            
            # Exit condition 1: Retrace to first entry (priority exit)
            if buy_count > 1 and self.buy_state.first_price > 0.0:
                if current_price >= self.buy_state.first_price:
                    self._close_all_side_orders(is_buy=True)
                    self._reset_side_state(is_buy=True)
                    
                    logger.info(f"Closed BUY basket - retrace to first entry @ {current_price:.5f}")
                    
                    return TradeSignal(
                        action=TradeDirection.CLOSE_BUY,
                        entry_price=current_price,
                        lot_size=0.0,  # Close all
                        take_profit=None,
                        stop_loss=None,
                        timestamp=candle.timestamp,
                        confidence=1.0,
                        strategy_name=self.strategy_name,
                        signal_strength=1.0
                    )
            
            # Exit condition 2: ATR take profit using VWAP
            if buy_count < self.config.max_trades and self.buy_state.atr_tp > 0.0:
                vwap, total_volume = self._side_vwap(is_buy=True)
                if total_volume > 0.0 and (current_price - vwap) >= self.buy_state.atr_tp:
                    self._close_all_side_orders(is_buy=True)
                    self._reset_side_state(is_buy=True)
                    
                    logger.info(f"Closed BUY basket - ATR TP @ {current_price:.5f} (VWAP {vwap:.5f})")
                    
                    return TradeSignal(
                        action=TradeDirection.CLOSE_BUY,
                        entry_price=current_price,
                        lot_size=0.0,  # Close all
                        take_profit=None,
                        stop_loss=None,
                        timestamp=candle.timestamp,
                        confidence=1.0,
                        strategy_name=self.strategy_name,
                        signal_strength=1.0
                    )
        
        return None
    
    # Helper methods for position management (simulated for BaseStrategy interface)
    
    def _count_side_orders(self, is_buy: bool) -> int:
        """Count positions on one side"""
        target_type = "BUY" if is_buy else "SELL"
        return sum(1 for pos in self.current_positions if pos["type"] == target_type)
    
    def _allowed_to_open(self, is_buy: bool) -> bool:
        """Check if allowed to open new position on side"""
        if self.config.allow_both_directions:
            return True
        other_count = self._count_side_orders(not is_buy)
        return other_count == 0
    
    def _get_latest_trade_price(self, is_buy: bool) -> float:
        """Get price of most recent trade on side"""
        target_type = "BUY" if is_buy else "SELL"
        latest_time = None
        latest_price = 0.0
        for pos in self.current_positions:
            if pos["type"] != target_type:
                continue
            if latest_time is None or pos["time"] > latest_time:
                latest_time = pos["time"]
                latest_price = pos["price"]
        return latest_price
    
    def _get_last_volume(self, is_buy: bool) -> float:
        """Get volume of most recent trade on side"""
        target_type = "BUY" if is_buy else "SELL"
        latest_time = None
        last_volume = 0.0
        for pos in self.current_positions:
            if pos["type"] != target_type:
                continue
            if latest_time is None or pos["time"] > latest_time:
                latest_time = pos["time"]
                last_volume = pos["volume"]
        return last_volume
    
    def _side_vwap(self, is_buy: bool) -> Tuple[float, float]:
        """Calculate Volume-Weighted Average Price for one side"""
        target_type = "BUY" if is_buy else "SELL"
        total_volume = 0.0
        weighted_price = 0.0
        for pos in self.current_positions:
            if pos["type"] != target_type:
                continue
            vol = pos["volume"]
            total_volume += vol
            weighted_price += vol * pos["price"]
        if total_volume <= 0.0:
            return 0.0, 0.0
        return weighted_price / total_volume, total_volume
    
    def _add_position(self, position_type: str, volume: float, price: float) -> None:
        """Add position to tracking"""
        if self._broker_positions_available:
            # When broker data is available we trust live positions rather than a local cache.
            return
        self.current_positions.append({
            "type": position_type,
            "volume": volume,
            "price": price,
            "time": datetime.now().timestamp()
        })
    
    def _close_all_side_orders(self, is_buy: bool) -> bool:
        """Close all positions on one side"""
        if self._broker_positions_available:
            # Execution layer will action the CLOSE signal; nothing to do locally.
            return True
        target_type = "BUY" if is_buy else "SELL"
        self.current_positions = [pos for pos in self.current_positions if pos["type"] != target_type]
        return True
    
    def _reset_side_state(self, is_buy: bool) -> None:
        """Reset state for one side"""
        side = self.buy_state if is_buy else self.sell_state
        side.first_rsi = 0.0
        side.next_index = 1
        side.zone_used = False
        side.first_price = 0.0
        side.atr_grid = 0.0
        side.atr_tp = 0.0
        side.expected_trades = 0
    
    def reset_strategy(self):
        """Reset strategy state"""
        self.state = RSI6TradesState()
        self.buy_state = SideState()
        self.sell_state = SideState()
        self.candle_data = []
        self.current_positions = []
        logger.info("RSI 6 Trades strategy state reset")
    
    def get_status(self) -> dict:
        """Get current strategy status"""
        return {
            'name': self.strategy_name,
            'timeframe': self.timeframe,
            'pair': self.pair,
            'buy_zone_used': self.buy_state.zone_used,
            'sell_zone_used': self.sell_state.zone_used,
            'buy_trades_count': self.buy_state.next_index - 1 if self.buy_state.zone_used else 0,
            'sell_trades_count': self.sell_state.next_index - 1 if self.sell_state.zone_used else 0,
            'total_positions': len(self.current_positions),
            'buy_positions': self._count_side_orders(is_buy=True),
            'sell_positions': self._count_side_orders(is_buy=False),
            'candles_loaded': len(self.candle_data),
            'last_tick_time': self.state.last_tick_time.isoformat() if self.state.last_tick_time else None,
            'state': {
                'buy': {
                    'zone_used': self.buy_state.zone_used,
                    'first_price': self.buy_state.first_price,
                    'next_index': self.buy_state.next_index,
                    'atr_grid': self.buy_state.atr_grid,
                    'atr_tp': self.buy_state.atr_tp,
                },
                'sell': {
                    'zone_used': self.sell_state.zone_used,
                    'first_price': self.sell_state.first_price,
                    'next_index': self.sell_state.next_index,
                    'atr_grid': self.sell_state.atr_grid,
                    'atr_tp': self.sell_state.atr_tp,
                }
            },
            'config': {
                'symbol': self.config.symbol,
                'rsi_period': self.config.rsi_period,
                'rsi_overbought': self.config.rsi_overbought,
                'rsi_oversold': self.config.rsi_oversold,
                'initial_lot': self.config.initial_lot,
                'max_trades': self.config.max_trades,
                'martingale_multiplier': self.config.martingale_multiplier,
                'wait_for_candle_close': self.config.wait_for_candle_close,
                'allow_both_directions': self.config.allow_both_directions
            }
        }
