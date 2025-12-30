diff --git a/app/indicators/atr.py b/app/indicators/atr.py
--- a/app/indicators/atr.py
+++ b/app/indicators/atr.py
@@ -0,0 +1,80 @@
+"""
+ATR (Average True Range) Indicator - Matches MT4 Logic Exactly
+
+This module implements ATR calculation that replicates the MT4 strategy
+logic exactly as defined in Gold Buy Dip.mq4 lines 397-417.
+
+MT4 Logic:
+- True Range = max(High-Low, |High-PrevClose|, |Low-PrevClose|)
+- ATR = average of True Range values over period
+- Uses candles from index 1 to period (skipping current candle 0)
+"""
+
+from typing import List
+
+def calculate_atr(candles: List, period: int) -> float:
+    """
+    Calculate ATR (Average True Range) matching MT4 logic exactly.
+    
+    Args:
+        candles: List of candle objects with high, low, close attributes
+                 (most recent candle is at the end)
+        period: Number of candles to use for ATR calculation
+        
+    Returns:
+        ATR value (0.001 if insufficient data - matches MT4 minimum)
+        
+    MT4 Reference:
+        Lines 397-417 in Gold Buy Dip.mq4
+        
+    Example:
+        candles = [candle1, candle2, candle3, ...]
+        period = 14
+        
+        For each candle i from 1 to period:
+            TR1 = High[i] - Low[i]
+            TR2 = |High[i] - Close[i+1]|
+            TR3 = |Low[i] - Close[i+1]|
+            TR = max(TR1, TR2, TR3)
+        
+        ATR = sum(TR) / period
+    """
+    # Need at least period + 2 candles for ATR calculation
+    # (period candles + 1 for previous close + 1 current candle)
+    if len(candles) < period + 2:
+        return 0.001  # Match MT4 minimum value
+    
+    true_ranges = []
+    
+    # Calculate True Range for each candle
+    # MT4: for(int i = 1; i <= actualPeriod; i++)
+    # Python: We need candles[-(i)] and candles[-(i+1)]
+    actual_period = min(period, len(candles) - 2)
+    
+    for i in range(1, actual_period + 1):
+        # Current candle: candles[-(i)]
+        # Previous candle: candles[-(i+1)]
+        current_candle = candles[-(i)]
+        prev_candle = candles[-(i + 1)]
+        
+        high = current_candle.high
+        low = current_candle.low
+        prev_close = prev_candle.close
+        
+        # Calculate three True Range components
+        tr1 = high - low
+        tr2 = abs(high - prev_close)
+        tr3 = abs(low - prev_close)
+        
+        # True Range is the maximum of the three
+        true_range = max(tr1, tr2, tr3)
+        true_ranges.append(true_range)
+    
+    # Calculate ATR as average of True Ranges
+    if not true_ranges:
+        return 0.001  # Match MT4 minimum value
+    
+    atr = sum(true_ranges) / len(true_ranges)
+    
+    return atr
+
