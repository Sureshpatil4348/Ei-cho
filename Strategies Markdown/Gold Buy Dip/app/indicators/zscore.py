diff --git a/app/indicators/zscore.py b/app/indicators/zscore.py
--- a/app/indicators/zscore.py
+++ b/app/indicators/zscore.py
@@ -0,0 +1,66 @@
+"""
+Z-Score Indicator - Matches MT4 Logic Exactly
+
+This module implements Z-score calculation that replicates the MT4 strategy
+logic exactly as defined in Gold Buy Dip.mq4 lines 234-265.
+
+MT4 Logic:
+- Current price: Close[0]
+- Historical prices: Close[1] to Close[period]
+- Mean = sum(historical prices) / period
+- StdDev = sqrt(sum((price - mean)^2) / period)
+- Z-score = (current_price - mean) / stdDev
+"""
+
+def calculate_zscore(closes: list, period: int) -> float:
+    """
+    Calculate Z-score matching MT4 logic exactly.
+    
+    Args:
+        closes: List of closing prices (most recent price is at the end)
+        period: Number of historical candles to use for calculation
+        
+    Returns:
+        Z-score value (0 if insufficient data or stdDev is 0)
+        
+    MT4 Reference:
+        Lines 234-265 in Gold Buy Dip.mq4
+        
+    Example:
+        closes = [100, 101, 102, 103, 104, 105]  # 6 candles
+        period = 5
+        
+        Current price: 105 (closes[-1])
+        Historical: [100, 101, 102, 103, 104] (closes[-(period+1):-1])
+        Mean: (100+101+102+103+104)/5 = 102
+        StdDev: sqrt(((100-102)^2 + ... + (104-102)^2)/5) = 1.414
+        Z-score: (105-102)/1.414 = 2.12
+    """
+    # Need at least period + 1 candles (period for history, 1 for current)
+    if len(closes) < period + 1:
+        return 0
+    
+    # Current price is the most recent close (Close[0] in MT4)
+    current_price = closes[-1]
+    
+    # Historical prices: Close[1] to Close[period] in MT4
+    # In Python list: closes[-(period+1):-1]
+    historical_closes = closes[-(period + 1):-1]
+    
+    # Calculate mean
+    mean = sum(historical_closes) / len(historical_closes)
+    
+    # Calculate standard deviation
+    sum_squares = sum((c - mean) ** 2 for c in historical_closes)
+    variance = sum_squares / len(historical_closes)
+    std_dev = variance ** 0.5
+    
+    # Prevent division by zero
+    if std_dev == 0:
+        return 0
+    
+    # Calculate Z-score
+    zscore = (current_price - mean) / std_dev
+    
+    return zscore
+
