
import pandas as pd
import zipfile
import os
import datetime

class CamarillaScanner:
    def __init__(self):
        pass

    def load_bhav_copy(self, zip_path):
        """Loads the CSV from the ZIP file into a DataFrame."""
        try:
            with zipfile.ZipFile(zip_path, 'r') as z:
                # Find the first CSV file
                csv_files = [f for f in z.namelist() if f.lower().endswith('.csv')]
                if not csv_files:
                    raise ValueError(f"No CSV found in {zip_path}")
                
                # Read CSV
                # Parse dates manually for XpryDt (format usually dd-MMM-yyyy)
                df = pd.read_csv(z.open(csv_files[0]))
                
                # Standardize columns (strip whitespace)
                df.columns = df.columns.str.strip()
                
                # Strip string columns
                str_cols = ['TckrSymb', 'FinInstrmTp', 'XpryDt', 'OptnTp']
                for c in str_cols:
                    if c in df.columns:
                        df[c] = df[c].astype(str).str.strip()

                # Convert Expiry to datetime for sorting
                # XpryDt format example: 2026-01-29 (ISO) or 29-Jan-2026
                # Using errors='coerce' is good. Mixed format warnings can be silenced by letting pandas guess without dayfirst or specifying format.
                df['XpryDt_Date'] = pd.to_datetime(df['XpryDt'], errors='coerce')
                
                return df
        except Exception as e:
            print(f"Error loading {zip_path}: {e}")
            return None

    def calculate_camarilla(self, high, low, close):
        """Calculates Camarilla pivots."""
        r = high - low
        
        # Avoid zero range division or issues if high==low
        if r == 0:
            return {
                'H4': close, 'H3': close, 'H2': close, 'H1': close,
                'L1': close, 'L2': close, 'L3': close, 'L4': close
            }

        data = {}
        data['H4'] = close + (r * 1.1 / 2)
        data['H3'] = close + (r * 1.1 / 4)
        data['H2'] = close + (r * 1.1 / 6)
        data['H1'] = close + (r * 1.1 / 12)
        data['L1'] = close - (r * 1.1 / 12)
        data['L2'] = close - (r * 1.1 / 6)
        data['L3'] = close - (r * 1.1 / 4)
        data['L4'] = close - (r * 1.1 / 2)
        return data

    def get_atm_strike(self, spot_price, available_strikes):
        """Returns the strike price closest to the spot price."""
        if len(available_strikes) == 0:
            return None
        # Find strike with minimum absolute difference
        return min(available_strikes, key=lambda x: abs(x - spot_price))

    def process_data(self, today_file, yesterday_file):
        print(f"Processing Today: {today_file}")
        print(f"Processing Yesterday: {yesterday_file}")

        df_today = self.load_bhav_copy(today_file)
        df_yest = self.load_bhav_copy(yesterday_file)

        if df_today is None or df_yest is None:
            return None

        # 1. Processing Yesterday's Data for Lookups
        yest_opts = df_yest[df_yest['FinInstrmTp'] == 'STO'].copy()
        
        # Create a dictionary for fast lookup: (Symbol, Strike, OptType) -> Data
        # Key by (Symbol, Strike, OptType, ExpiryDateString)
        
        print("Indexing Yesterday's data...")
        yest_lookup = {}
        for idx, row in yest_opts.iterrows():
            key = (row['TckrSymb'], float(row['StrkPric']), row['OptnTp'], row['XpryDt'])
            yest_lookup[key] = {
                'Open': row['OpnPric'],
                'High': row['HghPric'],
                'Low': row['LwPric'],
                'Close': row['ClsPric']
            }

        # 2. Process Today's Data
        # Filter FUTSTK for Underlying Close
        # 2. Process Today's Data
        # Filter FUTSTK for Underlying Close
        today_futs = df_today[df_today['FinInstrmTp'] == 'STF'].copy()
        today_opts = df_today[df_today['FinInstrmTp'] == 'STO'].copy()

        results = []

        # Get unique symbols
        symbols = today_futs['TckrSymb'].unique()
        print(f"Found {len(symbols)} underlying stocks in Futures.")

        for symbol in symbols:
            # Get Futures Steps
            # 1. Get Nearest Expiry Future for this symbol
            futs_sym = today_futs[today_futs['TckrSymb'] == symbol]
            if futs_sym.empty:
                continue
            
            # Find nearest expiry
            min_expiry = futs_sym['XpryDt_Date'].min()
            nearest_fut = futs_sym[futs_sym['XpryDt_Date'] == min_expiry].iloc[0]
            
            spot_close = nearest_fut['ClsPric']
            expiry_str = nearest_fut['XpryDt'] # Use string for matching
            
            # 2. Get Options for this symbol and SAME expiry
            # Filter options
            opts_sym = today_opts[
                (today_opts['TckrSymb'] == symbol) & 
                (today_opts['XpryDt'] == expiry_str)
            ]
            
            if opts_sym.empty:
                continue

            # 3. Find ATM Strike
            # Ensure strikes are float
            available_strikes = opts_sym['StrkPric'].astype(float).unique()
            atm_strike = self.get_atm_strike(spot_close, available_strikes)
            
            if atm_strike is None:
                continue

            # 4. Get CE and PE for ATM
            for opt_type in ['CE', 'PE']:
                opt_row = opts_sym[
                    (opts_sym['StrkPric'].astype(float) == atm_strike) & 
                    (opts_sym['OptnTp'] == opt_type)
                ]
                
                if opt_row.empty:
                    continue
                
                # Take the first one (should be unique per symbol-expiry-strike-type)
                row = opt_row.iloc[0]
                
                # Lookup Yesterday
                yest_key = (symbol, float(atm_strike), opt_type, expiry_str)
                yest_data = yest_lookup.get(yest_key)
                
                yest_levels = {}
                if yest_data:
                    # Calculate Yesterday's Camarilla Levels
                    yest_levels = self.calculate_camarilla(
                        yest_data['High'], yest_data['Low'], yest_data['Close']
                    )
                today_levels = self.calculate_camarilla(
                    row['HghPric'], row['LwPric'], row['ClsPric']
                )

                # Logic: Is Inside Camarilla?
                # Condition: Today H4 < Yest H3  AND  Today L4 > Yest L3
                is_inside = False
                if yest_levels:
                    # Ensure we have the necessary keys
                    if ('H4' in today_levels and 'L4' in today_levels and 
                        'H3' in yest_levels and 'L3' in yest_levels):
                        
                        cond1 = today_levels['H4'] < yest_levels['H3']
                        cond2 = today_levels['L4'] > yest_levels['L3']
                        
                        if cond1 and cond2:
                            is_inside = True

                # Logic: Is Inside Camarilla H4/L4 (New Sheet Condition)
                # Condition: Today H4 < Yest H4  AND  Today L4 > Yest L4
                is_inside_h4_l4 = False
                if yest_levels:
                    if ('H4' in today_levels and 'L4' in today_levels and 
                        'H4' in yest_levels and 'L4' in yest_levels):
                        
                        cond_h4 = today_levels['H4'] < yest_levels['H4']
                        cond_l4 = today_levels['L4'] > yest_levels['L4']
                        
                        if cond_h4 and cond_l4:
                            is_inside_h4_l4 = True

                # Store Result
                res = {
                    'Symbol': symbol,
                    'Expiry': expiry_str,
                    'Spot_Close': spot_close,
                    'ATM_Strike': atm_strike,
                    'Option_Type': opt_type,
                    'Today_Open': row['OpnPric'],
                    'Today_High': row['HghPric'],
                    'Today_Low': row['LwPric'],
                    'Today_Close': row['ClsPric'],
                    'Today_Close': row['ClsPric'],
                    'Is_Inside_Camarilla': is_inside,
                    'Is_Inside_H4_L4': is_inside_h4_l4,
                    'Is_Higher_Value': False,
                    'Is_Lower_Value': False,
                    'OpnIntrst': row['OpnIntrst'],
                    'ChngInOpnIntrst': row['ChngInOpnIntrst'],
                    'TtlTradgVol': row['TtlTradgVol'],
                    'TtlNbOfTxsExctd': row['TtlNbOfTxsExctd']
                }

                # Logic: Higher Value Camarilla (Today L4 > Yest H4)
                if yest_levels:
                     if 'L4' in today_levels and 'H4' in yest_levels:
                         if today_levels['L4'] > yest_levels['H4']:
                             res['Is_Higher_Value'] = True

                # Logic: Lower Value Camarilla (Today H4 < Yest L4)
                if yest_levels:
                     if 'H4' in today_levels and 'L4' in yest_levels:
                         if today_levels['H4'] < yest_levels['L4']:
                             res['Is_Lower_Value'] = True
                
                # Add Today's levels
                for k, v in today_levels.items():
                    res[f'Today_{k}'] = round(v, 2)
                
                # Add Yesterday's levels
                if yest_levels:
                    for k, v in yest_levels.items():
                        res[f'Yest_{k}'] = round(v, 2)
                
                results.append(res)

        return pd.DataFrame(results)

if __name__ == "__main__":
    # Test block
    import glob
    files = sorted(glob.glob("BhavCopy*.zip"))
    if len(files) >= 2:
        scanner = CamarillaScanner()
        df = scanner.process_data(files[-1], files[-2]) # Last is today, 2nd last is yesterday
        if df is not None:
            print(df.head())
            df.to_excel("Camarilla_Test_Output.xlsx", index=False)
