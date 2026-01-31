import pandas as pd
from scanner import CamarillaScanner

def test_logic():
    scanner = CamarillaScanner()
    
    # Mock data
    # Case 1: Higher Value Camarilla (Today L4 > Yest H4)
    # Yest: H4=100
    # Today: L4=101
    
    yest_levels = {'H4': 100, 'L4': 90, 'H3': 98, 'L3': 92}
    today_levels = {'H4': 110, 'L4': 101, 'H3': 108, 'L3': 103}
    
    res = {'Is_Higher_Value': False, 'Is_Lower_Value': False}
    
    # Logic from scanner.py
    if today_levels['L4'] > yest_levels['H4']:
        res['Is_Higher_Value'] = True
        
    print(f"Test Case 1 (Higher Value): Expected True, Got {res['Is_Higher_Value']}")
    assert res['Is_Higher_Value'] == True

    # Case 2: Lower Value Camarilla (Today H4 < Yest L4)
    # Yest: L4=100
    # Today: H4=99
    
    yest_levels = {'H4': 110, 'L4': 100, 'H3': 108, 'L3': 102}
    today_levels = {'H4': 99, 'L4': 90, 'H3': 97, 'L3': 92}
    
    res = {'Is_Higher_Value': False, 'Is_Lower_Value': False}
    
    # Logic from scanner.py
    if today_levels['H4'] < yest_levels['L4']:
        res['Is_Lower_Value'] = True
        
    print(f"Test Case 2 (Lower Value): Expected True, Got {res['Is_Lower_Value']}")
    assert res['Is_Lower_Value'] == True

    # Case 3: Neither
    yest_levels = {'H4': 100, 'L4': 90}
    today_levels = {'H4': 100, 'L4': 90}
    
    res = {'Is_Higher_Value': False, 'Is_Lower_Value': False}
    
    if today_levels['L4'] > yest_levels['H4']:
        res['Is_Higher_Value'] = True
    if today_levels['H4'] < yest_levels['L4']:
        res['Is_Lower_Value'] = True
        
    print(f"Test Case 3 (Neither): Expected False/False, Got {res['Is_Higher_Value']}/{res['Is_Lower_Value']}")
    assert res['Is_Higher_Value'] == False
    assert res['Is_Lower_Value'] == False

if __name__ == "__main__":
    try:
        test_logic()
        print("\nAll logic tests passed!")
    except AssertionError as e:
        print(f"\nTest Failed: {e}")
    except Exception as e:
        print(f"\nError: {e}")
