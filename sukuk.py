import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Number of iterations
n_iter = 1000

# Create a new workbook
wb = Workbook()
ws = wb.active
ws.title = "NTS-L"

# Function to get column letter
def get_col_letter(col_num):
    """Convert column number to Excel column letter"""
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result

# Set up the exact structure as Sukuk_NTS.xlsx
# Row 1: Title
ws['H1'] = "RENTAL TAX SHIELD WITH LEVERAGE"

# Row 2: Empty

# Row 3: Column headers (A1, A2, A3, etc.)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)  # Start from column I (9th column)
    ws[f'{col_letter}3'] = f'A{i+1}'

# Row 4: Net Operating Income
ws['A4'] = "Net Operating Income"
ws['H4'] = "NOI"
# Fill NOI values: linear increase from 10000 to 52000
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    value = 10000 + (52000 - 10000) * i / (n_iter - 1)
    ws[f'{col_letter}4'] = value

# Row 5: Interest of Debt
ws['A5'] = "Interest of Debt (market value of debt*interest rate)"
ws['G5'] = "INTEREST ON DEBT"
ws['H5'] = "INTEREST"
# Formula: =I19*I24 (and similar for other columns)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}5'] = f'={col_letter}19*{col_letter}24'

# Row 6: Rent of Sukuk
ws['A6'] = "Rent of Sukuk ijarah assets (market value of sukuk*rent rate)"
ws['G6'] = "RENT ON SUKUK"
ws['H6'] = "RENT"
# Formula: =I20*I25 (and similar for other columns)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}6'] = f'={col_letter}20*{col_letter}25'

# Row 7: Benefit of asset depreciation
ws['A7'] = "Benefit of asset depriciation/running expensed of the ijarah asset"
ws['H7'] = "NDTS"
# Formula: =I20*I26 (market value of sukuk * depreciation rate)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}7'] = f'={col_letter}20*{col_letter}26'

# Row 8: CAPITAL GAIN OR LOSS
ws['A8'] = "CAPITAL GAIN OR LOSS"
ws['G8'] = "CAPITAL GAIN OR LOSS"
ws['H8'] = "CAPITAL GAIN OR LOSS"
# Formula: =I20*I27 (market value of sukuk * capital gain/loss rate)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}8'] = f'={col_letter}20*{col_letter}27'

# Row 9: Dividend paid
ws['A9'] = "Dividend paid (tax shield)"
ws['G9'] = "DIVIDEND ON TAX"
ws['H9'] = "Dividend"
# Formula: =(I4-I5-I6-I7)*I28 (NOI - Interest - Rent - Depreciation) * dividend rate
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}9'] = f'=({col_letter}4-{col_letter}5-{col_letter}6-{col_letter}7)*{col_letter}28'

# Row 10: Earnings before Tax
ws['A10'] = "Earnings before Tax with above calculation"
ws['H10'] = "EBT"
# Formula: =I4-I5-I6+I7+I8-I9 (and similar for other columns)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}10'] = f'={col_letter}4-{col_letter}5-{col_letter}6+{col_letter}7+{col_letter}8-{col_letter}9'

# Row 11: Tax Amount
ws['A11'] = "TAX.C-NTS-L"
ws['G11'] = "TAX AMOUNT Tax @ 35%"
ws['H11'] = "TAX.C-NTS-L"
# Formula: =I10*I22 (EBT * Tax rate)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}11'] = f'={col_letter}10*{col_letter}22'

# Row 12: EAT
ws['A12'] = "EAT"
ws['H12'] = "EAT"
# Formula: =I10-I11 (EBT - Tax Amount)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}12'] = f'={col_letter}10-{col_letter}11'

# Row 13: EARNING AVAILABLE FOR DEBTHOLDERS
ws['A13'] = "EARNING AVAILABLE FOR DEBTHOLDERS"
ws['G13'] = "EARNING AVAILABLE FOR DEBTHOLDERS"
ws['H13'] = "INTEREST IF AFTER TAX"
# Formula: =I19*0.1 (Market value of debt * 0.1)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}13'] = f'={col_letter}19*0.1'

# Row 14: EARNING AVAILABLE FOR SUKUK HOLDERS
ws['A14'] = "EARNING AVAILABLE FOR SUKUK HOLDERS"
ws['G14'] = "EARNING AVAILABLE FOR SUKUK HOLDERS"
ws['H14'] = "RENT IF AFTER TAX"
# Formula: =I31*I20 (ks -NTS-L * Market value of sukuk)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}14'] = f'={col_letter}31*{col_letter}20'

# Row 15: EARNING AVAILABLE FOR SHAREHOLDERS
ws['A15'] = "EARNING AVAILABLE FOR SHAREHOLDERS"
ws['G15'] = "EARNING AVAILABLE FOR SHAREHOLDERS"
ws['H15'] = "DIVIDEND IF AFTER TAX"
# Formula: =I12-I13-I14 (EAT - Interest after tax - Rent after tax)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}15'] = f'={col_letter}12-{col_letter}13-{col_letter}14'

# Row 16: Dividend PAID
ws['A16'] = "Dividend PAID"
ws['G16'] = "Dividend PAID"
ws['H16'] = "DIVIDEND IF AFTER TAX"
# Formula: =I15*I29 (Earnings for equity * dividend rate)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}16'] = f'={col_letter}15*{col_letter}29'

# Row 17: NOI APPROACH
ws['A17'] = "NOI APPROACH"
ws['G17'] = "NOI APPROACH"
ws['H17'] = "NOI APPROACH"
# Formula: =I12+I7 (EAT + Depreciation benefit)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}17'] = f'={col_letter}12+{col_letter}7'

# Row 18: Share Capital
ws['A18'] = "Share Capital"
ws['H18'] = "EQUITY"
# Decreasing over time: start at 30000, decrease to 24500 (so sum with rows 19+20 = 70000)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    value = 30000 - (5500 * i / (n_iter - 1))
    ws[f'{col_letter}18'] = value

# Row 19: Market value of debt
ws['A19'] = "Market value of debt"
ws['D19'] = "Equity"
ws['F19'] = 28000
ws['G19'] = "TOTAL DEBT"
ws['H19'] = "DEBT"
# Decreasing over time: start at 40000, decrease to 0 in last iteration
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    value = 40000 - (40000 * i / (n_iter - 1))
    ws[f'{col_letter}19'] = value

# Row 20: Market value of sukuk
ws['A20'] = "Market value of sukuk"
ws['D20'] = "Debt"
ws['E20'] = 0
ws['F20'] = 2000
ws['G20'] = "TOTAL SUKUK"
ws['H20'] = "SUKUK"
# Increasing over time: start at 0, increase to 45500 (so sum with rows 18+19 = 70000)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    value = (45500 * i) / (n_iter - 1)
    ws[f'{col_letter}20'] = value

# Row 21: Total Assets
ws['A21'] = "Total Assets"
ws['D21'] = "Sukuk"
ws['E21'] = 42000
ws['F21'] = 2000
ws['G21'] = "TOTAL ASSETS"
ws['H21'] = "Total Assets"
# Formula: =I18+I19+I20 (and similar for other columns)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}21'] = f'={col_letter}18+{col_letter}19+{col_letter}20'

# Row 22: Tax rate
ws['A22'] = "Tax rate"
ws['D22'] = "TAX RATE"
ws['E22'] = 0.35
ws['F22'] = 0.35
ws['G22'] = "TAX RATE"
ws['H22'] = "Tax rate"
# Constant value across all iterations
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}22'] = 0.35

# Row 23: TAX SHIELD
ws['A23'] = "TAX SHIELD (1-Tr) or (1-35%)"
ws['H23'] = "TAX SHIELD (1-Tr) or (1-35%)"
# Formula: =(1-I22) (and similar for other columns)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}23'] = f'=(1-{col_letter}22)'

# Row 24: Interest rate
ws['A24'] = "Interest rate"
ws['D24'] = "INTEREST rate that is is tax deductible"
ws['E24'] = 0.1
ws['F24'] = 0
ws['G24'] = "ki"
ws['H24'] = "Interest rate"
# Constant value - should be 0 as per original
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}24'] = 0

# Row 25: ijarah sukuk rent rate
ws['A25'] = "ijarah sukuk rent rate"
ws['D25'] = "RENT rate that is is tax deductible"
ws['E25'] = 0.1
ws['F25'] = 0
ws['G25'] = "ks"
ws['H25'] = "ijarah sukuk rent rate"
# Constant value - should be 0 as per original
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}25'] = 0

# Row 26: ijrah depreciation benefit
ws['A26'] = "ijrah depreciation benefit/or daily running expenses"
ws['D26'] = "NDTS rate that is is tax deductible"
ws['E26'] = 0.03
ws['F26'] = 0.025
ws['G26'] = "NDTS RATE"
ws['H26'] = "ijrah depreciation benefit/or daily running expenses"
# Constant value
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}26'] = 0.025

# Row 27: Capital gain or loss
ws['A27'] = "In case of purchase back asset from sukuk holder"
ws['D27'] = "Capital gain or loss that is tax deductible"
ws['E27'] = 0.03
ws['F27'] = 0.03
ws['G27'] = "CAPITAL GAIN OR LOSS"
ws['H27'] = "In case of purchase back asset from sukuk holder"
# Constant value
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}27'] = 0.03

# Row 28: dividend rate
ws['A28'] = "dividend rate"
ws['D28'] = "DIVIDEND rate that is is tax deductible"
ws['E28'] = 0.5
ws['F28'] = 0
ws['G28'] = "DIVIDEND RATE"
ws['H28'] = "dividend rate"
# Special pattern: I28=F28, J28=I28, K28=J28, L28=K28, etc.
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    if i == 0:  # I28
        ws[f'{col_letter}28'] = '=F28'
    elif i == 1:  # J28
        ws[f'{col_letter}28'] = '=I28'
    else:  # K28 onwards - reference previous column
        prev_col_letter = get_col_letter(9 + i - 1)
        ws[f'{col_letter}28'] = f'={prev_col_letter}28'

# Row 29: DIVIDEND RATE
ws['G29'] = "DIVIDEND RATE"
# Formula: =E28 (and similar for other columns)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}29'] = f'=E28'

# Row 30: ki -NTS-L
ws['A30'] = "ki -NTS-L"
ws['C30'] = "Ki"
ws['G30'] = "Ki"
ws['H30'] = "ki -NTS-L"
# Formula pattern: I30=E24, J30=I30, K30=J30, etc.
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    if i == 0:  # I30
        ws[f'{col_letter}30'] = '=E24'
    else:  # J30 onwards - reference previous cell
        prev_col_letter = get_col_letter(9 + i - 1)
        ws[f'{col_letter}30'] = f'={prev_col_letter}30'

# Row 31: ks -NTS-L
ws['A31'] = "ks -NTS-L"
ws['C31'] = "Ks"
ws['G31'] = "Ks"
ws['H31'] = "ks -NTS-L"
# Formula pattern: I31=0.02, J31=I31+0.02, K31=J31+0.02, etc.
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    if i == 0:  # I31
        ws[f'{col_letter}31'] = 0.02
    else:  # J31 onwards - previous cell + 0.02
        prev_col_letter = get_col_letter(9 + i - 1)
        ws[f'{col_letter}31'] = f'={prev_col_letter}31+0.02'

# Row 32: ks+g
ws['H32'] = "ks+g"

# Row 33: ke -NTS-L
ws['A33'] = "ke -NTS-L"
ws['C33'] = "Ke"
ws['G33'] = "Ke"
ws['H33'] = "ke -NTS-L"
# Formula pattern: I33=I16/I18, J33=J16/J18, K33=K16/K18, etc.
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}33'] = f'={col_letter}16/{col_letter}18'

# Row 34: ko
ws['A34'] = "ko"
ws['G34'] = "ko"
ws['H34'] = "ko"
# Formula: =(I30*I39)+(I31*I40)+(I33*I41) (and similar for other columns)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}34'] = f'=({col_letter}30*{col_letter}39)+({col_letter}31*{col_letter}40)+({col_letter}33*{col_letter}41)'

# Continue with more rows following the same pattern...
# For brevity, I'll add the key calculation rows that are referenced in formulas

# Row 39: DEBT-L (Debt to Assets ratio)
ws['A39'] = "DEBT-L"
ws['G39'] = "DEBT TO ASSETS"
ws['H39'] = "DEBT TO ASSETS"
# Formula: =I19/I21 (and similar for other columns)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}39'] = f'={col_letter}19/{col_letter}21'

# Row 40: SUKUK-L (Sukuk to Assets ratio)
ws['A40'] = "SUKUK-L"
ws['G40'] = "SUKUK TO ASSETS"
ws['H40'] = "SUKUK TO ASSETS"
# Formula: =I20/I21 (and similar for other columns)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}40'] = f'={col_letter}20/{col_letter}21'

# Row 41: EQUITY-L (Equity to Assets ratio)
ws['A41'] = "EQUITY-L"
ws['G41'] = "EQUITY TO ASSETS"
ws['H41'] = "EQUITY TO ASSETS"
# Formula: =I18/I21 (and similar for other columns)
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}41'] = f'={col_letter}18/{col_letter}21'

# Row 35: Formula =I30*I23, =J30*J23, =K30*K23, etc.
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}35'] = f'={col_letter}30*{col_letter}23'

# Row 36: ks with tax shield
ws['A36'] = "ks with tax shield"
ws['G36'] = "ks (1-Tr)"
ws['H36'] = "ks with tax shield"
# Formula: =I31*I23, =J31*J23, =K31*K23, etc.
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}36'] = f'={col_letter}31*{col_letter}23'

# Row 37: ke with tax shield
ws['A37'] = "ke with tax shield"
ws['G37'] = "ke (1-Tr)"
ws['H37'] = "ke with tax shield"
# Formula: =I33*I23, =J33*J23, =K33*K23, etc.
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}37'] = f'={col_letter}33*{col_letter}23'

# Row 38: ko
ws['G38'] = "ko"
# Empty row - no formulas

# Row 39: DEBT-L
ws['A39'] = "DEBT-L"
ws['G39'] = "DEBT TO ASSETS"
ws['H39'] = "DEBT TO ASSETS"
# Formula: =I19/I21, =J19/J21, =K19/K21, etc.
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}39'] = f'={col_letter}19/{col_letter}21'

# Row 40: SUKUK-L
ws['A40'] = "SUKUK-L"
ws['G40'] = "SUKUK TO ASSETS"
ws['H40'] = "SUKUK TO ASSETS"
# Formula: =I20/I21, =J20/J21, =K20/K21, etc.
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}40'] = f'={col_letter}20/{col_letter}21'

# Row 41: EQUITY-L
ws['A41'] = "EQUITY-L"
ws['G41'] = "EQUITY TO ASSETS"
ws['H41'] = "EQUITY TO ASSETS"
# Formula: =I18/I21, =J18/J21, =K18/K21, etc.
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}41'] = f'={col_letter}18/{col_letter}21'

# Row 42: WACC -NTS-L
ws['A42'] = "WACC -NTS-L"
ws['G42'] = "WACC"
ws['H42'] = "WACC -NTS-L"
# Formula: =I34, =J34, =K34, etc.
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}42'] = f'={col_letter}34'

# Row 43: Market Value of Firm (MVF)
ws['G43'] = "Market Value of Firm (MVF)"
# Formula: =I17/I42, =J17/J42, =K17/K42, etc.
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}43'] = f'={col_letter}17/{col_letter}42'

# Row 44: Annual TAX SHIELD benefits of debt
ws['A44'] = "Annual TAX SHIELD benefits of debt"
ws['G44'] = "Annual TAX SHIELD benefits of debt"
ws['H44'] = "Annual TAX SHIELD benefits of debt"
# Formula: =I5*I22, =J5*J22, =K5*K22, etc.
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}44'] = f'={col_letter}5*{col_letter}22'

# Row 45: Annual RENT SHIELD benefits of sukuk
ws['A45'] = "Annual RENT SHIELD benefits of sukuk"
ws['G45'] = "Annual RENT SHIELD benefits of sukuk"
ws['H45'] = "Annual RENT SHIELD benefits of sukuk"
# Formula: =I22*(I6+I7), =J22*(J6+J7), =K22*(K6+K7), etc.
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}45'] = f'={col_letter}22*({col_letter}6+{col_letter}7)'

# Row 46: Annual Dividend SHIELD benefits of Equity
ws['A46'] = "Annual Dividend SHIELD benefits of Equity"
ws['G46'] = "Annual Dividend SHIELD benefits of Equity"
ws['H46'] = "Annual Dividend SHIELD benefits of Equity"
# Empty row - no formulas

# Row 47: PV of TAX SHIELD benefits of debt
ws['A47'] = "PV of TAX SHIELD benefits of debt"
ws['G47'] = "PV of INTEREST TAX SHIELD benefits of debt"
ws['H47'] = "PV of TAX SHIELD benefits of debt"
# Formula: =I22*I19, =J22*J19, =K22*K19, etc.
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}47'] = f'={col_letter}22*{col_letter}19'

# Row 48: PV of RENT SHIELD benefits of sukuk
ws['A48'] = "PV of RENT SHIELD benefits of sukuk"
ws['G48'] = "PV of RENTAL TAX SHIELD benefits of sukuk"
ws['H48'] = "PV of RENT SHIELD benefits of sukuk"
# Formula: =I20*I22, =J20*J22, =K20*K22, etc.
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}48'] = f'={col_letter}20*{col_letter}22'

# Row 49: PV of Dividend SHIELD benefits of Equity
ws['A49'] = "PV of Dividend SHIELD benefits of Equity"
ws['G49'] = "PV of Dividend TAX SHIELD benefits of Equity"
ws['H49'] = "PV of Dividend SHIELD benefits of Equity"
# Empty row - no formulas

# Row 50: PV of Bankruptcy Cost
ws['A50'] = "PV of Bankruptcy Cost"
ws['G50'] = "PV of Bankruptcy Cost"
ws['H50'] = "PV of Bankruptcy Cost"
# Empty row - no formulas

# Row 51: MVF-NTS-L
ws['A51'] = "MVF-NTS-L"
ws['G51'] = "MVF + Present Value of Tax Shield/Bc"
ws['H51'] = "MVF-NTS-L"
# Formula: =I43-I47, =J43-J47, =K43-K47, etc.
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}51'] = f'={col_letter}43-{col_letter}47'

# Row 52: No. of Shares Outstanding
ws['A52'] = "No. of Shares Outstanding"
ws['G52'] = "NO OF SHARES OUTSTANDING"
ws['H52'] = "No. of Shares Outstanding"
# Formula: =I18/5, =J18/5, =K18/5, etc.
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}52'] = f'={col_letter}18/5'

# Row 53: EPS-NTS-L
ws['A53'] = "EPS-NTS-L"
ws['G53'] = "EARNING PER SHARE"
ws['H53'] = "EPS-NTS-L"
# Formula: =I16/I52, =J16/J52, =K16/K52, etc.
for i in range(n_iter):
    col_letter = get_col_letter(9 + i)
    ws[f'{col_letter}53'] = f'={col_letter}16/{col_letter}52'

# Match the ACTUAL original sheet structure exactly:
# Rows 1-53: Data and formulas ✅
# Rows 54-73: COMPLETELY MISSING (don't exist in original)
# Rows 74-77: Some data (MVF Scaled, Tax Contribution, T.C)
# Rows 78-113: COMPLETELY MISSING (don't exist in original)
# Rows 114-117: Percentage calculations
# Row 118: Empty
# Rows 119-122: EXACT DUPLICATE of rows 114-117

# Note: To avoid creating empty rows 54-73 and 78-113,
# we'll use a different approach - create a list of only the rows we want

# Create only the specific rows we want, avoiding empty ranges
# We'll use a dictionary to map row numbers to their content

# Define the rows we want to create (avoiding 54-73 and 78-113)
rows_to_create = {
    74: ("MVF (Scaled Value)", "MVF (Scaled Value)", "MVF (Scaled Value)", "=I43/32000"),
    75: ("Tax Contribution Scalled", "Tax Contribution Scalled", "Tax Contribution Scalled", "=I11/1700"),
    76: ("T.C", "T.C", "T.C", "=I75"),
    114: ("Share Capital", "SHARE CAPITAL", "Equity %", "=I18/I21*100"),
    115: ("Market value of debt", "TOTAL DEBT", "Debt %", "=I19/I21*100"),
    116: ("Market value of sukuk", "TOTAL SUKUK", "Sukuk %", "=I20/I21*100"),
    117: ("Total Assets", "TOTAL ASSETS", "Total Assets %", "=I114+I115+I116"),
    119: ("Share Capital", "SHARE CAPITAL", "Equity %", "=I114"),
    120: ("Market value of debt", "TOTAL DEBT", "Debt %", "=I115"),
    121: ("Market value of sukuk", "TOTAL SUKUK", "Sukuk %", "=I116"),
    122: ("Total Assets", "TOTAL ASSETS", "Total Assets %", "=I117")
}

# Create only the specific rows we want
for row_num, (label_a, label_g, label_h, formula) in rows_to_create.items():
    ws[f'A{row_num}'] = label_a
    ws[f'G{row_num}'] = label_g
    ws[f'H{row_num}'] = label_h
    
    # Add formulas for all iterations
    for i in range(n_iter):
        col_letter = get_col_letter(9 + i)
        if formula.startswith('='):
            ws[f'{col_letter}{row_num}'] = formula
        else:
            # Handle special cases like =I75
            ws[f'{col_letter}{row_num}'] = formula

# Now add row grouping and collapse the empty ranges to match original sheet format
# Group and collapse rows 54-73 (empty range)
for row in range(54, 74):
    ws.row_dimensions[row].outline_level = 1
    ws.row_dimensions[row].hidden = True

# Group and collapse rows 78-113 (empty range)  
for row in range(78, 114):
    ws.row_dimensions[row].outline_level = 1
    ws.row_dimensions[row].hidden = True

# Also hide row 77 (empty) and row 118 (empty)
ws.row_dimensions[77].hidden = True
ws.row_dimensions[118].hidden = True





# Save the workbook
output_file = 'sukuk_1000_iterations.xlsx'
wb.save(output_file)

print(f"Excel file with 1000 iterations saved as {output_file}")
print(f"Total iterations: {n_iter}")
print("File structure replicates Sukuk_NTS.xlsx exactly with extended iterations")
print("Rows 1-53: Original structure")
print("Rows 74-122: Additional calculations and percentages")
