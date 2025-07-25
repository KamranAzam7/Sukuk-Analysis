import numpy as np
import pandas as pd
import openpyxl

# Number of iterations
n_iter = 1000

# Descriptions/labels for each row (replace with your actual labels if needed)
row_labels = [
    'NET OPERATING INCOME (Approach) NOI',
    'INTEREST ON DEBT INTEREST',
    'RENT ON SUKUK RENT',
    'Benefit of asset depreciation NDTS',
    'CAPITAL GAIN OR LOSS CAPITAL GAIN OR LOSS',
    'DIVIDEND ON TAX Dividend',
    'EBT EBT',
    'TAX AMOUNT @35% TAX.C -NTS-L',
    'EARNING AFTER TAX EAT',
    'EARNING AVAILABLE FOR DEBTHOLDERS INTEREST IF AFTER TAX',
    'EARNING AVAILABLE FOR SUKUKHOLDERS RENT IF AFTER TAX',
    'EARNING AVAILABLE FOR SHAREHOLDERS DIVIDEND IF AFTER TAX',
    'Dividend PAID DIVIDEND IF AFTER TAX',
    'NOI APPROACH NOI APPROACH',
    'SHARE CAPITAL EQUITY',
    'TOTAL DEBT DEBT',
    'TOTAL SUKUK SUKUK',
    'TOTAL ASSETS Total Assets',
    'TAX RATE Tax rate',
    'TAX SHIELD (1-Tr)',
    'ki Interest rate',
    'ks ijarah sukuk rent rate',
    'NDTS RATE ijarah depreciation benefit',
    'CAPITAL GAIN RATE',
    'DIVIDEND RATE',
    'ki ki -NTS-L',
    'ks ks -NTS-L',
    'ks+g',
    'ke ke -NTS-L',
    'ko ko',
    'Ki (1-Tr)',
    'ks (1-Tr)',
    'ke (1-Tr)',
    'ko',
    '',
    'DEBT TO ASSETS',
    'SUKUK TO ASSETS',
    'EQUITY TO ASSETS',
    'WACC WACC -NTS-L',
    'Market Value of Firm (MVF)',
    'Annual TAX SHIELD (Debt)',
    'Annual RENT SHIELD (Sukuk)',
    '',
    'Annual Dividend Shield (Equity)',
    'PV of TAX SHIELD (Debt)',
    '',
    '',
    'MVF + Present Value of Tax Shield/Bc MVF-NTS-L',
    'No. of Shares Outstanding',
    'EPS EPS-NTS-L',
    'MVF (Scaled Value)',
    'Tax Contribution Scaled Tax.C.S-NTS-L',
    'T.C T.C-NTS-L',
    '',
    '% Equity',
    '% Debt',
    '% Sukuk',
    'Total %',
    '',
    'Equity %',
    'Debt %',
    'Sukuk %',
    'Total Assets %'
]

# Preallocate a dictionary to hold all row data
rows = {}

# Load the Excel file to extract initial values for recursive rows
excel_file = 'c:/Users/Lenovo thinkBook/Downloads/Sukuk_NTS.xlsx'
wb = openpyxl.load_workbook(excel_file, data_only=True)
ws = wb.active  # or ws = wb['NTS-L'] if you want a specific sheet

# Helper to get the value from Excel row N+3, column 'I' (first iteration)
def get_initial_value(row_n):
    cell = ws[f'I{row_n+3}']
    return cell.value if cell.value is not None else 0

# 1. Linear from 10000 to 52000
rows[1] = np.linspace(10000, 52000, n_iter)
# 2. All 0
rows[2] = np.zeros(n_iter)
# 3. All 0
rows[3] = np.zeros(n_iter)
# 6. All 0
rows[6] = np.zeros(n_iter)
# 10. row19 * 0.1
rows[10] = np.full(n_iter, 0.35) * 0.1  # row19 is 0.35
# 11. row31 * row20 (will be filled later)
# 12. row12 - row13 - row14 (will be filled later)
# 13. row15 * row29 (will be filled later)
# 14. row12 + row7 (will be filled later)
# 15. 30000, then prev-500
row15 = np.zeros(n_iter)
row15[0] = get_initial_value(15)  # Excel row 18 (I18)
for i in range(1, n_iter):
    row15[i] = row15[i-1] - 500
rows[15] = row15
# 16. Linear from 40000 to 0
rows[16] = np.linspace(40000, 0, n_iter)
# 17. Linear from 0 to 50500
rows[17] = np.linspace(0, 50500, n_iter)
# 18. All 70000
rows[18] = np.full(n_iter, 70000)
# 19. 35%
rows[19] = np.full(n_iter, 0.35)
# 20. 65%
rows[20] = np.full(n_iter, 0.65)
# 21. 0
rows[21] = np.zeros(n_iter)
# 22. 0
rows[22] = np.zeros(n_iter)
# 23. 0.025
rows[23] = np.full(n_iter, 0.025)
# 24. 0.03
rows[24] = np.full(n_iter, 0.03)
# 25. 0
rows[25] = np.zeros(n_iter)
# 26. 0.5
rows[26] = np.full(n_iter, 0.5)
# 27. 0.1
rows[27] = np.full(n_iter, 0.1)
# 28. 0.02, then prev+0.02
row28 = np.zeros(n_iter)
row28[0] = get_initial_value(28)  # Excel row 31 (I31)
for i in range(1, n_iter):
    row28[i] = row28[i-1] + 0.02
rows[28] = row28
# 29. Empty
rows[29] = np.full(n_iter, np.nan)
# 35. Empty
rows[35] = np.full(n_iter, np.nan)
# 41. All 0
rows[41] = np.zeros(n_iter)
# 43. Empty
rows[43] = np.full(n_iter, np.nan)
# 46. Empty
rows[46] = np.full(n_iter, np.nan)
# 47. Empty
rows[47] = np.full(n_iter, np.nan)
# 54. Empty
rows[54] = np.full(n_iter, np.nan)
# 58. row114 + row115 + row116 (not defined, set as nan)
rows[58] = np.full(n_iter, np.nan)
# 59. Empty
rows[59] = np.full(n_iter, np.nan)
# 60. Linear from 42.8571428571429 to 27.857141852296
rows[60] = np.linspace(42.8571428571429, 27.857141852296, n_iter)
# 61. Linear from 57.1428571428571 to 0
rows[61] = np.linspace(57.1428571428571, 0, n_iter)
# 62. Linear from 0 to 72.142858147704
rows[62] = np.linspace(0, 72.142858147704, n_iter)
# 63. All 100
rows[63] = np.full(n_iter, 100)

# 4. row20 * row26
rows[4] = rows[20] * rows[26]
# 5. row20 * row27
rows[5] = rows[20] * rows[27]
# 8. row10 * row22
rows[8] = rows[10] * rows[22] if 10 in rows and 22 in rows else np.full(n_iter, np.nan)
# 9. row10 - row11
rows[9] = rows[10] - rows[11] if 10 in rows and 11 in rows else np.full(n_iter, np.nan)
# 7. Iterative: row4 - row5 - row6 - prev_row7 + row8 - row9
row7 = np.zeros(n_iter)
row7[0] = get_initial_value(7)  # Excel row 10 (I10)
for i in range(1, n_iter):
    row7[i] = (
        rows[4][i] - rows[5][i] - rows[6][i] - row7[i-1] + rows[8][i] - rows[9][i]
    )
rows[7] = row7

# Continue with the rest of the formula-based rows, ensuring dependencies are defined first
# 12. row12 - row13 - row14 (will be filled later)
# 13. row15 * row29 (will be filled later)
# 14. row12 + row7 (will be filled later)
# 30. row16 / row18
rows[30] = rows[16] / rows[18]
# 31. row30*row39 + row31*row40 + row33*row41 (will be filled later)
# 32. row30 * row23
rows[32] = rows[30] * rows[23]
# 33. row31 * row23 (will be filled later)
# 34. row33 * row23 (will be filled later)
# 36. row19 / row21
rows[36] = rows[19] / rows[21] if np.all(rows[21] != 0) else np.full(n_iter, np.nan)
# 37. row20 / row21
rows[37] = rows[20] / rows[21] if np.all(rows[21] != 0) else np.full(n_iter, np.nan)
# 38. row18 / row21
rows[38] = rows[18] / rows[21] if np.all(rows[21] != 0) else np.full(n_iter, np.nan)
# 39. row34 (will be filled later)
# 40. row17 / row42 (will be filled later)
# 42. row22 * (row6 + row7)
rows[42] = rows[22] * (rows[6] + rows[7])
# 44. row22 * row19
rows[44] = rows[22] * rows[19]
# 45. row20 * row22
rows[45] = rows[20] * rows[22]
# 48. row43 - row47
rows[48] = rows[43] - rows[47]
# 49. row18 / 5
rows[49] = rows[18] / 5
# 50. row16 / row52 (will be filled later)
# 51. row51 / 32000 (will be filled later)
# 52. row11 / 1700 (will be filled later)
# 53. row30 + row31 + row33 (will be filled later)
# 55. row18 / row21 * 100
rows[55] = rows[18] / rows[21] * 100 if np.all(rows[21] != 0) else np.full(n_iter, np.nan)
# 56. row19 / row21 * 100
rows[56] = rows[19] / rows[21] * 100 if np.all(rows[21] != 0) else np.full(n_iter, np.nan)
# 57. row20 / row21 * 100
rows[57] = rows[20] / rows[21] * 100 if np.all(rows[21] != 0) else np.full(n_iter, np.nan)

# Build DataFrame
row_indices = list(range(1, 64))
data = [rows.get(i, np.full(n_iter, np.nan)) for i in row_indices]
df = pd.DataFrame(data, index=row_labels[:len(data)])

# Write to Excel
output_file = 'sukuk_1000_iterations.xlsx'
df.to_excel(output_file, header=[f'Iter_{i+1}' for i in range(n_iter)])

print(f"Excel file with 1000 iterations saved as {output_file}")
