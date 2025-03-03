# Thai Baht Number to Words Converter

## Overview
This VBA project provides a function to convert numerical values into Thai Baht words. The function correctly formats both the Baht and Satang portions of the currency, ensuring proper grammatical structure and accurate conversions.

## Features
- Converts numbers into Thai Baht text format.
- Supports values with decimal places (Satangs).
- Correctly handles edge cases such as zero Baht, zero Satang, and full numbers without decimals.
- Uses "Zero Satang" instead of "No Satang" for clarity.
- Optimized for readability and performance.

## Usage
### Function: `SpellBaht`
Converts a number to Thai Baht words.

#### Example Usage in VBA:
```vba
Dim result As String
result = SpellBaht(1234567890.50)
' Output: "One Billion Two Hundred Thirty Four Million Five Hundred Sixty Seven Thousand Eight Hundred Ninety Baht and Fifty Satang"
```

### Expected Outputs:
| Input         | Expected Output |
|--------------|----------------|
| `1234567890.50` | One Billion Two Hundred Thirty Four Million Five Hundred Sixty Seven Thousand Eight Hundred Ninety Baht and Fifty Satang |
| `0.88` | Zero Baht and Eighty Eight Satang |
| `500.00` | Five Hundred Baht and Zero Satang |
| `1.01` | One Baht and One Satang |
| `1000000` | One Million Baht and Zero Satang |

## Functions Breakdown
### `SpellBaht(MyNumber)`
- Converts a number into Thai Baht words.
- Handles Baht and Satang separately.
- Uses helper functions to format different number ranges properly.

### `GetHundreds(MyNumber)`
- Converts numbers between 100-999 into text.

### `GetTens(TensText)`
- Converts numbers between 10-99 into text.

### `GetDigit(Digit)`
- Converts single-digit numbers (1-9) into text.

## Installation
1. Open Microsoft Excel.
2. Press `ALT + F11` to open the VBA editor.
3. Insert a new module (`Insert > Module`).
4. Copy and paste the VBA code into the module.
5. Save and use the `SpellBaht` function in your Excel formulas or VBA scripts.

## Contributions
Feel free to contribute by improving code readability, handling more edge cases, or optimizing performance. Fork the repository and submit a pull request!

## License
This project is open-source and available under the MIT License.

