# excel_range Tool - Number Formatting Guide

## Number Formatting Actions

**Two actions for number formatting - choose based on use case:**

### SetNumberFormat (RECOMMENDED)
**Use for:** Standard formatting that works correctly on any locale.

**Required:** `numberFormatCategory` - the type of format to apply.

**Categories:** General, Number, Currency, Accounting, Date, Time, Percentage, Fraction, Scientific, Text, Special

**Optional parameters by category:**

| Category | Parameters |
|----------|------------|
| Number | decimalPlaces, useThousandsSeparator, negativeNumberFormat |
| Currency | decimalPlaces, currencySymbol, useThousandsSeparator, negativeNumberFormat |
| Accounting | decimalPlaces, currencySymbol |
| Percentage | decimalPlaces |
| Scientific | decimalPlaces |
| Date | dateFormatStyle (ShortDate, LongDate, ISO, MonthYear, DayMonth, Year, Month, Day) |
| Time | timeFormatStyle (ShortTime, LongTime, Duration, HoursMinutes, HoursMinutesSeconds), includeDate |
| Fraction | fractionStyle (OneDigit, TwoDigits, Halves, Quarters, Eighths, Tenths, Hundredths) |
| Special | specialFormatType (ZipCode, ZipCodePlus4, PhoneNumber, SocialSecurityNumber) |

**Examples:**
- Currency with Euro: `numberFormatCategory='Currency', currencySymbol='€', decimalPlaces=2`
- Percentage: `numberFormatCategory='Percentage', decimalPlaces=1`
- Short date: `numberFormatCategory='Date', dateFormatStyle='ShortDate'`

### SetNumberFormatCustom (Expert Use)
**Use for:** Custom format codes when structured options don't cover your needs.

**Required:** `formatCode` - Excel format code string.

**Warning:** Format codes use US conventions (`#,##0.00` means comma=thousands, period=decimal). On non-US locales (German, French, etc.), these may display incorrectly. Use SetNumberFormat for locale-safe formatting.

**When to use:**
- Custom color formatting: `[Red]0.00;[Blue]-0.00`
- Conditional formatting: `[>1000]#,##0;[<0]-#,##0;0`
- Custom text: `0.00" units"`
- Complex patterns not covered by structured options

### SetNumberFormats (Bulk)
**Use for:** Applying different formats to different cells in a range.

**Required:** `formats` - 2D array matching range dimensions.

**Example:** `formats=[['$#,##0','0.00%'],['m/d/yyyy','General']]`

## Decision Tree

```
Need to format numbers?
├─ Standard format (currency, percent, date)?
│  └─ Use SetNumberFormat with numberFormatCategory
│
├─ Custom color/conditional format?
│  └─ Use SetNumberFormatCustom with formatCode
│
├─ Different formats per cell?
│  └─ Use SetNumberFormats with formats array
│
└─ Unsure about locale?
   └─ ALWAYS use SetNumberFormat (locale-safe)
```
