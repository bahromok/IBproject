# OfficeAI - Complete User Guide

## How to Control Excel and Word Files

### EXCEL FEATURES

#### Create a Spreadsheet
```
"Create a budget.xlsx with headers: Category, Budget, Actual. Add rows for Rent, Utilities, Groceries with amounts."
```

#### Add Data and Formulas
```
"Add SUM formula to column D for totals"
"Add AVERAGE formula for the Budget column"
"Add IF formula: if Actual > Budget, show 'Over Budget', else 'On Track'"
```

#### Style Cells
```
"Make headers bold with dark blue background and white text"
"Style column A with yellow background"
"Format column B as currency ($#,##0.00)"
"Format column C as percentage (0.00%)"
"Add borders to all cells in the table"
```

#### Advanced Operations
```
"Freeze the header row"
"Merge cells A1:E1 for the title"
"Set column A width to 30"
"Add conditional formatting: highlight cells > 1000 in green"
"Add data validation dropdown in column A with options Yes/No"
"Add a new sheet called Summary"
"Delete row 5"
"Sort by column B descending"
```

#### Formulas You Can Use
- **Math**: SUM, AVERAGE, COUNT, MAX, MIN, ROUND
- **Logic**: IF, IFS, AND, OR, NOT, IFERROR
- **Lookup**: VLOOKUP, HLOOKUP, INDEX, MATCH
- **Text**: CONCATENATE, LEFT, RIGHT, MID, LEN, TRIM
- **Date**: TODAY, NOW, DATE, YEAR, MONTH, DAY
- **Stats**: SUMIF, COUNTIF, AVERAGEIF, SUMIFS
- **Financial**: PMT, PV, FV, NPV, IRR

---

### WORD FEATURES

#### Create a Document
```
"Create a business_report.docx with title 'Q1 Report', sections for Executive Summary, Financial Data with a table"
```

#### Text Editing
```
"Change 'John Doe' to 'Jane Smith' in report.docx"
"Replace all instances of 'old company' with 'new company'"
```

#### Add Content
```
"Add a heading 'Conclusion' at the end"
"Add a paragraph summarizing the findings"
"Add a bullet list with key takeaways"
"Add a numbered list of steps"
"Add a table with columns: Name, Role, Department"
```

#### Table Operations
```
"Add a new row to the table with Alice, Manager, Sales"
"Add row with Bob, Developer, IT"
```

#### Formatting
```
"Add a page break before the appendix section"
"Add a separator line"
"Add header 'Confidential - Company Name'"
"Add footer with page numbers"
"Add table of contents"
```

#### Styling (in CREATE_DOCUMENT)
```
"Create document with Calibri font, title size 36, heading size 28"
"Use blue color (#1E40AF) for headings"
"Set line spacing to 1.15"
```

---

### ADVANCED PATTERNS

#### Creating Professional Invoices
```
"Create invoice.docx with:
- Title: INVOICE
- Section: Bill To (with company name, address, email)
- Table: Item, Description, Qty, Price, Total
- Add 3 line items with amounts
- Section: Payment Terms
- Footer: Thank you for your business"
```

#### Creating Budget Trackers
```
"Create budget.xlsx with:
- Headers: Category, Jan, Feb, Mar, Q1 Total, Budget, Variance
- Add rows for: Rent, Salaries, Marketing, Travel, Utilities
- Add SUM formulas for Q1 Total column
- Add Variance formula = Budget - Q1 Total
- Format currency columns as $#,##0
- Style headers with dark background"
```

#### Creating Project Plans
```
"Create project_plan.docx with:
- Title: Project Plan - Q2
- Section: Team Members (table with Name, Role, Email)
- Add 5 team members
- Section: Timeline (bullet list of milestones)
- Section: Budget (table with Category, Amount, Status)
- Add page break between sections"
```

---

### COLOR CODES (for styling)

| Color | Hex Code |
|-------|----------|
| Dark Blue | 1E40AF |
| Blue | 3B82F6 |
| Green | 10B981 |
| Red | EF4444 |
| Yellow | F59E0B |
| Purple | 7C3AED |
| Pink | EC4899 |
| Gray | 6B7280 |
| Dark Gray | 374151 |
| White | FFFFFF |
| Black | 000000 |
| Orange | F97316 |

---

### TIPS FOR BEST RESULTS

1. **Be specific**: "Create budget.xlsx with headers Category, Amount, Notes" is better than "create a spreadsheet"
2. **Include data**: "Add rows for Rent ($2000), Utilities ($300), Food ($500)" gives exact values
3. **Mention formulas**: "Add SUM formula in cell B10 for totals" ensures calculations
4. **Specify styling**: "Make headers bold with dark background" creates professional look
5. **Use context**: After creating a file, say "Add row to it" - AI remembers what you created

---

### COMMON MISTAKES TO AVOID

1. Don't say "make it look nice" - instead specify "bold headers, alternating row colors"
2. Don't say "add calculations" - instead say "add SUM formula in column D"
3. Don't say "format the numbers" - instead say "format column B as currency $#,##0.00"
