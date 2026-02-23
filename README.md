# Basic-Sales-Dashboard

## Overview
This tutorial teaches how to create a complete, dynamic sales dashboard in Excel that can be customized for various business applications.

## Tutorial Content Structure

### 1. Introduction to the Dynamic Dashboard
The final dashboard includes:
- **Key Metrics**: Revenue, profit, and profit percentage
- **Top Products**: Products with highest revenue
- **Top Categories**: Categories with highest revenue
- **Data Filtering**: By year (2023, 2024, 2025) and month
- **Time-Based Charts**: Monthly and daily trend visualization
- **Sales Channels**: Direct sales, distributors, wholesale, and online
- **Payment Methods**: Bank transfer and cash with percentage breakdown
- **Interactive Features**: Checkboxes to toggle data visibility

The final result is a **Sales Dashboard** built entirely in Excel—no add-ins required.

> Example outputs of the completed tracker.

### Sales Dashboard

![Final Dashboard](imgs/Dashboard.png)


### 2. Data Preparation

#### Raw Data Structure
**Data Sheet** includes:
- Product code
- Product name
- Category
- Unit
- Purchase price
- Selling price

**Sales Sheet** includes:
- Date
- Product code
- Quantity sold
- Sales channel (wholesale, online, direct sales)
- Payment method (bank transfer, cash)
- Discount field (set to 0% in example)

#### Data Validation Setup
Create dropdown lists for data consistency:

**For Sales Channel:**
- Data > Data Validation
- Choose "List" 
- Enter values: "Online, Distributor Wholesale, Direct Sales"
- Separating with commas (or semicolons depending on regional settings)

**For Payment Method:**
- Same process
- Enter values: "Bank Transfer, Cash"

#### Converting to Tables
- Select data range
- Press Ctrl+T or Insert > Table
- Name the tables:
  - Data sheet table: `tbl_data`
  - Sales sheet table: `tbl_sale`

#### Adding Calculated Columns
Add these columns to the Sales sheet using VLOOKUP:

**Product Name:**
```excel
=VLOOKUP([@[Product Code]], tbl_data, 2, 0)
```

**Category:**
```excel
=VLOOKUP([@[Product Code]], tbl_data, 3, 0)
```

**Unit:**
```excel
=VLOOKUP([@[Product Code]], tbl_data, 4, 0)
```

**Purchase Price:**
```excel
=VLOOKUP([@[Product Code]], tbl_data, 5, 0)
```

**Selling Price:**
```excel
=VLOOKUP([@[Product Code]], tbl_data, 6, 0)
```

**Total Purchase Cost:**
```excel
=[Quantity] * [Purchase Price]
```

**Total Sales Revenue:**
```excel
=[Quantity] * [Selling Price] * (100% - [Discount])
```

**Day Column:**
```excel
=DAY([Date])
```

**Month Column:**
```excel
=TEXT([Date], "mmm")
```
This displays abbreviated month names (Jan, Feb, Mar, etc.)

**Year Column:**
```excel
=YEAR([Date])
```

### 3. Creating the Calculate Sheet with Pivot Tables

#### Basic Metrics Pivot Table
1. Select any cell in tbl_sale
2. Insert > PivotTable
3. Place in Calculate sheet at cell B2
4. Add to Values area:
   - Total Purchase Cost
   - Total Sales Revenue
5. Format as thousands with no decimals

#### Creating Calculated Fields in Pivot Table
**Revenue:**
```excel
=[Total Sales Revenue]
```

**Profit:**
```excel
=[Revenue] - [Total Purchase Cost]
```

**Profit Percentage:**
```excel
=[Profit] / [Total Purchase Cost]
```
Format as percentage

#### Revenue by Day Pivot Table
- Copy the basic pivot table
- Add "Day" to Rows
- Keep Total Sales Revenue in Values

#### Revenue by Month Pivot Table
- Copy pivot table again
- Add "Month" to Rows
- Report Layout: Show in Tabular Form
- Keep Total Sales Revenue in Values

#### Monthly Analysis with Formulas
Outside the pivot table, create formula-based monthly summary:

**Revenue by Month:**
```excel
=VLOOKUP([Month], MonthPivotRange, 3, 0)
```

**Profit by Month:**
```excel
= Revenue by Month - VLOOKUP([Month], MonthPivotRange, 2, 0)
```

**Profit Percentage by Month:**
```excel
=(Profit by Month) / [Cost]
```

#### Creating Chart for Monthly Data
1. Select the monthly data range
2. Insert > Stacked Column Chart
3. Chart customization:
   - Remove chart title
   - Remove gridlines
   - Adjust gap width to 40%
   - Add data labels for profit percentage using "Value from Cells"
   - Format labels: white text, rotate text up, position inside

### 4. Interactive Checkboxes for Chart Elements

#### Creating Checkboxes
1. Developer tab > Insert > Checkbox
   - If Developer tab not visible: Right-click ribbon > Customize Ribbon > Check Developer
2. Create three checkboxes for Revenue, Profit, Profit %
3. Link each checkbox:
   - Right-click > Format Control
   - Cell Link: M1, N1, O1 (in Calculate sheet)
   - Important: Include sheet name in link when copying to Dashboard

#### Conditional Chart Formulas
Wrap chart data with IF statements:

**Revenue Column:**
```excel
=IF($M$1=TRUE, VLOOKUP([Month], MonthData, 3, 0), NA())
```

**Profit Column:**
```excel
=IF($N$1=TRUE, VLOOKUP([Month], MonthData, 3, 0) - VLOOKUP([Month], MonthData, 2, 0), NA())
```

**Profit % Column:**
```excel
=IF($O$1=TRUE, ProfitFormula, "")
```
Note: Use empty string "" instead of NA() for profit % to avoid ugly chart display

**Error Handling:**
```excel
=IFERROR([Formula], "")
```
Use this to handle NA errors when data is missing

### 5. Category Analysis

#### Category Pivot Table
- Copy existing pivot table
- Remove previous row fields
- Add "Category" to Rows
- Keep Total Sales Revenue in Values
- Remove Grand Total

#### Finding Top Category
**Maximum Revenue:**
```excel
=MAX(CategoryPivotRange)
```

**Top Category Name:**
```excel
=INDEX(CategoryColumn, MATCH(MaxRevenue, RevenueColumn, 0))
```

#### Creating Tree Map Chart
Cannot create directly from pivot table, so:
1. Copy category names and revenue to separate range
2. Use VLOOKUP to pull data:
```excel
=VLOOKUP([Category], CategoryPivot, 2, 0)
```
3. Select data range
4. Insert > Treemap Chart
5. Add data labels with category name and value

### 6. Product Analysis

#### Product Pivot Table
Add to Rows (in order):
- Product Name
- Unit
- Total Sales Revenue
- Quantity

Remove subtotals and use tabular form

#### Finding Top Product
**Maximum Product Revenue:**
```excel
=MAX(ProductRevenueColumn)
```

**Top Product Name:**
```excel
=INDEX(ProductNameColumn, MATCH(MaxRevenue, RevenueColumn, 0))
```

**Top Product Unit:**
```excel
=INDEX(UnitColumn, MATCH(MaxRevenue, RevenueColumn, 0))
```

**Top Product Quantity:**
```excel
=INDEX(QuantityColumn, MATCH(MaxRevenue, RevenueColumn, 0))
```

### 7. Scrollable Product List

#### Creating Scrollable Range with OFFSET
```excel
=OFFSET(ProductNameCell, 1 + $AB$8, 0, 1, 1)
```
- Starts from first product
- Offsets by value in AB8 (scroll bar linked cell)
- Returns 6 products at a time

#### Adding Scroll Bar Control
1. Developer > Insert > Scroll Bar
2. Right-click > Format Control
3. Cell Link: AB8 (in Calculate sheet - ensure sheet name included)
4. Page Change: 6 (to match number of visible products)

#### Limiting Scroll Range
```excel
=MIN($AB$8, COUNTA(ProductColumn) - 7)
```
- Store in AC8
- Prevents scrolling past available products
- Subtract 7 to account for header and 6-product display

#### Chart for Products
1. Select 6-product range
2. Insert > Bar Chart
3. Format:
   - Category axis in reverse order (bottom to top)
   - Gap width: 40%
   - Add data labels (white text, inside end position)

### 8. Sales Channel and Payment Method Analysis

#### Sales Channel Pivot Table
- Add "Sales Channel" to Rows
- Total Sales Revenue to Values

#### Payment Method Pivot Table
- Add "Payment Method" to Rows
- Total Sales Revenue to Values

#### Creating Pie Charts
1. Select pivot table
2. Insert > PivotChart > Pie Chart
3. Format:
   - Remove chart title
   - Legend at top
   - Add data labels showing percentage only
   - Remove values, keep only percentages

### 9. Building the Dashboard

#### Dashboard Setup
1. Create new sheet named "Dashboard"
2. Hide ribbon: Ctrl+F1 for more working space
3. Set background color

#### Adding Shape Elements
1. Insert > Shapes > Rounded Rectangle
2. Create containers for different sections
3. Format: No outline, custom colors
4. Copy shapes (Ctrl+D) to create consistent design

#### Dashboard Sections
Create labeled areas for:
- Sales Dashboard title
- Company name (e.g., "GA Excel Shop")
- Year filter
- Month filter
- Product filter
- Sales channel filter
- Payment method filter
- Chart areas

#### Adding Icons
Insert > Icons
Search for relevant icons:
- Money/revenue icon
- Growth/trend icon
- Calendar for date filters
- Shopping/product icons
- Payment icons

Format icons with white color for contrast

#### Creating Metric Display Boxes

**Revenue Display:**
1. Insert text box
2. In formula bar: `=Calculate!$C$6`
3. Format: White text, bold, centered, large font

**Profit Display:**
Link to: `=Calculate!$C$7`

**Profit Percentage Display:**
Link to: `=Calculate!$C$8`
Format as percentage

#### Top Product Display Boxes
**Revenue:**
Link to: `=Calculate!$AC$4`

**Product Name:**
Link to: `=Calculate!$AC$2`

**Unit:**
Link to: `=Calculate!$AC$3`

**Quantity:**
Link to: `=Calculate!$AC$5`

#### Top Category Display
Similar linking to category cells in Calculate sheet

### 10. Adding Interactive Elements

#### Creating Slicers
1. Click any pivot table
2. Insert > Slicer
3. Select fields: Year, Month, Product, Sales Channel, Payment Method

#### Connecting Slicers to All Pivot Tables
1. Right-click slicer > Report Connections
2. Check all pivot tables to connect

#### Custom Slicer Styling
1. Right-click slicer style > Duplicate
2. Name: "Slicer_GAExcel"
3. Modify:
   - Whole Slicer: Match dashboard background color
   - Border: No border
   - Font: White text
   - Header: White text, no border, transparent background

#### Positioning All Elements
1. Copy all charts from Calculate sheet (Ctrl+A on chart, then Ctrl+X)
2. Paste in Dashboard (Ctrl+V)
3. Arrange elements:
   - Monthly chart with checkboxes
   - Daily chart
   - Product chart with scroll bar
   - Category tree map
   - Pie charts for channels and payment methods
   - Slicers in designated areas
   - Top product/category displays

#### Final Alignment
Use alignment tools:
- Select multiple objects (hold Ctrl)
- Format > Align > Distribute Horizontally/Vertically
- Align Top/Bottom/Left/Right as needed

### 11. Testing and Refresh

#### Adding New Data
When new products or sales are added to tables:
1. Data automatically appears in tables
2. Right-click any chart > Refresh
3. All pivot tables and charts update automatically

#### Slicer Settings for Deleted Items
Right-click slicer > Slicer Settings
Check: "Hide items with no data"

### 12. Key Features of Completed Dashboard

**Dynamic Filtering:**
- Filter by year, month, product, sales channel, payment method
- All charts update simultaneously

**Interactive Charts:**
- Toggle revenue, profit, profit % visibility with checkboxes
- Scroll through product performance
- View category breakdown in tree map

**Automatic Updates:**
- Add data to source tables
- Refresh once - entire dashboard updates

**Professional Design:**
- Consistent color scheme
- Clear labeling
- Visual hierarchy
- Icon integration

## Important Excel Formulas Used

- **VLOOKUP**: Retrieve related data
- **TEXT**: Format dates and numbers
- **DAY/MONTH/YEAR**: Extract date components
- **INDEX/MATCH**: Find specific values
- **MAX/MIN**: Find extreme values
- **COUNTA**: Count non-empty cells
- **OFFSET**: Create dynamic ranges
- **IF**: Conditional logic
- **IFERROR**: Handle errors gracefully
- **NA()**: Return not available for charts

## Tips and Best Practices

1. **Always use tables** for source data - enables automatic expansion
2. **Link cell references across sheets** carefully - include sheet names
3. **Format consistently** - use thousands separators, appropriate decimals
4. **Test interactivity** - verify slicers connect to all pivots
5. **Save frequently** - especially when working on complex layouts
6. **Use Ctrl+D** to duplicate objects quickly
7. **Lock cell references** with F4 when needed
8. **Customize to your needs** - adapt categories, metrics, and layout
9. **Consider user experience** - make filters and controls intuitive
10. **Document your work** - especially complex formulas

---
