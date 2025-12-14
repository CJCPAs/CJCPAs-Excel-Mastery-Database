# Data Tab - Complete Reference

> **Import, transform, sort, filter, and analyze data**

**Ribbon Access:** Press **Alt+A** to activate Data tab

---

## Tab Overview

The Data tab handles all data management operations:

| Group | Purpose |
|-------|---------|
| [Get & Transform Data](#get--transform-data-group) | Import and Power Query |
| [Queries & Connections](#queries--connections-group) | Manage data sources |
| [Sort & Filter](#sort--filter-group) | Organize data |
| [Data Tools](#data-tools-group) | Text to Columns, Remove Duplicates |
| [Forecast](#forecast-group) | Predictive analysis |
| [Outline](#outline-group) | Group and subtotal |
| [Data Types](#data-types-group) | Linked data (Excel 365) |

---

## Get & Transform Data Group

### Get Data
| Item | Description | Shortcut |
|------|-------------|----------|
| **Get Data** | Import data from sources | - |

#### Get Data Sources
| Category | Options |
|----------|---------|
| **From File** | Excel Workbook, Text/CSV, XML, JSON, PDF, From Folder, From SharePoint Folder |
| **From Database** | SQL Server, Access, Analysis Services, SQL Server Analysis Services, Oracle, IBM Db2, MySQL, PostgreSQL, Sybase, Teradata, SAP HANA |
| **From Azure** | Azure SQL Database, Azure Synapse Analytics, Azure HDInsight, Azure Blob Storage, Azure Table Storage, Azure Data Lake Storage, Azure HDInsight Spark |
| **From Power Platform** | Power BI datasets, Dataflows, Dataverse |
| **From Online Services** | SharePoint Online List, Microsoft Exchange Online, Dynamics 365, Salesforce Objects, Salesforce Reports, Google Analytics, Adobe Analytics, etc. |
| **From Other Sources** | Web, OData Feed, ODBC, OLEDB, Active Directory, Microsoft Exchange, Hadoop File (HDFS), Spark, R script, Python script, Blank Query |
| **Combine Queries** | Merge (join queries), Append (stack queries) |
| **Launch Power Query Editor** | Open query editor directly |
| **Data Source Settings** | Manage credentials and connections |
| **Query Options** | Power Query settings |

### From Text/CSV
| Item | Description | Shortcut |
|------|-------------|----------|
| **From Text/CSV** | Import delimited files | - |

#### Import Options
| Option | Description |
|--------|-------------|
| **File Origin** | Character encoding |
| **Delimiter** | Comma, Tab, Semicolon, etc. |
| **Data Type Detection** | Automatic type inference |
| **Load** | Direct to worksheet |
| **Transform Data** | Open in Power Query |

### From Web
| Item | Description | Shortcut |
|------|-------------|----------|
| **From Web** | Import from URL | - |

#### Web Import Options
| Option | Description |
|--------|-------------|
| **Basic** | Enter URL directly |
| **Advanced** | URL parts, headers, parameters |

### From Table/Range
| Item | Description | Shortcut |
|------|-------------|----------|
| **From Table/Range** | Send existing data to Power Query | - |

### Recent Sources
| Item | Description | Shortcut |
|------|-------------|----------|
| **Recent Sources** | Recently used connections | - |

### Existing Connections
| Item | Description | Shortcut |
|------|-------------|----------|
| **Existing Connections** | Saved connections | - |

---

## Queries & Connections Group

### Queries & Connections
| Item | Description | Shortcut |
|------|-------------|----------|
| **Queries & Connections** | Opens side pane | - |

#### Pane Features
| Tab | Contents |
|-----|----------|
| **Queries** | All Power Query queries |
| **Connections** | All data connections |

#### Query Actions (Right-click)
| Action | Description |
|--------|-------------|
| **Edit** | Open in Power Query |
| **Delete** | Remove query |
| **Rename** | Change query name |
| **Refresh** | Update data |
| **Duplicate** | Copy query |
| **Reference** | Create dependent query |
| **Merge** | Join with another query |
| **Append** | Stack with another query |
| **Properties** | Query settings |
| **Load To...** | Change output destination |

### Refresh All
| Item | Description | Shortcut |
|------|-------------|----------|
| **Refresh All** | Update all data connections | Ctrl+Alt+F5 |

#### Refresh Options
| Option | Description | Shortcut |
|--------|-------------|----------|
| **Refresh All** | All queries and connections | Ctrl+Alt+F5 |
| **Refresh** | Current selection only | Alt+F5 |
| **Connection Properties** | Configure refresh settings |
| **Cancel Refresh** | Stop current refresh |

#### Connection Properties
| Tab | Settings |
|-----|----------|
| **Usage** | Description, refresh settings |
| **Definition** | Connection string, command |

#### Refresh Settings
| Option | Description |
|--------|-------------|
| **Enable background refresh** | Refresh asynchronously |
| **Refresh every X minutes** | Auto-refresh interval |
| **Refresh data when opening the file** | Refresh on open |
| **Remove data from the external data range before saving** | Privacy option |

---

## Sort & Filter Group

### Sort A to Z
| Item | Description | Shortcut |
|------|-------------|----------|
| **Sort A to Z** | Ascending sort | - |

### Sort Z to A
| Item | Description | Shortcut |
|------|-------------|----------|
| **Sort Z to A** | Descending sort | - |

### Sort
| Item | Description | Shortcut |
|------|-------------|----------|
| **Sort** | Custom sort dialog | - |

#### Sort Dialog Options
| Element | Description |
|---------|-------------|
| **Add Level** | Add sort criterion |
| **Delete Level** | Remove criterion |
| **Copy Level** | Duplicate criterion |
| **Move Up/Down** | Change priority |
| **Options...** | Case sensitive, orientation |
| **My data has headers** | First row is header |

#### Sort Options
| Option | Description |
|--------|-------------|
| **Column** | Which column to sort by |
| **Sort On** | Values, Cell Color, Font Color, Cell Icon |
| **Order** | A to Z, Z to A, Custom List |

### Filter
| Item | Description | Shortcut |
|------|-------------|----------|
| **Filter** | Toggle AutoFilter | Ctrl+Shift+L |

#### AutoFilter Features
| Feature | Description |
|---------|-------------|
| **Sort options** | A-Z, Z-A, by color |
| **Text Filters** | Contains, Begins With, Ends With, Custom |
| **Number Filters** | Equals, Greater Than, Top 10, Above Average |
| **Date Filters** | Today, This Week, This Month, This Year, Custom |
| **Filter by Color** | Cell color, font color |
| **Search box** | Type to filter list |
| **Select All** | Check/uncheck all |
| **Clear Filter** | Remove from column |

#### Number Filter Options
| Filter | Description |
|--------|-------------|
| **Equals** | Exact match |
| **Does Not Equal** | Exclude value |
| **Greater Than** | Above threshold |
| **Greater Than Or Equal To** | At or above |
| **Less Than** | Below threshold |
| **Less Than Or Equal To** | At or below |
| **Between** | Range of values |
| **Top 10** | Top/bottom N items/percent |
| **Above Average** | Above mean |
| **Below Average** | Below mean |
| **Custom Filter** | Complex conditions |

#### Text Filter Options
| Filter | Description |
|--------|-------------|
| **Equals** | Exact match |
| **Does Not Equal** | Exclude text |
| **Begins With** | Starts with text |
| **Ends With** | Ends with text |
| **Contains** | Includes text |
| **Does Not Contain** | Excludes text |
| **Custom Filter** | Complex conditions |

#### Date Filter Options
| Filter | Description |
|--------|-------------|
| **Equals** | Specific date |
| **Before** | Earlier than date |
| **After** | Later than date |
| **Between** | Date range |
| **Tomorrow/Today/Yesterday** | Relative dates |
| **Next/This/Last Week** | Week ranges |
| **Next/This/Last Month** | Month ranges |
| **Next/This/Last Quarter** | Quarter ranges |
| **Next/This/Last Year** | Year ranges |
| **Year to Date** | Jan 1 to today |
| **All Dates in Period** | Quarter or month |
| **Custom Filter** | Complex conditions |

### Clear
| Item | Description | Shortcut |
|------|-------------|----------|
| **Clear** | Remove all filters | - |

### Reapply
| Item | Description | Shortcut |
|------|-------------|----------|
| **Reapply** | Refresh filter results | Ctrl+Alt+L |

### Advanced
| Item | Description | Shortcut |
|------|-------------|----------|
| **Advanced** | Complex filtering | - |

#### Advanced Filter Options
| Option | Description |
|--------|-------------|
| **Filter the list, in-place** | Hide non-matching rows |
| **Copy to another location** | Extract matches |
| **List range** | Data to filter |
| **Criteria range** | Filter conditions |
| **Copy to** | Output destination |
| **Unique records only** | Remove duplicates |

---

## Data Tools Group

### Text to Columns
| Item | Description | Shortcut |
|------|-------------|----------|
| **Text to Columns** | Split text into columns | - |

#### Text to Columns Wizard
| Step | Options |
|------|---------|
| **Step 1** | Delimited or Fixed width |
| **Step 2** | Choose delimiters (Tab, Semicolon, Comma, Space, Other) or set column breaks |
| **Step 3** | Set column data formats (General, Text, Date, Skip column) |

### Flash Fill
| Item | Description | Shortcut |
|------|-------------|----------|
| **Flash Fill** | Pattern-based auto-fill | Ctrl+E |

#### Flash Fill Uses
| Example | Description |
|---------|-------------|
| **Extract first name** | Type pattern, Flash Fill completes |
| **Combine columns** | Join text with formatting |
| **Reformat data** | Change phone number format |
| **Extract domains** | Get domain from email |

### Remove Duplicates
| Item | Description | Shortcut |
|------|-------------|----------|
| **Remove Duplicates** | Delete duplicate rows | - |

#### Remove Duplicates Dialog
| Option | Description |
|--------|-------------|
| **Select All** | Check all columns |
| **Unselect All** | Uncheck all columns |
| **Column checkboxes** | Include in comparison |
| **My data has headers** | First row is header |

### Data Validation
| Item | Description | Shortcut |
|------|-------------|----------|
| **Data Validation** | Control cell input | - |

#### Data Validation Dialog
| Tab | Purpose |
|-----|---------|
| **Settings** | Define validation rules |
| **Input Message** | Show message on cell entry |
| **Error Alert** | Message when invalid data entered |

#### Validation Types
| Type | Description |
|------|-------------|
| **Any value** | No restriction |
| **Whole number** | Integers only |
| **Decimal** | Numbers with decimals |
| **List** | Dropdown from list |
| **Date** | Valid dates only |
| **Time** | Valid times only |
| **Text length** | Character count limits |
| **Custom** | Formula-based validation |

#### Validation Conditions
| Condition | Description |
|-----------|-------------|
| **between** | Within range |
| **not between** | Outside range |
| **equal to** | Exact match |
| **not equal to** | Anything but |
| **greater than** | Above value |
| **less than** | Below value |
| **greater than or equal to** | At or above |
| **less than or equal to** | At or below |

#### Error Alert Styles
| Style | Description |
|-------|-------------|
| **Stop** | Reject invalid entry |
| **Warning** | Allow override |
| **Information** | Allow override |

### Consolidate
| Item | Description | Shortcut |
|------|-------------|----------|
| **Consolidate** | Combine data from ranges | - |

#### Consolidate Options
| Function | Description |
|----------|-------------|
| **Sum** | Add values |
| **Count** | Count values |
| **Average** | Average values |
| **Max** | Maximum value |
| **Min** | Minimum value |
| **Product** | Multiply values |
| **Count Nums** | Count numbers |
| **StdDev** | Standard deviation |
| **StdDevp** | Population std dev |
| **Var** | Variance |
| **Varp** | Population variance |

| Option | Description |
|--------|-------------|
| **All references** | Ranges to consolidate |
| **Top row** | Labels in top row |
| **Left column** | Labels in left column |
| **Create links** | Link to source data |

### What-If Analysis
| Item | Description | Shortcut |
|------|-------------|----------|
| **What-If Analysis** | Scenario and sensitivity tools | - |

#### What-If Options
| Tool | Description |
|------|-------------|
| **Scenario Manager...** | Create/compare scenarios |
| **Goal Seek...** | Find input for target output |
| **Data Table...** | Sensitivity analysis table |

#### Scenario Manager
| Action | Description |
|--------|-------------|
| **Add** | Create new scenario |
| **Delete** | Remove scenario |
| **Edit** | Modify scenario |
| **Merge** | Import from other sheets |
| **Summary** | Generate report |

#### Goal Seek Dialog
| Field | Description |
|-------|-------------|
| **Set cell** | Cell with formula |
| **To value** | Target result |
| **By changing cell** | Input to vary |

#### Data Table
| Type | Description |
|------|-------------|
| **One-variable** | One input, multiple scenarios |
| **Two-variable** | Two inputs, matrix of results |

---

## Forecast Group

### What-If Analysis
*(See above - same button)*

### Forecast Sheet
| Item | Description | Shortcut |
|------|-------------|----------|
| **Forecast Sheet** | Create forecast chart | - |

#### Forecast Options
| Option | Description |
|--------|-------------|
| **Forecast End** | How far to predict |
| **Confidence Interval** | Prediction range |
| **Seasonality** | Detect or set manually |
| **Timeline Range** | Date column |
| **Values Range** | Data to forecast |
| **Fill Missing Points** | Interpolation method |
| **Aggregate Duplicates** | How to handle duplicates |

---

## Outline Group

### Group
| Item | Description | Shortcut |
|------|-------------|----------|
| **Group** | Create expandable sections | Alt+Shift+Right |

#### Group Options
| Option | Description |
|--------|-------------|
| **Group** | Group selected rows/columns |
| **Ungroup** | Remove grouping |
| **Auto Outline** | Create outline automatically |
| **Clear Outline** | Remove all groups |
| **Settings** | Outline direction settings |

### Ungroup
| Item | Description | Shortcut |
|------|-------------|----------|
| **Ungroup** | Remove grouping | Alt+Shift+Left |

### Subtotal
| Item | Description | Shortcut |
|------|-------------|----------|
| **Subtotal** | Add subtotals to list | - |

#### Subtotal Dialog
| Option | Description |
|--------|-------------|
| **At each change in** | Group by column |
| **Use function** | SUM, COUNT, AVERAGE, etc. |
| **Add subtotal to** | Columns to subtotal |
| **Replace current subtotals** | Overwrite existing |
| **Page break between groups** | For printing |
| **Summary below data** | Subtotal position |
| **Remove All** | Delete all subtotals |

### Show Detail
| Item | Description | Shortcut |
|------|-------------|----------|
| **Show Detail** | Expand group | - |

### Hide Detail
| Item | Description | Shortcut |
|------|-------------|----------|
| **Hide Detail** | Collapse group | - |

---

## Data Types Group (Excel 365)

### Stocks
| Item | Description | Shortcut |
|------|-------------|----------|
| **Stocks** | Convert to stock data type | - |

#### Stock Data Fields
| Field | Description |
|-------|-------------|
| **Price** | Current stock price |
| **Change** | Price change |
| **Change (%)** | Percentage change |
| **Open** | Opening price |
| **High** | Day high |
| **Low** | Day low |
| **52-week high** | Year high |
| **52-week low** | Year low |
| **Market cap** | Market capitalization |
| **PE ratio** | Price to earnings |
| **Volume** | Trading volume |

### Geography
| Item | Description | Shortcut |
|------|-------------|----------|
| **Geography** | Convert to geographic data | - |

#### Geography Data Fields
| Field | Description |
|-------|-------------|
| **Population** | Total population |
| **Area** | Land area |
| **Capital** | Capital city |
| **Leader(s)** | Government leaders |
| **GDP** | Gross domestic product |
| **Currency** | Official currency |
| **Time zone** | Time zone |
| **Largest city** | Most populous city |

---

## Complete Alt Shortcuts Reference

### Get & Transform Data Group (Alt+A)
| Action | Alt Shortcut |
|--------|--------------|
| Get Data | Alt+A, P, N |
| From Text/CSV | Alt+A, F, T |
| From Web | Alt+A, F, W |
| From Table/Range | Alt+A, F, R |
| Recent Sources | Alt+A, C, R |
| Existing Connections | Alt+A, X |

### Queries & Connections Group (Alt+A)
| Action | Alt Shortcut | Other |
|--------|--------------|-------|
| Queries & Connections Pane | Alt+A, O | - |
| Refresh | Alt+A, R, R | Alt+F5 |
| Refresh All | Alt+A, R, A | Ctrl+Alt+F5 |

### Sort & Filter Group (Alt+A)
| Action | Alt Shortcut | Other |
|--------|--------------|-------|
| Sort A to Z | Alt+A, S, A | - |
| Sort Z to A | Alt+A, S, D | - |
| Sort (Custom) | Alt+A, S, S | - |
| Filter Toggle | Alt+A, T | Ctrl+Shift+L |
| Clear Filter | Alt+A, C | - |
| Reapply Filter | Alt+A, Y | Ctrl+Alt+L |
| Advanced Filter | Alt+A, Q | - |

### Data Tools Group (Alt+A)
| Action | Alt Shortcut | Other |
|--------|--------------|-------|
| Text to Columns | Alt+A, E | - |
| Flash Fill | Alt+A, I | Ctrl+E |
| Remove Duplicates | Alt+A, M | - |
| Data Validation | Alt+A, V, V | - |
| Data Validation (Circle Invalid) | Alt+A, V, I | - |
| Data Validation (Clear Circles) | Alt+A, V, C | - |
| Consolidate | Alt+A, N | - |
| What-If Analysis | Alt+A, W | - |
| Goal Seek | Alt+A, W, G | - |
| Scenario Manager | Alt+A, W, S | - |
| Data Table | Alt+A, W, T | - |

### Forecast Group (Alt+A)
| Action | Alt Shortcut |
|--------|--------------|
| Forecast Sheet | Alt+A, D |

### Outline Group (Alt+A)
| Action | Alt Shortcut | Other |
|--------|--------------|-------|
| Group | Alt+A, G, G | Alt+Shift+Right |
| Ungroup | Alt+A, U, U | Alt+Shift+Left |
| Subtotal | Alt+A, B | - |
| Auto Outline | Alt+A, G, A | - |
| Clear Outline | Alt+A, G, C | - |
| Show Detail | Alt+A, J | - |
| Hide Detail | Alt+A, H | - |

### Data Types Group (Alt+A)
| Action | Alt Shortcut |
|--------|--------------|
| Stocks | Alt+A, K |
| Geography | Alt+A, G, E |

---

## Keyboard Shortcuts Summary

| Action | Alt Shortcut | Other |
|--------|--------------|-------|
| Toggle AutoFilter | Alt+A, T | Ctrl+Shift+L |
| Reapply filter | Alt+A, Y | Ctrl+Alt+L |
| Flash Fill | Alt+A, I | Ctrl+E |
| Refresh connection | Alt+A, R, R | Alt+F5 |
| Refresh all | Alt+A, R, A | Ctrl+Alt+F5 |
| Sort A to Z | Alt+A, S, A | - |
| Sort Z to A | Alt+A, S, D | - |
| Text to Columns | Alt+A, E | - |
| Remove Duplicates | Alt+A, M | - |
| Data Validation | Alt+A, V, V | - |
| Group rows/columns | Alt+A, G, G | Alt+Shift+Right |
| Ungroup rows/columns | Alt+A, U, U | Alt+Shift+Left |
| Subtotal | Alt+A, B | - |

---

[🎗️ Back to Ribbon Reference](./README.md) | [🏠 Back to Home](../README.md)
