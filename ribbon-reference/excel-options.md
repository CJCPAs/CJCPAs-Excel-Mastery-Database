# Excel Options - Complete Reference

> **Every setting in File → Options, fully documented**

Access: **File → Options** or **Alt+F, T**

---

## Options Categories

| Category | Purpose |
|----------|---------|
| [General](#general) | User interface and startup |
| [Formulas](#formulas) | Calculation and error checking |
| [Data](#data) | Data import and refresh |
| [Proofing](#proofing) | Spelling and AutoCorrect |
| [Save](#save) | File saving behavior |
| [Language](#language) | Display and editing languages |
| [Ease of Access](#ease-of-access) | Accessibility features |
| [Advanced](#advanced) | Detailed behavior settings |
| [Customize Ribbon](#customize-ribbon) | Ribbon customization |
| [Quick Access Toolbar](#quick-access-toolbar) | QAT customization |
| [Add-ins](#add-ins) | Add-in management |
| [Trust Center](#trust-center) | Security settings |

---

## General

### User Interface Options
| Option | Description | Default |
|--------|-------------|---------|
| **Show Mini Toolbar on selection** | Floating format toolbar | On |
| **Enable Live Preview** | Preview formatting before applying | On |
| **Show Quick Analysis options on selection** | Analysis button on selection | On |
| **ScreenTip style** | Full, partial, or no ScreenTips | Show feature descriptions |

### When creating new workbooks
| Option | Description | Default |
|--------|-------------|---------|
| **Use this as the default font** | Font for new workbooks | Calibri |
| **Font size** | Size for new workbooks | 11 |
| **Default view for new sheets** | Normal, Page Break, Page Layout | Normal |
| **Include this many sheets** | Starting sheet count | 1 |

### Personalize your copy of Microsoft Office
| Option | Description |
|--------|-------------|
| **User name** | Name shown in comments |
| **Initials** | Used in some features |
| **Office Background** | Decorative pattern |
| **Office Theme** | Colorful, Dark Gray, Black, White |
| **Always use these values** | Override sign-in name |

### Start up options
| Option | Description | Default |
|--------|-------------|---------|
| **Show the Start screen** | Show recent files on launch | On |
| **Open all files in this folder** | XLSTART alternative path | (blank) |
| **At startup, open all files in** | Alternate startup folder | (blank) |

### Real-Time Collaboration Options
| Option | Description | Default |
|--------|-------------|---------|
| **Show changes** | Highlight collaborator edits | On |
| **Show names on presence flags** | Display names on selections | On |

---

## Formulas

### Calculation options
| Option | Description | Default |
|--------|-------------|---------|
| **Automatic** | Recalculate on change | Selected |
| **Automatic except for data tables** | Skip data tables | - |
| **Manual** | Recalculate only on request | - |
| **Recalculate workbook before saving** | Ensure accuracy when saving | On |
| **Enable iterative calculation** | Allow circular references | Off |
| **Maximum Iterations** | Iteration limit | 100 |
| **Maximum Change** | Convergence threshold | 0.001 |

### Working with formulas
| Option | Description | Default |
|--------|-------------|---------|
| **R1C1 reference style** | Use R1C1 instead of A1 | Off |
| **Formula AutoComplete** | Show function suggestions | On |
| **Use table names in formulas** | Structured references | On |
| **Use GetPivotData functions** | For PivotTable references | On |

### Error Checking
| Option | Description | Default |
|--------|-------------|---------|
| **Enable background error checking** | Check for errors automatically | On |
| **Indicate errors using this color** | Error indicator color | Green |
| **Reset Ignored Errors** | Restore dismissed errors | Button |

### Error checking rules
| Rule | Description | Default |
|------|-------------|---------|
| **Cells containing formulas that result in an error** | Flag #VALUE!, #REF!, etc. | On |
| **Inconsistent calculated column formula in tables** | Broken table formulas | On |
| **Cells containing years represented as 2 digits** | Y2K check | On |
| **Numbers formatted as text** | Numbers stored as text | On |
| **Formulas inconsistent with other formulas** | Pattern violations | On |
| **Formulas which omit cells in a region** | Range expansion suggestions | On |
| **Unlocked cells containing formulas** | Security check | Off |
| **Formulas referring to empty cells** | Blank cell references | Off |
| **Data entered in a table is invalid** | Validation errors | On |
| **Misleading number formats** | Format mismatches | On |

---

## Data

### Data options
| Option | Description | Default |
|--------|-------------|---------|
| **Disable undo for large PivotTable refresh** | Performance optimization | Off |
| **Disable undo for large data model operations** | Performance optimization | Off |
| **Prefer the Excel Data Model** | Default for relationships | Off |

### Show legacy data import wizards
| Option | Description | Default |
|--------|-------------|---------|
| **From Text (Legacy)** | Old text import | Off |
| **From Access (Legacy)** | Old Access import | Off |
| **From Web (Legacy)** | Old web query | Off |

### Automatic Data Conversion
| Option | Description | Default |
|--------|-------------|---------|
| **Convert continuous letters and numbers to a date** | Auto-date detection | On |
| **Convert scientific notation to number** | Auto-number detection | On |
| **Convert numeric data that starts with "0" to text** | Preserve leading zeros | On |
| **Truncate leading zeros of numeric data** | Remove leading zeros | Off |

---

## Proofing

### AutoCorrect options
| Button | Opens |
|--------|-------|
| **AutoCorrect Options...** | AutoCorrect dialog |

#### AutoCorrect Tab
| Option | Description | Default |
|--------|-------------|---------|
| **Show AutoCorrect Options buttons** | Show smart tag | On |
| **Correct TWo INitial CApitals** | Fix double caps | On |
| **Capitalize first letter of sentences** | Auto-capitalize | Off |
| **Capitalize names of days** | Monday, Tuesday, etc. | On |
| **Correct accidental use of cAPS LOCK key** | Fix caps lock | On |
| **Replace text as you type** | AutoCorrect replacements | On |

#### AutoFormat As You Type Tab
| Option | Description | Default |
|--------|-------------|---------|
| **Internet and network paths with hyperlinks** | Auto-link URLs | On |
| **Include new rows and columns in table** | Table auto-expand | On |
| **Fill formulas in tables** | Auto-fill formulas | On |
| **Automatic bulleted lists** | Start bullets | Off |
| **Set left- and first-indent with tabs** | Tab behavior | On |

### When correcting spelling in Microsoft Office programs
| Option | Description | Default |
|--------|-------------|---------|
| **Ignore words in UPPERCASE** | Skip acronyms | On |
| **Ignore words that contain numbers** | Skip alphanumeric | On |
| **Ignore Internet and file addresses** | Skip URLs | On |
| **Flag repeated words** | Find duplicates | On |
| **Suggest from main dictionary only** | Limit suggestions | Off |
| **Custom Dictionaries...** | Manage dictionaries | Button |

---

## Save

### Save workbooks
| Option | Description | Default |
|--------|-------------|---------|
| **Save files in this format** | Default file format | .xlsx |
| **AutoRecover information every X minutes** | Auto-save interval | 10 |
| **Keep the last AutoRecovered version** | Keep if closed without saving | On |
| **AutoRecover file location** | Recovery file path | (system) |
| **Default local file location** | Default save path | Documents |
| **Default personal templates location** | Custom templates folder | (blank) |
| **Don't show the Backstage** | Skip file info screen | Off |
| **Show additional places for saving** | OneDrive, SharePoint | On |
| **Save to Computer by default** | Local save default | Off |

### AutoRecover exceptions for this workbook
| Option | Description | Default |
|--------|-------------|---------|
| **Disable AutoRecover for this workbook only** | Turn off for current file | Off |

### Offline editing options
| Option | Description | Default |
|--------|-------------|---------|
| **Save checked-out files to** | Server draft or local | Server |

### Preserve visual appearance
| Option | Description | Default |
|--------|-------------|---------|
| **Choose what colors will be seen** | Color fidelity | (current) |

---

## Language

### Office display Language
| Column | Description |
|--------|-------------|
| **Language** | UI language |
| **Installed** | Status |
| **Preferred** | Drag to set preference |

### Office authoring languages and proofing
| Column | Description |
|--------|-------------|
| **Language** | Content language |
| **Spelling** | Spell check available |
| **Grammar** | Grammar check available |

| Button | Description |
|--------|-------------|
| **Add a Language...** | Install new language |
| **Not installed** | Get missing proofing tools |

---

## Ease of Access

### Feedback options
| Option | Description | Default |
|--------|-------------|---------|
| **Provide feedback with sound** | Audio cues | Off |
| **Provide feedback with animation** | Visual motion | On |
| **ScreenTip style** | Tooltip behavior | Show |

### Application display options
| Option | Description | Default |
|--------|-------------|---------|
| **Show function ScreenTips** | Help on hover | On |

### Automatic Alt Text
| Option | Description | Default |
|--------|-------------|---------|
| **Automatically generate alt text for me** | AI descriptions | On |

### Document display options
| Option | Description | Default |
|--------|-------------|---------|
| **Optimize for compatibility** | Assistive technology support | Off |

---

## Advanced

### Editing options
| Option | Description | Default |
|--------|-------------|---------|
| **After pressing Enter, move selection** | Direction | Down |
| **Automatically insert a decimal point** | Fixed decimal | Off |
| **Enable fill handle and cell drag-and-drop** | Drag to fill | On |
| **Alert before overwriting cells** | Confirm overwrite | On |
| **Allow editing directly in cells** | In-cell editing | On |
| **Extend data range formats and formulas** | Auto-extend | On |
| **Enable automatic percent entry** | % when < 1 | On |
| **Enable AutoComplete for cell values** | Suggest completions | On |
| **Flash Fill automatically** | Auto Flash Fill | On |
| **Zoom on roll with IntelliMouse** | Scroll = zoom | Off |
| **Alert the user when a potentially time consuming operation occurs** | Long operation warning | On |
| **Use system separators** | Decimal and thousands | On |
| **Decimal separator** | Decimal point | . |
| **Thousands separator** | Thousands comma | , |
| **Cursor movement** | Logical or Visual | Logical |

### Cut, copy, and paste
| Option | Description | Default |
|--------|-------------|---------|
| **Show Paste Options button** | Smart tag after paste | On |
| **Show Insert Options buttons** | Smart tag after insert | On |
| **Cut, copy, and sort inserted objects with their parent cells** | Move with cells | On |

### Image Size and Quality
| Option | Description | Default |
|--------|-------------|---------|
| **Discard editing data** | Remove undo history | Off |
| **Do not compress images in file** | Keep full resolution | Off |
| **Set default target output** | PPI setting | 220 ppi |

### Chart
| Option | Description | Default |
|--------|-------------|---------|
| **Show chart element names on hover** | Chart ScreenTips | On |
| **Show data point values on hover** | Value ScreenTips | On |
| **Properties follow chart data point** | Formatting follows data | On |
| **Current workbook** | Apply settings to | (current) |
| **For new workbooks, use this chart template** | Default chart | (none) |

### Display
| Option | Description | Default |
|--------|-------------|---------|
| **Show this number of Recent Documents** | Recent file count | 50 |
| **Quickly access this number of Recent Workbooks** | Pinned count | 0 |
| **Show this number of unpinned Recent Folders** | Folder count | 5 |
| **Ruler units** | Inches, Centimeters, Millimeters | Inches |
| **Show formula bar in Editing view** | Show formula bar | On |
| **Show function ScreenTips** | Help tooltips | On |
| **For cells with comments, show** | Indicator or comment | Indicators only |
| **Default direction** | Left-to-right or Right-to-left | Left-to-right |

### Display options for this workbook
| Option | Description | Default |
|--------|-------------|---------|
| **Show horizontal scroll bar** | Bottom scroll bar | On |
| **Show vertical scroll bar** | Right scroll bar | On |
| **Show sheet tabs** | Tab bar | On |
| **Group dates in the AutoFilter menu** | Date grouping | On |
| **For objects, show** | All, Placeholders, Nothing | All |

### Display options for this worksheet
| Option | Description | Default |
|--------|-------------|---------|
| **Show row and column headers** | A, B, C and 1, 2, 3 | On |
| **Show formulas in cells** | Show formulas | Off |
| **Show page breaks** | Page break lines | On |
| **Show a zero in cells that have zero value** | Display zeros | On |
| **Show outline symbols** | +/- buttons | On |
| **Show gridlines** | Cell grid | On |
| **Gridline color** | Grid line color | Automatic |

### Formulas
| Option | Description | Default |
|--------|-------------|---------|
| **Enable multi-threaded calculation** | Use multiple cores | On |
| **Number of calculation threads** | Core count | All processors |

### When calculating this workbook
| Option | Description | Default |
|--------|-------------|---------|
| **Update links to other documents** | Refresh links | On |
| **Save external link values** | Store linked data | On |

### General
| Option | Description | Default |
|--------|-------------|---------|
| **Ignore other applications that use DDE** | DDE behavior | Off |
| **Ask to update automatic links** | Prompt for links | On |
| **Show add-in user interface errors** | Show add-in errors | Off |
| **Scale content for A4 or 8.5" x 11" paper sizes** | Auto-scale for paper | On |
| **At startup, open all files in** | Startup folder | (blank) |
| **Web Options...** | Web publishing settings | Button |
| **Service Options...** | Online services | Button |

---

## Customize Ribbon

### Customize the Ribbon
| Panel | Description |
|-------|-------------|
| **Left** | Available commands |
| **Right** | Current ribbon tabs |

#### Command Sources
| Source | Description |
|--------|-------------|
| **Popular Commands** | Frequently used |
| **Commands Not in the Ribbon** | Unlisted commands |
| **All Commands** | Everything |
| **Macros** | Your macros |
| **All Tabs** | From all tabs |
| **(specific tab)** | From one tab |

#### Ribbon Customization
| Action | Description |
|--------|-------------|
| **New Tab** | Create custom tab |
| **New Group** | Create custom group |
| **Rename** | Change tab/group name |
| **Add >>** | Add command to ribbon |
| **<< Remove** | Remove command |
| **Move Up/Down** | Reorder items |

### Main Tabs Checkboxes
| Tab | Default |
|-----|---------|
| **Home** | Checked |
| **Insert** | Checked |
| **Page Layout** | Checked |
| **Formulas** | Checked |
| **Data** | Checked |
| **Review** | Checked |
| **View** | Checked |
| **Developer** | Unchecked |
| **Help** | Checked |

| Button | Description |
|--------|-------------|
| **Reset** | Restore defaults |
| **Import/Export** | Save/load customizations |

---

## Quick Access Toolbar

### Customize Quick Access Toolbar
| Panel | Description |
|-------|-------------|
| **Left** | Available commands |
| **Right** | Current QAT commands |

| Option | Description |
|--------|-------------|
| **Choose commands from** | Command source |
| **Customize Quick Access Toolbar** | All docs or this workbook |
| **Show Quick Access Toolbar below the Ribbon** | Position toggle |

### Default QAT Commands
| Command | Description |
|---------|-------------|
| **AutoSave** | Cloud auto-save toggle |
| **Save** | Save workbook |
| **Undo** | Undo last action |
| **Redo** | Redo last action |

---

## Add-ins

### View and manage Office Add-ins
| Column | Description |
|--------|-------------|
| **Name** | Add-in name |
| **Location** | File path |
| **Type** | COM, Excel, etc. |

### Manage dropdown
| Option | Description |
|--------|-------------|
| **Excel Add-ins** | XLA/XLAM files |
| **COM Add-ins** | COM components |
| **Actions** | Legacy Actions |
| **XML Expansion Packs** | XML schemas |
| **Disabled Items** | Crashed add-ins |

| Button | Description |
|--------|-------------|
| **Go...** | Open selected add-in manager |

---

## Trust Center

### Microsoft Excel Trust Center
| Button | Description |
|--------|-------------|
| **Trust Center Settings...** | Open Trust Center |

### Trust Center Categories

#### Trusted Publishers
| Column | Description |
|--------|-------------|
| **Name** | Certificate name |
| **Issued By** | Certificate authority |
| **Expiration** | Valid until |
| **Issued To** | Certificate owner |

#### Trusted Locations
| Column | Description |
|--------|-------------|
| **Path** | Trusted folder path |
| **User/All** | User-specific or all users |
| **Description** | Location description |

| Option | Description | Default |
|--------|-------------|---------|
| **Allow Trusted Locations on my network** | Network trust | Off |
| **Disable all Trusted Locations** | Turn off feature | Off |

#### Trusted Documents
| Option | Description | Default |
|--------|-------------|---------|
| **Allow documents on a network to be trusted** | Network trust | On |
| **Disable Trusted Documents** | Turn off feature | Off |
| **Clear** | Reset trusted docs | Button |

#### Trusted Add-in Catalogs
| Option | Description |
|--------|-------------|
| **Catalog URLs** | Web add-in sources |
| **Require apps be signed by Trusted Publisher** | Signature required |

#### Add-ins
| Option | Description | Default |
|--------|-------------|---------|
| **Require Application Add-ins to be signed** | Signature required | Off |
| **Disable notification for unsigned add-ins** | Silent block | Off |
| **Disable all Application Add-ins** | Block all | Off |

#### ActiveX Settings
| Option | Description | Default |
|--------|-------------|---------|
| **Disable all controls without notification** | Silent block | - |
| **Prompt me before enabling UFI controls** | Prompt for unsafe | Selected |
| **Prompt me before enabling all controls** | Prompt for all | - |
| **Enable all controls without restrictions** | Allow all | - |
| **Safe Mode** | Safe initialization | On |

#### Macro Settings
| Option | Description | Default |
|--------|-------------|---------|
| **Disable VBA macros without notification** | Silent block | - |
| **Disable VBA macros with notification** | Prompt user | Selected |
| **Disable VBA macros except digitally signed** | Require signature | - |
| **Enable VBA macros** | Allow all | - |
| **Trust access to the VBA project object model** | Automation access | Off |

#### Protected View
| Option | Description | Default |
|--------|-------------|---------|
| **Enable for files from the Internet** | Web files | On |
| **Enable for files in potentially unsafe locations** | Unsafe paths | On |
| **Enable for Outlook attachments** | Email attachments | On |

#### Message Bar
| Option | Description | Default |
|--------|-------------|---------|
| **Show the Message Bar when content is blocked** | Show security bar | On |

#### External Content
| Option | Description | Default |
|--------|-------------|---------|
| **Data Connections security** | Block, Prompt, or Enable | Prompt |
| **Workbook Links security** | Block, Prompt, or Enable | Prompt |
| **Dynamic Data Exchange security** | DDE behavior | Prompt |

#### File Block Settings
| File Type | Open | Save | Open in Protected View |
|-----------|------|------|------------------------|
| **Excel 2007 and later** | - | - | - |
| **Excel 97-2003** | - | - | - |
| **Excel 4** | Block | Block | - |
| **Excel 3** | Block | Block | - |
| **Excel 2** | Block | Block | - |
| **dBase files** | - | - | - |

#### Privacy Options
| Option | Description | Default |
|--------|-------------|---------|
| **Allow the Office Connected Experiences** | Cloud features | On |
| **Optional connected experiences** | Additional services | On |
| **Document-specific settings** | Inspector settings | Various |

---

## Keyboard Shortcut

| Action | Shortcut |
|--------|----------|
| Open Options | Alt+F, T |

---

[🎗️ Back to Ribbon Reference](./README.md) | [🏠 Back to Home](../README.md)
