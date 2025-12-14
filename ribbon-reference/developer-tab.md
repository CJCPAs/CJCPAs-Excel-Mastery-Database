# Developer Tab - Complete Reference

> **VBA, macros, add-ins, form controls, and XML tools**

---

## Enabling the Developer Tab

The Developer tab is hidden by default. To show it:

1. **File → Options → Customize Ribbon**
2. Check **Developer** in the right panel
3. Click **OK**

---

## Tab Overview

The Developer tab provides programming and advanced features:

| Group | Purpose |
|-------|---------|
| [Code](#code-group) | VBA, macros, recording |
| [Add-ins](#add-ins-group) | COM and Office add-ins |
| [Controls](#controls-group) | Form and ActiveX controls |
| [XML](#xml-group) | XML mapping and tools |

---

## Code Group

### Visual Basic
| Item | Description | Shortcut |
|------|-------------|----------|
| **Visual Basic** | Open VBA Editor | Alt+F11 |

#### VBA Editor Features
| Window | Purpose |
|--------|---------|
| **Project Explorer** | Navigate workbook objects |
| **Properties Window** | Object properties |
| **Code Window** | Write/edit VBA code |
| **Immediate Window** | Test code (Ctrl+G) |
| **Watch Window** | Debug variables |
| **Locals Window** | View local variables |

#### VBA Editor Shortcuts
| Action | Shortcut |
|--------|----------|
| Run code | F5 |
| Step into | F8 |
| Step over | Shift+F8 |
| Toggle breakpoint | F9 |
| Stop execution | Ctrl+Break |
| Object browser | F2 |
| Find | Ctrl+F |
| Replace | Ctrl+H |

### Macros
| Item | Description | Shortcut |
|------|-------------|----------|
| **Macros** | View/run macros | Alt+F8 |

#### Macros Dialog
| Button | Description |
|--------|-------------|
| **Run** | Execute selected macro |
| **Step Into** | Start debugging |
| **Edit** | Open in VBA Editor |
| **Create** | Create new macro |
| **Delete** | Remove macro |
| **Options...** | Shortcut and description |

### Record Macro
| Item | Description | Shortcut |
|------|-------------|----------|
| **Record Macro** | Start macro recording | - |

#### Record Macro Dialog
| Field | Description |
|-------|-------------|
| **Macro name** | Name (no spaces) |
| **Shortcut key** | Ctrl+[letter] |
| **Store macro in** | Workbook location |
| **Description** | What macro does |

#### Storage Options
| Location | Description |
|----------|-------------|
| **This Workbook** | Saved with current file |
| **New Workbook** | New file for macro |
| **Personal Macro Workbook** | PERSONAL.XLSB (always available) |

### Use Relative References
| Item | Description | Shortcut |
|------|-------------|----------|
| **Use Relative References** | Record relative cell moves | - |

#### Absolute vs Relative Recording
| Mode | Result |
|------|--------|
| **Absolute (default)** | Records exact cell addresses (e.g., A1) |
| **Relative** | Records offsets (e.g., one cell down) |

### Macro Security
| Item | Description | Shortcut |
|------|-------------|----------|
| **Macro Security** | Trust Center macro settings | - |

#### Macro Security Levels
| Level | Description |
|-------|-------------|
| **Disable all macros without notification** | Silent block |
| **Disable all macros with notification** | Warning prompt |
| **Disable all macros except digitally signed macros** | Require signature |
| **Enable all macros** | No security (not recommended) |

#### Trusted Locations
| Feature | Description |
|---------|-------------|
| **Add folder** | Trust all files in folder |
| **Subfolders** | Include nested folders |
| **Network** | Allow network locations |

---

## Add-ins Group

### Add-ins
| Item | Description | Shortcut |
|------|-------------|----------|
| **Add-ins** | Manage Excel add-ins | - |

#### Add-ins Button Options
| Category | Description |
|----------|-------------|
| **My Add-ins** | Your installed add-ins |
| **Office Add-ins** | From Office Store |

### Excel Add-ins
| Item | Description | Shortcut |
|------|-------------|----------|
| **Excel Add-ins** | XLA/XLAM add-ins | - |

#### Add-ins Dialog
| Button | Description |
|--------|-------------|
| **Browse...** | Find add-in file |
| **Automation...** | COM add-ins |

#### Built-in Add-ins
| Add-in | Purpose |
|--------|---------|
| **Analysis ToolPak** | Statistical analysis |
| **Analysis ToolPak - VBA** | VBA functions for analysis |
| **Euro Currency Tools** | Euro conversion |
| **Solver Add-in** | Optimization solver |

### COM Add-ins
| Item | Description | Shortcut |
|------|-------------|----------|
| **COM Add-ins** | .NET and COM components | - |

#### COM Add-ins Dialog
| Column | Description |
|--------|-------------|
| **Available Add-ins** | Installed COM add-ins |
| **Load Behavior** | How add-in loads |
| **Location** | File path |

| Load Behavior | Description |
|---------------|-------------|
| **Load at Startup** | Always load |
| **Load on Demand** | Load when used |
| **Unloaded** | Not loaded |

---

## Controls Group

### Insert
| Item | Description | Shortcut |
|------|-------------|----------|
| **Insert** | Add form controls | - |

#### Form Controls
| Control | Purpose |
|---------|---------|
| **Button** | Run macro on click |
| **Combo Box** | Dropdown list |
| **Check Box** | Yes/No option |
| **Spin Button** | Increment/decrement value |
| **List Box** | Select from list |
| **Option Button** | Mutually exclusive choice |
| **Group Box** | Group related controls |
| **Label** | Display text |
| **Scroll Bar** | Select from range |

#### ActiveX Controls
| Control | Purpose |
|---------|---------|
| **Command Button** | Programmable button |
| **Combo Box** | Enhanced dropdown |
| **Check Box** | Enhanced checkbox |
| **Spin Button** | Enhanced spinner |
| **List Box** | Enhanced list |
| **Option Button** | Enhanced radio button |
| **Toggle Button** | On/off button |
| **Label** | Programmable label |
| **Text Box** | Text input |
| **Scroll Bar** | Enhanced scroll bar |
| **Image** | Display picture |
| **More Controls...** | Additional ActiveX |

#### Form vs ActiveX Controls
| Feature | Form Controls | ActiveX Controls |
|---------|---------------|------------------|
| **Ease of use** | Simpler | More complex |
| **Formatting** | Limited | Extensive |
| **Events** | Macro assignment | VBA events |
| **Compatibility** | Better | Can have issues |
| **Web** | Works | May not work |

### Design Mode
| Item | Description | Shortcut |
|------|-------------|----------|
| **Design Mode** | Toggle edit mode for ActiveX | - |

#### Design Mode Features
| Mode | Description |
|------|-------------|
| **On** | Edit controls, view properties |
| **Off** | Controls are functional |

### Properties
| Item | Description | Shortcut |
|------|-------------|----------|
| **Properties** | Control properties dialog | - |

#### Common Properties
| Property | Description |
|----------|-------------|
| **Name** | Control identifier |
| **Caption** | Display text |
| **Value** | Current value |
| **LinkedCell** | Cell for value |
| **ListFillRange** | Data source |
| **Font** | Text formatting |
| **BackColor** | Background color |
| **ForeColor** | Text color |
| **Enabled** | Can interact |
| **Visible** | Can see |

### View Code
| Item | Description | Shortcut |
|------|-------------|----------|
| **View Code** | Open control's VBA code | - |

### Run Dialog
| Item | Description | Shortcut |
|------|-------------|----------|
| **Run Dialog** | Display custom dialog | - |

---

## XML Group

### Source
| Item | Description | Shortcut |
|------|-------------|----------|
| **Source** | XML Source task pane | - |

#### XML Source Pane
| Feature | Description |
|---------|-------------|
| **XML Maps** | Loaded schema mappings |
| **Drag elements** | Map to cells |
| **Properties** | Element settings |

### Map Properties
| Item | Description | Shortcut |
|------|-------------|----------|
| **Map Properties** | XML map settings | - |

#### Map Properties Options
| Option | Description |
|--------|-------------|
| **Validate data** | Check against schema |
| **Preserve formatting** | Keep cell formats |
| **Adjust column width** | Auto-fit on import |

### Expansion Packs
| Item | Description | Shortcut |
|------|-------------|----------|
| **Expansion Packs** | Office XML schemas | - |

### Refresh Data
| Item | Description | Shortcut |
|------|-------------|----------|
| **Refresh Data** | Reload XML data | - |

### Import
| Item | Description | Shortcut |
|------|-------------|----------|
| **Import** | Import XML file | - |

#### Import Options
| Option | Description |
|--------|-------------|
| **XML table** | As structured table |
| **Existing worksheet** | Into current sheet |
| **New worksheet** | New sheet |

### Export
| Item | Description | Shortcut |
|------|-------------|----------|
| **Export** | Export to XML file | - |

---

## Personal Macro Workbook

### About PERSONAL.XLSB
| Feature | Description |
|---------|-------------|
| **Location** | XLSTART folder |
| **Always loaded** | Opens with Excel (hidden) |
| **Global macros** | Available in all workbooks |
| **Persists** | Survives workbook close |

### Creating PERSONAL.XLSB
1. Record a macro
2. Choose "Personal Macro Workbook" for storage
3. File is created automatically
4. Access via VBA Editor

### PERSONAL.XLSB Location
| OS | Path |
|----|------|
| **Windows** | C:\Users\[Username]\AppData\Roaming\Microsoft\Excel\XLSTART |
| **Mac** | /Users/[Username]/Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Excel |

---

## Common VBA Tasks

### Assign Macro to Button
1. Insert → Form Controls → Button
2. Draw button on sheet
3. Select macro in dialog
4. Click OK

### Edit Button Macro
1. Right-click button
2. Select "Assign Macro..."
3. Choose different macro or Edit

### Remove Control
1. For Form Controls: Select and Delete
2. For ActiveX: Enter Design Mode first

---

## Keyboard Shortcuts Summary

| Action | Shortcut |
|--------|----------|
| Open VBA Editor | Alt+F11 |
| Macros dialog | Alt+F8 |
| Run macro | F5 (in VBA Editor) |
| Step through code | F8 |
| Toggle breakpoint | F9 |
| Stop code | Ctrl+Break |
| Object Browser | F2 (in VBA Editor) |
| Immediate Window | Ctrl+G (in VBA Editor) |
| Close VBA Editor | Alt+Q |

---

## Security Best Practices

| Practice | Description |
|----------|-------------|
| **Disable by default** | Use "with notification" setting |
| **Trusted locations** | Only for known-safe folders |
| **Digital signatures** | Sign your own macros |
| **Review code** | Check macros before running |
| **Backup** | Save before running unknown macros |

---

[🎗️ Back to Ribbon Reference](./README.md) | [🏠 Back to Home](../README.md)
