#   CSV Editor - Complete Documentation

## Overview
The   CSV Editor is a specialized spreadsheet application designed for editing CSV, TSV, and TXT files with advanced features tailored for turbine component data management. Built with Python using tkinter, tksheet, and ttkbootstrap libraries.

---

## Core Features

### 1. File Operations
- **New File Creation**: Start with a blank spreadsheet containing default columns
- **File Opening**: Support for multiple file formats:
  - CSV (Comma-separated values)
  - TSV (Tab-separated values) 
  - TXT (Text files with various delimiters)
- **Save Functions**: 
  - Save to current file
  - Save As with format selection
- **Unsaved Changes Protection**: Prompts user before discarding modifications

### 2. User Interface
- **Modern Bootstrap Theme**: Clean, professional interface using ttkbootstrap
- **Dark Mode Toggle**: Switch between light ("flatly") and dark ("darkly") themes
- **Responsive Layout**: Automatically adjusts to window size
- **Status Bar**: Shows current file status and copyright information
- **Toolbar**: Quick access to all major functions

### 3. Spreadsheet Functionality
- **Full Sheet Editing**: Edit cells, headers, rows, and columns
- **Selection Modes**:
  - Single cell selection
  - Row selection
  - Column selection
  - Drag selection
- **Data Operations**:
  - Copy, cut, paste
  - Undo/redo functionality
  - Find and replace
  - Sort columns and rows
- **Structural Operations**:
  - Insert/delete rows and columns
  - Resize row heights and column widths
  - Right-click context menus

### 4. Suggestion Mode
A specialized feature for turbine component data validation and assistance.

### 5. Data Validation
- **Visual Feedback**: Invalid entries highlighted in red when not matching dropdown options
- **Input Validation**: Dropdown constraints ensure data consistency
- **Flexible Input**: Allows custom values while highlighting non-standard entries

### 6. Keyboard Shortcuts
- **Ctrl+N**: New file
- **Ctrl+O**: Open file
- **Ctrl+S**: Save file
- **Ctrl+Shift+S**: Save As
- **Ctrl+Z**: Undo
- **Ctrl+Y**: Redo

### 7. Theme Management
- **Automatic Color Coordination**: All UI elements update when switching themes
- **Sheet Styling**: Spreadsheet colors automatically adjust to match theme
- **Status Bar Integration**: Footer elements follow theme colors
- **Persistent Settings**: Theme preference maintained during session

---

## Technical Specifications

### Dependencies
- **Python 3.x**
- **pandas**: Data manipulation and CSV handling
- **tkinter**: GUI framework
- **tksheet**: Advanced spreadsheet widget
- **ttkbootstrap**: Modern styling framework
- **webbrowser**: Help documentation access

### File Format Support
- **CSV**: Standard comma-separated values
- **TSV**: Tab-separated values  
- **TXT**: Space-delimited text files
- **Encoding**: Handles various text encodings automatically
- **Empty Values**: Preserves empty cells without converting to NaN

### Data Handling
- **Memory Management**: Efficient handling of large datasets
- **Type Preservation**: Maintains original data types where possible
- **Error Recovery**: Graceful handling of malformed files

---

## Advanced Features

### 1. Context-Aware Dropdowns
- **Dynamic Population**: Parasolid dropdown automatically scans directory for relevant files
- **Real-time Updates**: Dropdowns update when suggestion mode is toggled
- **Value Preservation**: Existing cell values maintained when dropdowns are applied

### 2. Visual Feedback System
- **Color-coded Validation**: Invalid entries highlighted in contrasting colors
- **Status Updates**: Real-time feedback on operations
- **Modification Tracking**: Visual indicators for unsaved changes

### 3. Help Integration
- **Web-based Documentation**: Direct links to SharePoint documentation
- **Contextual Assistance**: Help button provides access to detailed guides

### 4. Session Management
- **Auto-save Prompts**: Prevents accidental data loss
- **State Preservation**: Maintains editing state during theme changes
- **Clean Exit**: Proper cleanup on application close

---

## Usage Guidelines

### Getting Started
1. Launch the application
2. Use "New File" to start fresh or "Open File" to load existing data
3. Enable "Suggestion Mode" for turbine-specific data validation
4. Toggle "Dark Mode" for preferred visual theme

### Best Practices
- **Save Frequently**: Use Ctrl+S to save changes regularly
- **Use Suggestion Mode**: Enable for turbine component files to ensure data validity
- **Check Validation**: Red-highlighted cells indicate values outside expected ranges
- **Backup Important Files**: Use "Save As" to create backup copies

### Troubleshooting
- **File Opening Issues**: Try different delimiter options if initial load fails
- **Dropdown Problems**: Toggle suggestion mode off and on to refresh dropdowns
- **Theme Issues**: Restart application if visual elements don't update properly

---

## File Naming Conventions

The application recognizes specific file naming patterns:


These patterns trigger specialized dropdown menus and validation rules.

---

## Copyright & Support

© 2025 Pramod Kumar
