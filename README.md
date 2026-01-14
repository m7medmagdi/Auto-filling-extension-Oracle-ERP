# Auto Excel Form Filler

A Chrome browser extension that automatically fills Java Faces form fields from Excel data with persistent state management.

## Features

- **Excel Data Import**: Upload and parse Excel files (.xlsx, .xls) containing form data
- **Persistent State**: Remembers current row position and loaded data across browser sessions
- **Form Filling Options**:
  - Fill Next Row: Fills the current row data into form fields
  - Fill and Submit: Fills data and submits the form
  - Fill and Submit Full: Fills all remaining rows and submits
- **Reset Functionality**: Clear all stored data and reset to initial state
- **Java Faces Support**: Specifically designed for Java Faces web applications

## Installation

1. Download or clone this repository
2. Open Chrome and navigate to `chrome://extensions/`
3. Enable "Developer mode" in the top right corner
4. Click "Load unpacked" and select the extension directory
5. The extension will be installed and ready to use

## Usage

1. Click the extension icon in the Chrome toolbar
2. Select an Excel file containing your form data
3. Navigate to the Java Faces form page you want to fill
4. Use the buttons in the extension popup:
   - **Fill Next Row**: Fills the current row data into the form fields
   - **Fill and Submit**: Fills the data and automatically submits the form
   - **Fill and Submit Full**: Processes all remaining rows automatically
   - **Reset Extension**: Clears all data and resets the row counter

## Excel File Format

The Excel file should contain data in the following column order:
1. Department
2. Cost Center
3. Future1
4. Future2
5. Location Barcode

Each row represents one form entry to be filled.

## Permissions

The extension requires the following permissions:
- `scripting`: To inject scripts into web pages
- `activeTab`: To access the currently active tab
- `tabs`: To manage tab information
- `storage`: To persist Excel data and current row state
- Host permissions for all URLs to work on any website

## Files Structure

- `manifest.json`: Extension manifest and configuration
- `popup.html`: Extension popup interface (in Arabic)
- `popup.js`: Popup functionality and Excel processing
- `background.js`: Background service worker
- `content.js`: Content script for form interaction
- `styles.css`: Popup styling
- `xlsx.full.min.js`: Excel file parsing library

## Development

This extension uses Manifest V3 and is built with vanilla JavaScript. It utilizes the SheetJS library (xlsx.full.min.js) for Excel file processing.

## Author

Developed by Mohamed Magdy for the Information Technology Department.

## Version

Current version: 1.8
