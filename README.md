# VBA Office Bridge: Excel-to-PowerPoint Data Extraction

## Overview
This VBA macro automates the extraction of structured data from external Excel (`.xlsx`) files directly into a PowerPoint (`.pptm`) presentation. It is designed to run from within PowerPoint, silently opening target Excel files, searching for specific data markers, reading dynamic ranges into memory, and bypassing system-generated temporary lock files.

## Business Problem Solved
When weekly reports are generated in Excel by multiple users, pulling that data into presentation slides manually is time-consuming and prone to copy-paste errors. This script automates that workflow by:
1. Scanning the presentation's directory for relevant Excel data files.
2. Ignoring hidden temporary lock files (`~$`) that crash naive loops.
3. Dynamically locating data tables based on string markers (e.g., `ATIVADO~?`) instead of hardcoded cell references.

## Key Features & Techniques
* **Cross-Application Automation:** Uses late binding to create an invisible `Excel.Application` instance from PowerPoint, manipulating workbooks and worksheets in the background.
* **Robust File Handling:** Uses `Scripting.FileSystemObject` to iterate through directories and safely filter out open-file lock markers.
* **Dynamic Range Parsing:** Uses `.Find` methods to locate the upper-left boundary of the target data, counts matches to determine the row depth, and dynamically defines the extraction range (e.g., `ws.Range(ws.Cells(foundRow, foundColumn), ws.Cells(limitRowForSearch, limitColumnForSearch))`).
* **Memory Optimization:** Loads entire Excel ranges into a 2D Variant array (`.Value2`) for high-speed, in-memory processing instead of looping through cells one by one.
* **Data Aggregation:** Includes loop structures designed to aggregate metrics (like totals and quantities) from the 2D array directly into VBScript variables.

## Requirements
* Microsoft PowerPoint (with macros enabled `.pptm`)
* Target data files must be in `.xlsx` format and reside in the same directory as the PowerPoint file.
* Requires Windows OS (relies on `Scripting.FileSystemObject` and ActiveX COM objects).

## Usage Notes
- The sub `CRIAR_ZLIDES_FOTOS_WEEK_ATUAL` is the main execution block.
- Development tracking `MsgBox` alerts are currently active to verify range boundaries and target files during execution.
