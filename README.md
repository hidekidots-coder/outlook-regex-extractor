# Outlook Logistics Document Extractor

VBA automation for extracting AWB and HAWB references from
Microsoft Outlook emails and organizing the extracted information
in Excel.

## Overview

This project automates part of a logistics document processing
workflow using Microsoft Outlook, VBA and Regular Expressions.

The automation searches Outlook emails, analyzes the email subject
and HTML content, identifies AWB and HAWB references, downloads
attachments and stores the extracted information in Excel.

## Problem

Logistics and import/export operations often require users to
manually search through emails, identify shipment references,
download attachments and organize information.

This process can be repetitive and time-consuming.

## Solution

This project automates the extraction and organization of
logistics references from Outlook emails.

### Workflow

```text
Microsoft Outlook
        │
        ▼
   Email Search
        │
        ▼
Subject / HTML Body
        │
        ▼
Regular Expressions
        │
        ▼
AWB / HAWB Extraction
        │
        ▼
Attachment Download
        │
        ▼
      Excel
## How to Use

1. Open Excel and press `Alt + F11`
2. Import the file `src/ExtrairAWBs.bas`
3. Update the configuration constants at the top of the `ExtrairAWBs` procedure
4. Run the macro `ExtrairAWBs`

## Configuration

| Constant           | Description                          | Example                     |
|--------------------|--------------------------------------|-----------------------------|
| `ACCOUNT_NAME`     | Outlook account name                 | `"seu.email@empresa.com"`   |
| `ROOT_FOLDER`      | Main folder                          | `"dhl"`                     |
| `FOLDER_1`         | First subfolder                      | `"notificação de embarque"` |
| `FOLDER_2`         | Second subfolder                     | `"pre alerta carga"`        |
| `DOWNLOAD_FOLDER`  | Path to save attachments             | `"C:\Temp\AWB\"`            |
