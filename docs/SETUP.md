# Acquisition Workspace Setup Guide

Complete setup guide for the Real Estate Acquisition Tracking System.

## Table of Contents

1. [Quick Start](#quick-start)
2. [Installing the Apps Script](#installing-the-apps-script)
3. [Initializing Your Workspace](#initializing-your-workspace)
4. [Configuring Monthly Goals](#configuring-monthly-goals)
5. [Setting Up Automation Triggers](#setting-up-automation-triggers)
6. [Importing Zip Codes](#importing-zip-codes)
7. [Using the System](#using-the-system)

---

## Quick Start

### Step 1: Open Your Google Sheet

Open your existing Google Sheet or create a new one:
- Go to [Google Sheets](https://sheets.google.com)
- Open your Acquisition Workspace or create new

### Step 2: Open Apps Script Editor

1. Click **Extensions** > **Apps Script**
2. Delete any existing code in the editor
3. Copy the contents of `google-apps-script/Code.gs`
4. Paste into the Apps Script editor
5. Create additional files for `Triggers.gs` and `Utilities.gs`
6. Click the **Save** icon (or Ctrl+S)

### Step 3: Initialize the Workspace

1. Refresh your Google Sheet (the menu might take a moment to appear)
2. Click **Acquisition Tools** > **Setup** > **Initialize All Sheets**
3. Click **Yes** when prompted
4. Wait for the initialization to complete

---

## Installing the Apps Script

### Creating the Script Files

1. In Apps Script editor, you'll see a file called `Code.gs`
2. Click the **+** next to "Files" to create new files
3. Create these files and paste the corresponding content:

| File Name | Source File |
|-----------|-------------|
| Code.gs | `google-apps-script/Code.gs` |
| Triggers.gs | `google-apps-script/Triggers.gs` |
| Utilities.gs | `google-apps-script/Utilities.gs` |

### Authorizing the Script

When you first run any function, Google will ask for permissions:

1. Click **Review Permissions**
2. Select your Google account
3. Click **Advanced** > **Go to Acquisition Workspace (unsafe)**
4. Click **Allow**

This grants the script permission to:
- Edit your spreadsheet
- Send emails (for reports)
- Create time-based triggers

---

## Initializing Your Workspace

### What Gets Created

Running "Initialize All Sheets" creates these sheets:

| Sheet | Purpose |
|-------|---------|
| **On Market Leads** | Track Zillow, Redfin, MLS leads |
| **Off Market Leads** | Track wholesaler, agent, direct mail leads |
| **Contract Pipeline** | Track deals under contract |
| **Main Page** | Consolidated view of all leads |
| **Lead KPIs** | Monthly performance metrics |
| **Zip Codes** | Target markets and performance |
| **Dashboard** | Visual overview and goal tracking |
| **Buyers List** | Cash buyer database |

### Data Validation

The script automatically sets up:
- Dropdown menus for Status, Source, Yes/No fields
- Conditional formatting for status colors
- Auto-calculated fields (MAO, profit, etc.)

---

## Configuring Monthly Goals

### Default Goals

The system comes with these default monthly goals:

```javascript
MONTHLY_GOALS: {
  LEADS_GENERATED: 100,
  LEADS_QUALIFIED: 50,
  OFFERS_MADE: 25,
  CONTRACTS_SIGNED: 5,
  DEALS_CLOSED: 3,
  REVENUE_TARGET: 50000
}
```

### Customizing Goals

1. Open Apps Script editor (Extensions > Apps Script)
2. In `Code.gs`, find the `CONFIG` object at the top
3. Modify the `MONTHLY_GOALS` values
4. Save the script
5. Run **Acquisition Tools** > **Update KPIs** to refresh

---

## Setting Up Automation Triggers

### Available Automations

| Trigger | Frequency | Description |
|---------|-----------|-------------|
| Daily Morning Report | 8 AM daily | Summary of yesterday's activity |
| Weekly KPI Report | Monday 9 AM | Full KPI analysis |
| Monthly Reset | 1st of month | Archive KPIs, start fresh |
| Archive Old Leads | Sunday 2 AM | Move 90+ day dead leads |

### Setting Up Triggers

**Option 1: Automatic Setup**
1. In Apps Script editor
2. Run the function `setupAllTriggers()`
3. All recommended triggers will be created

**Option 2: Manual Setup**
1. In Apps Script editor, click the **clock icon** (Triggers)
2. Click **+ Add Trigger**
3. Configure:
   - Function: `dailyMorningReport`
   - Event source: Time-driven
   - Type: Day timer
   - Time: 8am to 9am
4. Repeat for other functions

### Email Notifications

To enable email reports:
1. Open `Triggers.gs`
2. Find the `dailyMorningReport` function
3. Uncomment the `MailApp.sendEmail` section
4. Replace `your-email@example.com` with your email

---

## Importing Zip Codes

### Using the Import Tool

1. Click **Acquisition Tools** > **Setup** > **Import Zip Codes**
2. Enter zip codes (one per line or comma-separated)
3. Click **Import**

### Bulk Import from CSV

1. Open the `Zip Codes` sheet
2. Copy data from `templates/sample-zip-codes.csv`
3. Paste starting from row 2

### Required Columns

| Column | Description |
|--------|-------------|
| Zip Code | 5-digit zip code |
| City | City name |
| County | County name |
| State | State abbreviation |
| Market Status | Active/Paused/Retired |
| Priority Level | High/Medium/Low |

---

## Using the System

### Adding Leads

**Method 1: Quick Entry Forms**
1. Click **Acquisition Tools** > **Add New Lead**
2. Select **On Market Lead** or **Off Market Lead**
3. Fill out the form and submit

**Method 2: Direct Entry**
1. Go to the appropriate sheet
2. Start typing in the first empty row
3. Lead ID and Date are auto-generated

### Lead Workflow

```
New → Contacted → Analyzing → Offer Pending → Under Contract → Closed
                                                    ↓
                                                  Dead
```

### Qualifying Leads

1. Find the lead in Main Page or source sheet
2. Set **Qualified (Yes/No)** to **Yes**
3. The lead will now count in qualified metrics

### Moving to Contract Pipeline

**Automatic:** Set "Contract Signed" to "Yes" - lead auto-moves to pipeline

**Manual:**
1. Select the lead row
2. Click **Acquisition Tools** > **Move to Contract Pipeline**

### Tracking Deals

In Contract Pipeline, update:
- Inspection Period End date
- Close Date
- Disposition Strategy
- Buyer information
- Final Sale Price (when closed)
- Status (Under Contract → Pending Close → Closed)

### Viewing KPIs

- **Lead KPIs sheet**: Detailed metrics and conversion rates
- **Dashboard**: Visual overview with goal progress
- Click **Refresh Dashboard** to update all calculations

---

## Custom Formulas

The system adds custom functions you can use in cells:

| Formula | Description | Example |
|---------|-------------|---------|
| `=MAO(ARV, Repairs, Fee)` | Calculate Maximum Allowable Offer | `=MAO(200000, 30000, 10000)` |
| `=POTENTIAL_PROFIT(...)` | Calculate flip profit potential | See Utilities.gs |
| `=ASSIGNMENT_FEE(Contract, EndBuyer)` | Calculate assignment fee | `=ASSIGNMENT_FEE(150000, 165000)` |
| `=FORMAT_PHONE(phone)` | Format phone number | `=FORMAT_PHONE("5551234567")` |
| `=DAYS_UNTIL(date)` | Days until a deadline | `=DAYS_UNTIL(O2)` |
| `=LEAD_SCORE(equity, motivation, days, condition)` | Score a lead 1-100 | See docs |
| `=MATCHING_BUYERS(zip, type, price)` | Find matching buyers | `=MATCHING_BUYERS("75001", "SFH", 200000)` |

---

## Troubleshooting

### Menu Not Appearing

1. Refresh the spreadsheet
2. Wait 30 seconds
3. If still missing, go to Extensions > Apps Script > Run > onOpen

### Permission Errors

1. Go to Apps Script editor
2. Run any function manually
3. Complete the authorization flow

### Formulas Not Calculating

1. Click **Acquisition Tools** > **Refresh Dashboard**
2. Or press Ctrl+Shift+E to force recalculation

### Triggers Not Running

1. Check Apps Script > Triggers
2. Look for failed executions
3. Verify your email/permissions

---

## Support

For issues or feature requests:
1. Check the Help dialog: **Acquisition Tools** > **Help**
2. Review this documentation
3. Check the Apps Script execution log for errors

---

## Version History

- **v3.1** - Current version with full automation
- **v3.0** - Added KPI tracking and Dashboard
- **v2.0** - Added Contract Pipeline
- **v1.0** - Initial release with basic lead tracking
