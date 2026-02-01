# Keystone Acquisition Workspace

A comprehensive Google Sheets-based real estate acquisition tracking system with automated KPIs, lead management, and deal pipeline tracking.

## Features

### Lead Tracking
- **On-Market Leads**: Track properties from Zillow, Redfin, Realtor.com, and MLS
- **Off-Market Leads**: Track deals from wholesalers, agents, direct mail, cold calling, driving for dollars, probate, and tax liens
- **Automated Lead IDs**: Auto-generated unique identifiers
- **Status Tracking**: Full workflow from New to Closed

### Contract Pipeline
- Track deals under contract through closing
- Inspection period and close date monitoring
- Disposition strategy tracking (Assignment, Double Close, Wholetail, Fix & Flip, Buy & Hold)
- Automatic profit calculations

### KPI Dashboard
- Monthly performance metrics
- Goal tracking with visual progress indicators
- Conversion rate analysis:
  - Lead to Qualified
  - Qualified to Offer
  - Offer to Contract
  - Contract to Close
- Source performance comparison

### Automation
- Auto-populate Lead IDs and dates
- MAO (Maximum Allowable Offer) calculator using 70% rule
- Auto-move leads to Contract Pipeline when contract signed
- Scheduled reports (daily, weekly, monthly)
- Stale lead detection
- Contract deadline alerts

### Zip Code Management
- Track target markets by zip code
- Performance metrics per zip
- Priority ranking (High/Medium/Low)
- Success rate tracking

### Buyers List
- Cash buyer database
- Buyer matching by zip code, property type, and price range
- Deal history tracking

## Project Structure

```
keystone-acquisition/
├── google-apps-script/
│   ├── Code.gs           # Main script with sheet creation & core functions
│   ├── Triggers.gs       # Time-based automation and triggers
│   └── Utilities.gs      # Helper functions and calculations
├── templates/
│   ├── sheet-structure.json   # Complete sheet schema
│   └── sample-zip-codes.csv   # 120+ DFW zip codes ready to import
├── docs/
│   └── SETUP.md          # Detailed setup instructions
└── README.md             # This file
```

## Quick Start

### 1. Open Google Sheets
Open your existing Acquisition Workspace or create a new Google Sheet.

### 2. Add the Script
1. Go to **Extensions** > **Apps Script**
2. Copy contents of `google-apps-script/Code.gs` into the editor
3. Create additional files for `Triggers.gs` and `Utilities.gs`
4. Save the project

### 3. Initialize
1. Refresh the spreadsheet
2. Click **Acquisition Tools** > **Setup** > **Initialize All Sheets**
3. Authorize when prompted

### 4. Import Zip Codes
1. Click **Acquisition Tools** > **Setup** > **Import Zip Codes**
2. Or paste from `templates/sample-zip-codes.csv`

## Sheet Overview

| Sheet | Purpose |
|-------|---------|
| On Market Leads | Zillow, Redfin, MLS tracking |
| Off Market Leads | Wholesaler, agent, direct mail tracking |
| Contract Pipeline | Deals under contract to close |
| Main Page | Consolidated view of all leads |
| Lead KPIs | Monthly metrics and conversion rates |
| Zip Codes | Target markets and performance |
| Dashboard | Visual overview and goal progress |
| Buyers List | Cash buyer database |

## Default Monthly Goals

| Metric | Goal |
|--------|------|
| Leads Generated | 100 |
| Leads Qualified | 50 |
| Offers Made | 25 |
| Contracts Signed | 5 |
| Deals Closed | 3 |
| Revenue Target | $50,000 |

Customize these in the `CONFIG` object in `Code.gs`.

## Automation Features

### On Edit Triggers
- Auto-generate Lead ID
- Auto-set Date Added
- Auto-calculate MAO
- Auto-move to Contract Pipeline when "Contract Signed" = Yes

### Time-Based Triggers
- **Daily (8 AM)**: Morning activity summary
- **Weekly (Monday 9 AM)**: KPI report
- **Monthly (1st)**: Archive KPIs, reset for new month
- **Weekly (Sunday 2 AM)**: Archive old dead leads

## Custom Formulas

Use these in any cell:

```
=MAO(ARV, Repairs, WholesaleFee)     # Calculate Maximum Allowable Offer
=POTENTIAL_PROFIT(...)               # Calculate flip profit
=ASSIGNMENT_FEE(Contract, EndBuyer)  # Calculate assignment fee
=FORMAT_PHONE(phone)                 # Format (XXX) XXX-XXXX
=DAYS_UNTIL(date)                    # Days until deadline
=LEAD_SCORE(equity, motivation, days, condition)  # Score 1-100
=MATCHING_BUYERS(zip, type, price)   # Find matching buyers
```

## Menu Options

After initialization, find **Acquisition Tools** in the menu:

- **Refresh Dashboard** - Update all calculations
- **Update KPIs** - Recalculate KPI sheet
- **Add New Lead** - Quick entry forms
- **Move to Contract Pipeline** - Manual pipeline move
- **Setup** - Initialize sheets, reset KPIs, import zip codes
- **Help** - Usage documentation

## Workflow

```
1. Lead Entry
   └── On Market (Zillow/Redfin) OR Off Market (Wholesaler/Agent)

2. Lead Qualification
   └── Analyze → Set Qualified = Yes → Pass to Team Lead

3. Offer Stage
   └── Calculate MAO → Submit Offer → Track in system

4. Contract Stage
   └── Contract Signed = Yes → Auto-moves to Contract Pipeline

5. Closing
   └── Track deadlines → Update disposition → Record final sale
```

## Data Sources Supported

### On-Market
- Zillow
- Redfin
- Realtor.com
- MLS

### Off-Market
- Wholesalers
- Agent Referrals
- Direct Mail
- Cold Calling
- Driving for Dollars
- Probate
- Tax Liens

## Customization

### Modify Lead Sources
Edit the `CONFIG.SOURCES` object in `Code.gs`:

```javascript
SOURCES: {
  ON_MARKET: ['Zillow', 'Redfin', 'Realtor.com', 'MLS', 'YourSource'],
  OFF_MARKET: ['Wholesaler', 'Agent Referral', 'Direct Mail', ...]
}
```

### Modify Status Options
Edit the `CONFIG.STATUS` object:

```javascript
STATUS: {
  LEAD: ['New', 'Contacted', 'Analyzing', 'Offer Pending', ...],
  DISPOSITION: ['Assigned', 'Double Close', 'Wholetail', ...]
}
```

### Modify Monthly Goals
Edit the `CONFIG.MONTHLY_GOALS` object:

```javascript
MONTHLY_GOALS: {
  LEADS_GENERATED: 150,
  LEADS_QUALIFIED: 75,
  // ...
}
```

## Included Zip Codes

The `sample-zip-codes.csv` includes 120+ zip codes in the Dallas-Fort Worth metroplex:
- Dallas County
- Collin County
- Denton County
- Tarrant County
- Rockwall County

Each includes city, county, market status, priority level, and notes.

## License

This project is provided as-is for real estate acquisition tracking purposes.

## Support

See `docs/SETUP.md` for detailed setup instructions and troubleshooting.
