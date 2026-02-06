# Keystone Acquisition Workspace v4.0

A streamlined Google Sheets-based real estate acquisition tracking system with automated pipeline management.

## Features

### Simplified Lead Tracking
- **On Market Leads**: Zillow, Redfin, MLS, Crexi, LoopNet
- **Off Market Leads**: Wholesalers, Direct Mail, Expired, Foreclosure, Tired Landlord, Probate, Tax Lien

### Automatic Pipeline
- When "Contract Signed" = Yes, lead automatically moves to Contract Pipeline
- No manual data entry for pipeline - all data transfers automatically

### Team Tracking
- Acquisition Team Member assignment (Tess Walter, Nick Barr)
- Team performance metrics and leaderboard
- Per-member conversion rates

### KPI Dashboard
- Monthly KPIs with goal tracking
- Yearly KPI tracking by month
- Conversion rates: Lead → Offer → Contract → Close
- Source/Issue performance comparison

### Future Integrations
- Gmail integration placeholder
- DealMachine.com integration placeholder

## Project Structure

```
keystone-acquisition/
├── google-apps-script/
│   ├── Code.gs           # Main script - sheets, menus, automation
│   ├── Triggers.gs       # Time-based automation
│   └── Utilities.gs      # Helper functions & custom formulas
├── templates/
│   ├── sheet-structure.json   # Schema reference
│   └── sample-zip-codes.csv   # 120+ DFW zip codes
├── docs/
│   └── SETUP.md          # Setup guide
└── README.md
```

## Sheets Created

| Sheet | Purpose |
|-------|---------|
| On Market Leads | Zillow, Redfin, MLS, Crexi, LoopNet tracking |
| Off Market Leads | Wholesaler, Direct Mail, Expired, Foreclosure leads |
| Contract Pipeline | Auto-populated from lead sheets |
| Lead KPIs | Monthly metrics |
| Yearly KPIs | Track each month's KPIs for the year |
| Zip Codes | Target markets with performance stats |
| Dashboard | Visual overview |
| Buyers List | Import from Dispo Workspace 3.0 |

## Column Structure

### On Market Leads
- Lead ID (auto)
- Date Added (auto)
- Acquisition Team Member
- Source (Zillow, Redfin, MLS, Crexi, LoopNet)
- MLS #, Address, City, County, State, Zip
- Property details (Type, Beds, Baths, Sqft, Year Built)
- List Price, Days on Market
- ARV, Repair Estimate, MAO (auto-calculated)
- Offer Amount
- Offer Made (Yes/No)
- **Contract Signed (Yes/No)** ← Triggers pipeline move
- Notes, Agent info, URL

### Off Market Leads
- Lead ID (auto)
- Date Added (auto)
- Acquisition Team Member
- **Issue** (Wholesaler, Direct Mail, Expired, Foreclosure, Tired Landlord, etc.)
- Address, City, County, State, Zip
- Property details
- Asking Price, ARV, Repair Estimate, MAO (auto)
- Offer Amount, Assignment Fee
- Offer Made (Yes/No)
- **Contract Signed (Yes/No)** ← Triggers pipeline move
- Motivation Level, Seller info, Notes

## Automation

### On Edit Triggers
- Auto-generate Lead ID and Date
- Auto-calculate MAO (70% rule)
- Auto-move to Pipeline when Contract Signed = Yes

### Menu Functions
- Acquisition Tools > Refresh Dashboard
- Acquisition Tools > Update KPIs
- Acquisition Tools > Add New Lead (forms)
- Acquisition Tools > Setup > Initialize All Sheets
- Acquisition Tools > Setup > Import Buyers from Dispo Workspace
- Acquisition Tools > Setup > Start New Month
- Acquisition Tools > Setup > Import Zip Codes

### Time-Based (configure in Apps Script)
- Daily morning report (8 AM)
- Weekly summary (Monday 9 AM)
- Contract deadline checks (daily 7 AM)
- Monthly KPI archive (1st of month)

## Quick Start

1. Open your Google Sheet
2. Go to Extensions > Apps Script
3. Create files: Code.gs, Triggers.gs, Utilities.gs
4. Copy content from `google-apps-script/` folder
5. Save and refresh sheet
6. Click **Acquisition Tools > Setup > Initialize All Sheets**

## Monthly Goals (Customize in Code.gs)

```javascript
MONTHLY_GOALS: {
  LEADS_GENERATED: 100,
  OFFERS_MADE: 25,
  CONTRACTS_SIGNED: 5,
  DEALS_CLOSED: 3,
  REVENUE_TARGET: 50000
}
```

## Team Members (Customize in Code.gs)

```javascript
TEAM_MEMBERS: ['Tess Walter', 'Nick Barr']
```

## Custom Formulas

Use in any cell:
- `=MAO(ARV, Repairs, Fee)` - Maximum Allowable Offer
- `=PROFIT(Purchase, ARV, Repairs, Holding, Selling)` - Potential profit
- `=ASSIGNMENT_FEE(Contract, EndBuyer)` - Assignment fee
- `=DAYS_UNTIL(date)` - Days until deadline
- `=FORMAT_PHONE(phone)` - Format phone number
- `=MATCHING_BUYERS(Zip, Type, Price)` - Find matching buyers

## Importing Buyers

To import buyers from your Dispo Workspace 3.0:
1. Click **Acquisition Tools > Setup > Import Buyers from Dispo Workspace**
2. Enter the Spreadsheet ID from your Dispo Workspace URL
3. Buyers will be imported to the Buyers List sheet

## Adding Zip Codes

1. Click **Acquisition Tools > Setup > Import Zip Codes**
2. Enter zip codes (one per line or comma-separated)
3. Or paste directly from `templates/sample-zip-codes.csv`

## Version History

- **v4.0** - Simplified structure, removed Main Page, added team tracking, yearly KPIs
- **v3.1** - Full automation with Main Page consolidation
- **v3.0** - Added KPI tracking and Dashboard
