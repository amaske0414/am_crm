window.AM_CRM_SCHEMA_CATALOG = {
  "appName": "AM_CRM",
  "generatedAt": "2026-05-01",
  "sourceFiles": [
    "schema(1).xlsx"
  ],
  "tables": {
    "ref_ALG": {
      "datasetName": "ref_ALG",
      "fields": [
        {
          "datasetName": "ref_ALG",
          "fieldName": "ID_NAME_ALG",
          "fieldOrder": 1,
          "dataType": "string",
          "nullable": false,
          "primaryKey": true,
          "unique": true,
          "note": "ID_NAME_ALG aka client group ID"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "Slg Key Acct",
          "fieldOrder": 2,
          "dataType": "string",
          "nullable": false,
          "primaryKey": true,
          "unique": true,
          "note": "the internal unique ID associated with the client group ID/ID_NAME_ALG"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "url_ALG",
          "fieldOrder": 3,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": true,
          "note": "the url for viewing the client data in Morgan Stanley's internal custodian tool (3D)"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "NOTE_ALG",
          "fieldOrder": 4,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "status of the ALG data"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "url_MSO",
          "fieldOrder": 5,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": true,
          "note": "the url link for opening a view of the client's online access portal (Morgan Stanley Online aka MSO)"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "ID_PARTY_GPS",
          "fieldOrder": 6,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "unique ID associated with the client's financial plan"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "ID_NAME_GPS",
          "fieldOrder": 7,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "the name that shows on the client's financial plan"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "ID_3D GROUP_GPS",
          "fieldOrder": 8,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "unique ID associated with the client's financial plan"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "url_GPS",
          "fieldOrder": 9,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "the url link for opening a view of the client's financial plan (if available)"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "url_ALG_GPS",
          "fieldOrder": 10,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "the url for viewing the client data in Morgan Stanley's internal custodian tool including externally linked and manually added accounts (3D)"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "hasGPS",
          "fieldOrder": 11,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "indicates whether or not the client has a financial plan on file"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "NOTE_GPS",
          "fieldOrder": 12,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "ignore"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "STATUS",
          "fieldOrder": 13,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the client tier"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "NAME_FA",
          "fieldOrder": 14,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the name of the primary FA associated with the relationship"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "LAST NAME",
          "fieldOrder": 15,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the client's last name"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "SALUTATION",
          "fieldOrder": 16,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "this is the salutation that should be used at the beginning of any email correspondence with the client(s)"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "EMAIL",
          "fieldOrder": 17,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "the client(s) email address(es)"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "PHONE",
          "fieldOrder": 18,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "the primary client phone number"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "REVIEW LOCATION",
          "fieldOrder": 19,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "how/where the client(s) prefer to meet"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "GIC ALLOCATION",
          "fieldOrder": 20,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "the risk profile assigned to the client"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "CONTACT FREQUENCY",
          "fieldOrder": 21,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the target contact frequency for the client"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "REVIEW FREQUENCY",
          "fieldOrder": 22,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the target number of reviews conducted annually with the client"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "LAST REVIEW",
          "fieldOrder": 23,
          "dataType": "date",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "the date of the last review meeting with the client"
        },
        {
          "datasetName": "ref_ALG",
          "fieldName": "count_ID_NAME_ALG",
          "fieldOrder": 24,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "a check for duplicates (the field should always be = 1 for all rows)"
        }
      ],
      "primaryKeys": [
        "ID_NAME_ALG",
        "Slg Key Acct"
      ],
      "uniqueFields": [
        "ID_NAME_ALG",
        "Slg Key Acct",
        "url_ALG",
        "url_MSO"
      ]
    },
    "ref_ACCOUNT": {
      "datasetName": "ref_ACCOUNT",
      "fields": [
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Account Number",
          "fieldOrder": 1,
          "dataType": "string",
          "nullable": false,
          "primaryKey": true,
          "unique": true,
          "note": "The account number + 3 digit JPN number at the end of the string"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Contact",
          "fieldOrder": 2,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Contact display name"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Sort Name",
          "fieldOrder": 3,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "this is the ID_NAME_ALG aka client group ID"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Short Name",
          "fieldOrder": 4,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Short name / alias"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Status",
          "fieldOrder": 5,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "the client tier (should not be blank for any rows)"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Account Type",
          "fieldOrder": 6,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "e.g., IRA, AAA"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Account Category",
          "fieldOrder": 7,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "e.g., Investments, Trust"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Account Taxable / Non-Taxable",
          "fieldOrder": 8,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Taxability classification"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Opened Date",
          "fieldOrder": 9,
          "dataType": "date",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Account open date"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "SS / Tax ID",
          "fieldOrder": 10,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Masked in file (still treat as sensitive)"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "First Name",
          "fieldOrder": 11,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": null
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Middle Name",
          "fieldOrder": 12,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": null
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Last Name",
          "fieldOrder": 13,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": null
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Name Suffix",
          "fieldOrder": 14,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": null
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Date of Birth",
          "fieldOrder": 15,
          "dataType": "date",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "client date of birth"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Tax Bracket",
          "fieldOrder": 16,
          "dataType": "number",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "Percent-like values observed (e.g., 25)"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Total Net Worth",
          "fieldOrder": 17,
          "dataType": "number",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "Currency-like numeric"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Liquid Net Worth",
          "fieldOrder": 18,
          "dataType": "number",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "Currency-like numeric"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Total Annual Income",
          "fieldOrder": 19,
          "dataType": "number",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "Currency-like numeric"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Risk Tolerance",
          "fieldOrder": 20,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "e.g., Moderate"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Retirement Date",
          "fieldOrder": 21,
          "dataType": "date",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "Nullable"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Review Month",
          "fieldOrder": 22,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "Month label/value (nullable)"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Total Assets($)",
          "fieldOrder": 23,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Currency"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "SLG Total Assets($)",
          "fieldOrder": 24,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Currency"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "CG Fee Based Assets",
          "fieldOrder": 25,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Currency"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "SLG Prior Year Total Assets($)",
          "fieldOrder": 26,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Currency"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Total Liabilities($)",
          "fieldOrder": 27,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Currency"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Home Loans Balance($)",
          "fieldOrder": 28,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Currency"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Asset Allocation - Annuities & Insurance ($)",
          "fieldOrder": 29,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Currency"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Investment Assets from adopted Plan/Analysis",
          "fieldOrder": 30,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Currency"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "T12 Revenue($)",
          "fieldOrder": 31,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Currency"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "SLG T12 Total Revenue($)",
          "fieldOrder": 32,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Currency"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "SLG Prior Year Total Revenue($)",
          "fieldOrder": 33,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Currency"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "CG Fee-based T12 Revenue",
          "fieldOrder": 34,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Currency"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "RL Gain/Loss (Current Year)-($)",
          "fieldOrder": 35,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Currency (can be negative)"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Net Real G/L Prior Year($)",
          "fieldOrder": 36,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Currency"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Unrl Gain/Loss($)",
          "fieldOrder": 37,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Currency"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Tax Year",
          "fieldOrder": 38,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "e.g., 2026"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Checks",
          "fieldOrder": 39,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "Enrollment/eligibility status"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Debit Cards",
          "fieldOrder": 40,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "Enrollment/eligibility status"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "American Express Card",
          "fieldOrder": 41,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "Enrollment/eligibility status"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Bill Pay",
          "fieldOrder": 42,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "Enrollment/eligibility status"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Funds Transfer Service",
          "fieldOrder": 43,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "Enrollment/eligibility status"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Online Transfers",
          "fieldOrder": 44,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "Enrollment/eligibility status"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Account All Enrolled in eDelivery",
          "fieldOrder": 45,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "Yes/No"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "FPL Potential Annual Income($)",
          "fieldOrder": 46,
          "dataType": "number",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "Currency (nullable)"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Auto RMD",
          "fieldOrder": 47,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "Yes/No (nullable)"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "VDA on File",
          "fieldOrder": 48,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "Yes/No"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Last Traded",
          "fieldOrder": 49,
          "dataType": "date",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "Nullable"
        },
        {
          "datasetName": "ref_ACCOUNT",
          "fieldName": "Account Red Flagged",
          "fieldOrder": 50,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "Yes/No"
        }
      ],
      "primaryKeys": [
        "Account Number"
      ],
      "uniqueFields": [
        "Account Number"
      ]
    },
    "ref_SECURITY": {
      "datasetName": "ref_SECURITY",
      "fields": [
        {
          "datasetName": "ref_SECURITY",
          "fieldName": "ID_MSDW_SECURITY",
          "fieldOrder": 1,
          "dataType": "string",
          "nullable": false,
          "primaryKey": true,
          "unique": true,
          "note": "internal unique ID"
        },
        {
          "datasetName": "ref_SECURITY",
          "fieldName": "ID_CUSIP",
          "fieldOrder": 2,
          "dataType": "string",
          "nullable": false,
          "primaryKey": true,
          "unique": true,
          "note": "external/public unique ID"
        },
        {
          "datasetName": "ref_SECURITY",
          "fieldName": "Pref Iss Nme",
          "fieldOrder": 3,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "description of the security"
        },
        {
          "datasetName": "ref_SECURITY",
          "fieldName": "STRATEGY",
          "fieldOrder": 4,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "clean description of the strategy"
        },
        {
          "datasetName": "ref_SECURITY",
          "fieldName": "isALT",
          "fieldOrder": 5,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "indicates if the security is an alternative investment"
        },
        {
          "datasetName": "ref_SECURITY",
          "fieldName": "isESCROW",
          "fieldOrder": 6,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "indicates if the security is in escrow"
        },
        {
          "datasetName": "ref_SECURITY",
          "fieldName": "isCASH",
          "fieldOrder": 7,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "indicates if the security is classified as cash & equivalents"
        },
        {
          "datasetName": "ref_SECURITY",
          "fieldName": "isBDPS",
          "fieldOrder": 8,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "indicates if the security is pure cash"
        },
        {
          "datasetName": "ref_SECURITY",
          "fieldName": "isTAXFREE",
          "fieldOrder": 9,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "indicates if income from the security is subject to income taxes"
        },
        {
          "datasetName": "ref_SECURITY",
          "fieldName": "count_ID_CUSIP",
          "fieldOrder": 10,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "a check to ensure there are no duplicates (all rows should always = 1)"
        },
        {
          "datasetName": "ref_SECURITY",
          "fieldName": "count_ID_MSDW_SECURITY",
          "fieldOrder": 11,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "a check to ensure there are no duplicates (all rows should always = 1)"
        }
      ],
      "primaryKeys": [
        "ID_MSDW_SECURITY",
        "ID_CUSIP"
      ],
      "uniqueFields": [
        "ID_MSDW_SECURITY",
        "ID_CUSIP"
      ]
    },
    "ref_PRODUCT": {
      "datasetName": "ref_PRODUCT",
      "fields": [
        {
          "datasetName": "ref_PRODUCT",
          "fieldName": "Cde Prd",
          "fieldOrder": 1,
          "dataType": "string",
          "nullable": false,
          "primaryKey": true,
          "unique": true,
          "note": "the product code"
        },
        {
          "datasetName": "ref_PRODUCT",
          "fieldName": "PRODUCT_L1",
          "fieldOrder": 2,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Morgan Stanley grouping level 1"
        },
        {
          "datasetName": "ref_PRODUCT",
          "fieldName": "PRODUCT_L2",
          "fieldOrder": 3,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Morgan Stanley grouping level 2"
        },
        {
          "datasetName": "ref_PRODUCT",
          "fieldName": "PRODUCT_L3",
          "fieldOrder": 4,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Morgan Stanley grouping level 3"
        },
        {
          "datasetName": "ref_PRODUCT",
          "fieldName": "RECURRING",
          "fieldOrder": 5,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "indicates whether or not the revenue stream is recurring for the product"
        },
        {
          "datasetName": "ref_PRODUCT",
          "fieldName": "GROUP",
          "fieldOrder": 6,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "custom grouping"
        },
        {
          "datasetName": "ref_PRODUCT",
          "fieldName": "FREQUENCY",
          "fieldOrder": 7,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "indicates how often the revenue is received if recurring"
        },
        {
          "datasetName": "ref_PRODUCT",
          "fieldName": "isREVENUESHARE",
          "fieldOrder": 8,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "indicates whether or not the revenue from this product is included in revenue share agreements"
        }
      ],
      "primaryKeys": [
        "Cde Prd"
      ],
      "uniqueFields": [
        "Cde Prd"
      ]
    },
    "ref_WEALTHDESK": {
      "datasetName": "ref_WEALTHDESK",
      "fields": [
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "CRG Name",
          "fieldOrder": 1,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "ignore"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "Client Name",
          "fieldOrder": 2,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the client name in the system"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "Client type",
          "fieldOrder": 3,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "ignore"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "Account Number",
          "fieldOrder": 4,
          "dataType": "string",
          "nullable": false,
          "primaryKey": true,
          "unique": true,
          "note": "the account number + the 3 digit JPN number included at the end of the string"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "Account Type",
          "fieldOrder": 5,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "a long description of the account type"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "Account Restrictions",
          "fieldOrder": 6,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "ignore"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "Account Tax Status",
          "fieldOrder": 7,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "indicates whether or not the account is subject to taxes"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "Inception Date",
          "fieldOrder": 8,
          "dataType": "date",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the date the account was enrolled in the current strategy"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "Termination Date",
          "fieldOrder": 9,
          "dataType": "date",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "the date the account was terminated from the program"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "Inception Status",
          "fieldOrder": 10,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "ignore"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "Program",
          "fieldOrder": 11,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "The program the strategy was created in"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "Portfolio Name",
          "fieldOrder": 12,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "ignore"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "Investment Type",
          "fieldOrder": 13,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the name of the strategy the account is enrolled in"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "Available Cash",
          "fieldOrder": 14,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "ignore"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "Account Mkt Value",
          "fieldOrder": 15,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "ignore"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "Account Risk profile",
          "fieldOrder": 16,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "ignore"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "Account Volatility Score",
          "fieldOrder": 17,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the calculated risk associated with the strategy/strategies the account is invested in"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "Custodian",
          "fieldOrder": 18,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "ignore"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "UMA Program Discretion",
          "fieldOrder": 19,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "ignore"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "UMA Tax Managed Ind",
          "fieldOrder": 20,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "indicates whether or not the account is enrolled in UMA tax management"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "Client Rate",
          "fieldOrder": 21,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the total percent fee the client pays"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "Advisory Fee",
          "fieldOrder": 22,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the percent fee the advisor receives (subtract 7 bps if the Program = UMA)"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "Manager Fee",
          "fieldOrder": 23,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the percent fee the SMA manager receives (if applicable)"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "Other Fee",
          "fieldOrder": 24,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "ignore"
        },
        {
          "datasetName": "ref_WEALTHDESK",
          "fieldName": "YTD Fee",
          "fieldOrder": 25,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "ignore"
        }
      ],
      "primaryKeys": [
        "Account Number"
      ],
      "uniqueFields": [
        "Account Number"
      ]
    },
    "data_REVENUE": {
      "datasetName": "data_REVENUE",
      "fields": [
        {
          "datasetName": "data_REVENUE",
          "fieldName": "Num Ofc",
          "fieldOrder": 1,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the office code for the advisor associated with the transaction"
        },
        {
          "datasetName": "data_REVENUE",
          "fieldName": "Num Fa",
          "fieldOrder": 2,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the individual FA code for the advisor associated with the transaction"
        },
        {
          "datasetName": "data_REVENUE",
          "fieldName": "Num Acc",
          "fieldOrder": 3,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the account number"
        },
        {
          "datasetName": "data_REVENUE",
          "fieldName": "Key Acct",
          "fieldOrder": 4,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "a unique identifier for the account"
        },
        {
          "datasetName": "data_REVENUE",
          "fieldName": "Num Alt Ofc",
          "fieldOrder": 5,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the office code for the account"
        },
        {
          "datasetName": "data_REVENUE",
          "fieldName": "Num Alt Fa",
          "fieldOrder": 6,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the JPN number for the account"
        },
        {
          "datasetName": "data_REVENUE",
          "fieldName": "Cde Prd",
          "fieldOrder": 7,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the product code associated with the revenue transaction"
        },
        {
          "datasetName": "data_REVENUE",
          "fieldName": "Cde Tnd Cntl Bill Comb",
          "fieldOrder": 8,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "ignore"
        },
        {
          "datasetName": "data_REVENUE",
          "fieldName": "Cde Choice Acct",
          "fieldOrder": 9,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "ignore"
        },
        {
          "datasetName": "data_REVENUE",
          "fieldName": "Dt2 Trd Entr",
          "fieldOrder": 10,
          "dataType": "datetime",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the date of the transaction"
        },
        {
          "datasetName": "data_REVENUE",
          "fieldName": "Cde Pyout Basis",
          "fieldOrder": 11,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "ignore"
        },
        {
          "datasetName": "data_REVENUE",
          "fieldName": "Txt Pyout Basis Desc",
          "fieldOrder": 12,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "ignore"
        },
        {
          "datasetName": "data_REVENUE",
          "fieldName": "Amt Cmp Grs",
          "fieldOrder": 13,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Gross amount of revenue credited to the advisor for the transaction"
        },
        {
          "datasetName": "data_REVENUE",
          "fieldName": "Amt Gl Grs",
          "fieldOrder": 14,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Ignore"
        },
        {
          "datasetName": "data_REVENUE",
          "fieldName": "Amt Pyout",
          "fieldOrder": 15,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Gross amount paid to the advisor for the revenue transaction"
        },
        {
          "datasetName": "data_REVENUE",
          "fieldName": "Amt Agg Asst Cmsn Alloc",
          "fieldOrder": 16,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Ignore"
        },
        {
          "datasetName": "data_REVENUE",
          "fieldName": "Amt Cr Rvnu",
          "fieldOrder": 17,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Ignore"
        },
        {
          "datasetName": "data_REVENUE",
          "fieldName": "Amt Disc",
          "fieldOrder": 18,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Ignore"
        },
        {
          "datasetName": "data_REVENUE",
          "fieldName": "Amt Net",
          "fieldOrder": 19,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Ignore"
        },
        {
          "datasetName": "data_REVENUE",
          "fieldName": "Num Emp Fa",
          "fieldOrder": 20,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Unique ID indicating the advisor receiving credit for the transaction."
        }
      ],
      "primaryKeys": [],
      "uniqueFields": []
    },
    "data_POSITIONS": {
      "datasetName": "data_POSITIONS",
      "fields": [
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Key Acct",
          "fieldOrder": 1,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Unique ID that can be used to link account reference files"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Branch Number",
          "fieldOrder": 2,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the 3 digit office code"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Account Number",
          "fieldOrder": 3,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the 6 digit account number"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "FA Number",
          "fieldOrder": 4,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the 3 digit JPN number"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Name Account",
          "fieldOrder": 5,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the name listed on the account"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Managed",
          "fieldOrder": 6,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "\"Y\" indicates this account is enrolled in an advisory program.  \"N\" indicates that the account is unmanaged."
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Dt2 End",
          "fieldOrder": 7,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "this is the as of date for the market values.  If the data is for the current month, this field will always show the last business day of the month."
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Nme Sym",
          "fieldOrder": 8,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "the security symbol"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Num Cusip",
          "fieldOrder": 9,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "the security CUSIP"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Cde Msdw Sec",
          "fieldOrder": 10,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the 9-character unique ID for the security.  This is an internal identfier."
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Pref Iss Nme",
          "fieldOrder": 11,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "the long description for the security"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Txt Lvl1 Ast Desc",
          "fieldOrder": 12,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the most granular asset class/category for the security"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Txt Lvl2 Ast Desc",
          "fieldOrder": 13,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the second most granular asset class/category for the security"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Txt Lvl3 Ast Desc",
          "fieldOrder": 14,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the highest level asset class/category grouping for the security"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Txt Lvl4 Ast Desc",
          "fieldOrder": 15,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Ignore"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Cde Ms Rtg",
          "fieldOrder": 16,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "Morgan Stanley Research Analyst rating for the security (if applicable)"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Nme Iss Typ",
          "fieldOrder": 17,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "product groupings"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Nme Iss Sub Typ",
          "fieldOrder": 18,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "product groupings"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Nme Asset Typ",
          "fieldOrder": 19,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "product groupings"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Num Face Val",
          "fieldOrder": 20,
          "dataType": "number",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "Ignore"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Dt2 Call",
          "fieldOrder": 21,
          "dataType": "date",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "the next date the security can by called by the issuer (if applicable)"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Mat Exp Date",
          "fieldOrder": 22,
          "dataType": "date",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "the date the security is scheduled to mature (if applicable)"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Pct Coupon",
          "fieldOrder": 23,
          "dataType": "number",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "the annual coupon percent for the security (if applicable)"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Amt Avg Dly Price",
          "fieldOrder": 24,
          "dataType": "number",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "the average daily price for the security"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Num Qty",
          "fieldOrder": 25,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the number of shares/units owned"
        },
        {
          "datasetName": "data_POSITIONS",
          "fieldName": "Amt Mkt Val",
          "fieldOrder": 26,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "the market value of the security held in the account"
        }
      ],
      "primaryKeys": [],
      "uniqueFields": []
    },
    "data_ACTIVITY": {
      "datasetName": "data_ACTIVITY",
      "fields": [
        {
          "datasetName": "data_ACTIVITY",
          "fieldName": "Client",
          "fieldOrder": 1,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "This is the system/database name for the client that owns the account"
        },
        {
          "datasetName": "data_ACTIVITY",
          "fieldName": "Account Number",
          "fieldOrder": 2,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "This is the account number"
        },
        {
          "datasetName": "data_ACTIVITY",
          "fieldName": "Sort Name",
          "fieldOrder": 3,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "This is the ID_NAME_ALG"
        },
        {
          "datasetName": "data_ACTIVITY",
          "fieldName": "Activity Type",
          "fieldOrder": 4,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "This the system defined category for the activity"
        },
        {
          "datasetName": "data_ACTIVITY",
          "fieldName": "Symbol",
          "fieldOrder": 5,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "This is the security symbol (if applicable)"
        },
        {
          "datasetName": "data_ACTIVITY",
          "fieldName": "CUSIP",
          "fieldOrder": 6,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "This is the security CUSIP (if applicable)"
        },
        {
          "datasetName": "data_ACTIVITY",
          "fieldName": "Activity Description",
          "fieldOrder": 7,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "This is the long description for the activity"
        },
        {
          "datasetName": "data_ACTIVITY",
          "fieldName": "Quantity",
          "fieldOrder": 8,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "This is the number of shares/units included in the transaction (if applicable)"
        },
        {
          "datasetName": "data_ACTIVITY",
          "fieldName": "Price($)",
          "fieldOrder": 9,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "This is the price of the security at the time of the activity (if applicable)"
        },
        {
          "datasetName": "data_ACTIVITY",
          "fieldName": "Amount($)",
          "fieldOrder": 10,
          "dataType": "number",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "This is the dollar amount associated with the activity"
        },
        {
          "datasetName": "data_ACTIVITY",
          "fieldName": "Activity Date",
          "fieldOrder": 11,
          "dataType": "date",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "This is the date the activity occurred"
        },
        {
          "datasetName": "data_ACTIVITY",
          "fieldName": "Settlement Date",
          "fieldOrder": 12,
          "dataType": "date",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "This is the date the activity settled"
        },
        {
          "datasetName": "data_ACTIVITY",
          "fieldName": "Account Type",
          "fieldOrder": 13,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "This is the type of account the activity took place in"
        },
        {
          "datasetName": "data_ACTIVITY",
          "fieldName": "Entity Type",
          "fieldOrder": 14,
          "dataType": "string",
          "nullable": false,
          "primaryKey": false,
          "unique": false,
          "note": "Ignore"
        },
        {
          "datasetName": "data_ACTIVITY",
          "fieldName": "MSID",
          "fieldOrder": 15,
          "dataType": "string",
          "nullable": true,
          "primaryKey": false,
          "unique": false,
          "note": "This is the unique ID of the employee who posted the activity"
        }
      ],
      "primaryKeys": [],
      "uniqueFields": []
    }
  },
  "metadata": {
    "ref_ALG": {
      "grain": "one row per client group ID (aka ALG)",
      "refreshCadence": "ad-hoc",
      "recordCount": "80 - 120",
      "filtersExclusions": null
    },
    "data_POSITIONS": {
      "grain": "one row per position per account",
      "refreshCadence": "daily",
      "recordCount": "10,000 - 20,000",
      "filtersExclusions": null
    },
    "data_ACTIVITY": {
      "grain": "one row per transaction",
      "refreshCadence": "daily",
      "recordCount": "variable (1,000 - 100,000)",
      "filtersExclusions": "excludes trade activity and dividend activity"
    },
    "data_REVENUE": {
      "grain": "one row per account code per product code per date per advisor (Num Emp Fa)",
      "refreshCadence": "weekly",
      "recordCount": "variable (80,000 - 120,000)",
      "filtersExclusions": null
    },
    "ref_PRODUCT": {
      "grain": "one row per product code",
      "refreshCadence": "ad-hoc",
      "recordCount": "97 + growth",
      "filtersExclusions": "missing certain product codes that have never been observed in data_REVENUE"
    },
    "ref_SECURITY": {
      "grain": "one row per security",
      "refreshCadence": "ad-hoc",
      "recordCount": "140 + growth",
      "filtersExclusions": "Only includes securities that data_POSITIONS is unable to classify"
    },
    "ref_WEALTHDESK": {
      "grain": "one row per account",
      "refreshCadence": "ad-hoc",
      "recordCount": "500 - 800",
      "filtersExclusions": "only includes accounts currently enrolled in an advisory program"
    }
  },
  "relationshipMapping": {
    "title": "Relationship Mapping Instructions",
    "purpose": "Defines the approved join logic for linking the workbook's reference and data tables. Use these rows as the technical relationship map when building queries, Power Query steps, BI models, or validation rules.",
    "accountKeyStandard": "Canonical account key format is ###-######. For account-number fields that may include suffixes or extra characters, compare LEFT([Account Number], 10). For branch/office + account fields, build the key by zero-padding branch/office to 3 digits with a trailing dash and zero-padding account number to 6 digits.",
    "matchHandling": "Treat join keys as text, preserve leading zeroes, and trim imported flat-file values before comparison. Default to LEFT JOIN from the fact/detail table to the lookup/reference table when enriching data; use INNER JOIN only when unmatched records should be excluded.",
    "relationships": [
      {
        "relationshipId": "R001",
        "lookupParentTable": "ref_PRODUCT",
        "lookupKeyColumnOrExpression": "ref_PRODUCT.[Cde Prd]",
        "detailRelatedTable": "data_REVENUE",
        "detailKeyColumnOrExpression": "data_REVENUE.[Cde Prd]",
        "canonicalJoinPredicate": "ref_PRODUCT.[Cde Prd] = data_REVENUE.[Cde Prd]",
        "recommendedJoinUsage": "Use data_REVENUE LEFT JOIN ref_PRODUCT when enriching revenue records with product attributes.",
        "expectedCardinality": "ref_PRODUCT 1 : many data_REVENUE",
        "keyNormalizationRule": "Compare product code as text. Preserve the exact code value; trim whitespace if loaded from flat files.",
        "validationChecks": "Confirm ref_PRODUCT.[Cde Prd] is unique. Check for non-null data_REVENUE.[Cde Prd] values with no matching ref_PRODUCT row.",
        "implementationNotes": "Product reference supplies product hierarchy/classification for each revenue record."
      },
      {
        "relationshipId": "R002",
        "lookupParentTable": "ref_ACCOUNT",
        "lookupKeyColumnOrExpression": "LEFT(ref_ACCOUNT.[Account Number], 10)",
        "detailRelatedTable": "data_REVENUE",
        "detailKeyColumnOrExpression": "TEXT(data_REVENUE.[Num Alt Ofc], \"000-\") & TEXT(data_REVENUE.[Num Acc], \"000000\")",
        "canonicalJoinPredicate": "LEFT(ref_ACCOUNT.[Account Number], 10) = TEXT(data_REVENUE.[Num Alt Ofc], \"000-\") & TEXT(data_REVENUE.[Num Acc], \"000000\")",
        "recommendedJoinUsage": "Use data_REVENUE LEFT JOIN ref_ACCOUNT when enriching revenue records with account attributes.",
        "expectedCardinality": "ref_ACCOUNT 1 : many data_REVENUE",
        "keyNormalizationRule": "Build the revenue-side account key as ###-######. Zero-pad Num Alt Ofc to 3 digits plus dash and Num Acc to 6 digits. Compare to the first 10 characters of ref_ACCOUNT.[Account Number].",
        "validationChecks": "Check for blank Num Alt Ofc or Num Acc values, duplicate LEFT(ref_ACCOUNT.[Account Number], 10) keys, and revenue records with no account match.",
        "implementationNotes": "Do not join to the full ref_ACCOUNT.[Account Number] if it contains suffixes or characters beyond the first 10-character account key."
      },
      {
        "relationshipId": "R003",
        "lookupParentTable": "ref_ACCOUNT",
        "lookupKeyColumnOrExpression": "LEFT(ref_ACCOUNT.[Account Number], 10)",
        "detailRelatedTable": "data_POSITIONS",
        "detailKeyColumnOrExpression": "TEXT(data_POSITIONS.[Branch Number], \"000-\") & TEXT(data_POSITIONS.[Account Number], \"000000\")",
        "canonicalJoinPredicate": "LEFT(ref_ACCOUNT.[Account Number], 10) = TEXT(data_POSITIONS.[Branch Number], \"000-\") & TEXT(data_POSITIONS.[Account Number], \"000000\")",
        "recommendedJoinUsage": "Use data_POSITIONS LEFT JOIN ref_ACCOUNT when enriching position records with account attributes.",
        "expectedCardinality": "ref_ACCOUNT 1 : many data_POSITIONS",
        "keyNormalizationRule": "Build the positions-side account key as ###-######. Zero-pad Branch Number to 3 digits plus dash and Account Number to 6 digits. Compare to the first 10 characters of ref_ACCOUNT.[Account Number].",
        "validationChecks": "Check for blank Branch Number or Account Number values, duplicate LEFT(ref_ACCOUNT.[Account Number], 10) keys, and positions with no account match.",
        "implementationNotes": "Positions can contain multiple rows per account because holdings may vary by security and reporting date."
      },
      {
        "relationshipId": "R004",
        "lookupParentTable": "ref_WEALTHDESK",
        "lookupKeyColumnOrExpression": "LEFT(ref_WEALTHDESK.[Account Number], 10)",
        "detailRelatedTable": "data_POSITIONS",
        "detailKeyColumnOrExpression": "TEXT(data_POSITIONS.[Branch Number], \"000-\") & TEXT(data_POSITIONS.[Account Number], \"000000\")",
        "canonicalJoinPredicate": "LEFT(ref_WEALTHDESK.[Account Number], 10) = TEXT(data_POSITIONS.[Branch Number], \"000-\") & TEXT(data_POSITIONS.[Account Number], \"000000\")",
        "recommendedJoinUsage": "Use data_POSITIONS LEFT JOIN ref_WEALTHDESK when enriching positions with WealthDesk account/program attributes. Aggregate positions first if the output must stay one row per WealthDesk account.",
        "expectedCardinality": "ref_WEALTHDESK 1 : many data_POSITIONS",
        "keyNormalizationRule": "Use the first 10 characters of ref_WEALTHDESK.[Account Number] as the WealthDesk account key. Build the positions-side key as ###-###### from Branch Number and Account Number.",
        "validationChecks": "Check that LEFT(ref_WEALTHDESK.[Account Number], 10) is unique. Check for position accounts that do not match a WealthDesk account.",
        "implementationNotes": "Use ref_WEALTHDESK as the workbook table name for this relationship."
      },
      {
        "relationshipId": "R005",
        "lookupParentTable": "ref_ACCOUNT",
        "lookupKeyColumnOrExpression": "LEFT(ref_ACCOUNT.[Account Number], 10)",
        "detailRelatedTable": "data_ACTIVITY",
        "detailKeyColumnOrExpression": "LEFT(data_ACTIVITY.[Account Number], 10)",
        "canonicalJoinPredicate": "LEFT(ref_ACCOUNT.[Account Number], 10) = LEFT(data_ACTIVITY.[Account Number], 10)",
        "recommendedJoinUsage": "Use data_ACTIVITY LEFT JOIN ref_ACCOUNT when enriching activity records with account attributes.",
        "expectedCardinality": "ref_ACCOUNT 1 : many data_ACTIVITY",
        "keyNormalizationRule": "Compare the first 10 characters of both account-number fields as the canonical account key. Treat both values as text and trim before applying LEFT.",
        "validationChecks": "Check duplicate LEFT(ref_ACCOUNT.[Account Number], 10) keys and activity rows with blank or unmatched account keys.",
        "implementationNotes": "This relationship assumes the activity account number begins with the same 10-character ###-###### account key."
      },
      {
        "relationshipId": "R006",
        "lookupParentTable": "ref_SECURITY",
        "lookupKeyColumnOrExpression": "ref_SECURITY.[ID_MSDW_SECURITY]",
        "detailRelatedTable": "data_POSITIONS",
        "detailKeyColumnOrExpression": "data_POSITIONS.[Cde Msdw Sec]",
        "canonicalJoinPredicate": "ref_SECURITY.[ID_MSDW_SECURITY] = data_POSITIONS.[Cde Msdw Sec]",
        "recommendedJoinUsage": "Use data_POSITIONS LEFT JOIN ref_SECURITY when enriching positions with security attributes.",
        "expectedCardinality": "ref_SECURITY 1 : many data_POSITIONS",
        "keyNormalizationRule": "Compare security identifiers as text. Preserve leading zeroes or special characters if present.",
        "validationChecks": "Confirm ref_SECURITY.[ID_MSDW_SECURITY] is unique. Check for non-null data_POSITIONS.[Cde Msdw Sec] values with no matching security row.",
        "implementationNotes": "Security reference supplies CUSIP, preferred issuer name, strategy, and security classification flags."
      },
      {
        "relationshipId": "R007",
        "lookupParentTable": "ref_ALG",
        "lookupKeyColumnOrExpression": "UPPER(ref_ALG.[ID_NAME_ALG])",
        "detailRelatedTable": "ref_ACCOUNT",
        "detailKeyColumnOrExpression": "UPPER(ref_ACCOUNT.[Sort Name])",
        "canonicalJoinPredicate": "UPPER(ref_ACCOUNT.[Sort Name]) = UPPER(ref_ALG.[ID_NAME_ALG])",
        "recommendedJoinUsage": "Use ref_ACCOUNT LEFT JOIN ref_ALG when enriching account records with ALG links/status. Match names case-insensitively.",
        "expectedCardinality": "ref_ALG 1 : many ref_ACCOUNT, based on unique ref_ALG.[ID_NAME_ALG]",
        "keyNormalizationRule": "Normalize both name fields to UPPER(TRIM(value)) for implementation. Use exact text comparison after normalization.",
        "validationChecks": "Check duplicate normalized ref_ALG.[ID_NAME_ALG] values, duplicate or ambiguous normalized ref_ACCOUNT.[Sort Name] values, and unmatched account names.",
        "implementationNotes": "Name-based joins are less stable than ID/account-number joins; review ambiguous or missing matches manually."
      }
    ]
  }
};
