# AM_CRM Data Schema Summary

Generated from `source_workbooks/schema.xlsx`.

## Table Metadata
| Dataset | Grain | Refresh Cadence | Record Count | Filters / Exclusions |
| --- | --- | --- | --- | --- |
| ref_ALG | one row per client group ID (aka ALG) | ad-hoc | 80 - 120 |  |
| data_POSITIONS | one row per position per account | daily | 10,000 - 20,000 |  |
| data_ACTIVITY | one row per transaction | daily | variable (1,000 - 100,000) | excludes trade activity and dividend activity |
| data_REVENUE | one row per account code per product code per date per advisor (Num Emp Fa) | weekly | variable (80,000 - 120,000) |  |
| ref_PRODUCT | one row per product code | ad-hoc | 97 + growth | missing certain product codes that have never been observed in data_REVENUE |
| ref_SECURITY | one row per security | ad-hoc | 140 + growth | Only includes securities that data_POSITIONS is unable to classify |
| ref_WEALTHDESK | one row per account | ad-hoc | 500 - 800 | only includes accounts currently enrolled in an advisory program |

## Relationship Mapping
Canonical account key format is ###-######. For account-number fields that may include suffixes or extra characters, compare LEFT([Account Number], 10). For branch/office + account fields, build the key by zero-padding branch/office to 3 digits with a trailing dash and zero-padding account number to 6 digits.

| ID | Lookup / Parent | Lookup Key | Detail / Related | Detail Key | Canonical Predicate |
| --- | --- | --- | --- | --- | --- |
| R001 | ref_PRODUCT | ref_PRODUCT.[Cde Prd] | data_REVENUE | data_REVENUE.[Cde Prd] | ref_PRODUCT.[Cde Prd] = data_REVENUE.[Cde Prd] |
| R002 | ref_ACCOUNT | LEFT(ref_ACCOUNT.[Account Number], 10) | data_REVENUE | TEXT(data_REVENUE.[Num Alt Ofc], "000-") & TEXT(data_REVENUE.[Num Acc], "000000") | LEFT(ref_ACCOUNT.[Account Number], 10) = TEXT(data_REVENUE.[Num Alt Ofc], "000-") & TEXT(data_REVENUE.[Num Acc], "000000") |
| R003 | ref_ACCOUNT | LEFT(ref_ACCOUNT.[Account Number], 10) | data_POSITIONS | TEXT(data_POSITIONS.[Branch Number], "000-") & TEXT(data_POSITIONS.[Account Number], "000000") | LEFT(ref_ACCOUNT.[Account Number], 10) = TEXT(data_POSITIONS.[Branch Number], "000-") & TEXT(data_POSITIONS.[Account Number], "000000") |
| R004 | ref_WEALTHDESK | LEFT(ref_WEALTHDESK.[Account Number], 10) | data_POSITIONS | TEXT(data_POSITIONS.[Branch Number], "000-") & TEXT(data_POSITIONS.[Account Number], "000000") | LEFT(ref_WEALTHDESK.[Account Number], 10) = TEXT(data_POSITIONS.[Branch Number], "000-") & TEXT(data_POSITIONS.[Account Number], "000000") |
| R005 | ref_ACCOUNT | LEFT(ref_ACCOUNT.[Account Number], 10) | data_ACTIVITY | LEFT(data_ACTIVITY.[Account Number], 10) | LEFT(ref_ACCOUNT.[Account Number], 10) = LEFT(data_ACTIVITY.[Account Number], 10) |
| R006 | ref_SECURITY | ref_SECURITY.[ID_MSDW_SECURITY] | data_POSITIONS | data_POSITIONS.[Cde Msdw Sec] | ref_SECURITY.[ID_MSDW_SECURITY] = data_POSITIONS.[Cde Msdw Sec] |
| R007 | ref_ALG | UPPER(ref_ALG.[ID_NAME_ALG]) | ref_ACCOUNT | UPPER(ref_ACCOUNT.[Sort Name]) | UPPER(ref_ACCOUNT.[Sort Name]) = UPPER(ref_ALG.[ID_NAME_ALG]) |

## ref_ALG
| Order | Field | Type | Nullable | Primary Key | Unique | Note |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | ID_NAME_ALG | string | N | Y | Y | ID_NAME_ALG aka client group ID |
| 2 | Slg Key Acct | string | N | Y | Y | the internal unique ID associated with the client group ID/ID_NAME_ALG |
| 3 | url_ALG | string | N | N | Y | the url for viewing the client data in Morgan Stanley's internal custodian tool (3D) |
| 4 | NOTE_ALG | string | N | N | N | status of the ALG data |
| 5 | url_MSO | string | N | N | Y | the url link for opening a view of the client's online access portal (Morgan Stanley Online aka MSO) |
| 6 | ID_PARTY_GPS | string | Y | N | N | unique ID associated with the client's financial plan |
| 7 | ID_NAME_GPS | string | Y | N | N | the name that shows on the client's financial plan |
| 8 | ID_3D GROUP_GPS | string | Y | N | N | unique ID associated with the client's financial plan |
| 9 | url_GPS | string | Y | N | N | the url link for opening a view of the client's financial plan (if available) |
| 10 | url_ALG_GPS | string | Y | N | N | the url for viewing the client data in Morgan Stanley's internal custodian tool including externally linked and manually added accounts (3D) |
| 11 | hasGPS | string | Y | N | N | indicates whether or not the client has a financial plan on file |
| 12 | NOTE_GPS | string | Y | N | N | ignore |
| 13 | STATUS | string | N | N | N | the client tier |
| 14 | NAME_FA | string | N | N | N | the name of the primary FA associated with the relationship |
| 15 | LAST NAME | string | N | N | N | the client's last name |
| 16 | SALUTATION | string | N | N | N | this is the salutation that should be used at the beginning of any email correspondence with the client(s) |
| 17 | EMAIL | string | Y | N | N | the client(s) email address(es) |
| 18 | PHONE | string | Y | N | N | the primary client phone number |
| 19 | REVIEW LOCATION | string | Y | N | N | how/where the client(s) prefer to meet |
| 20 | GIC ALLOCATION | string | Y | N | N | the risk profile assigned to the client |
| 21 | CONTACT FREQUENCY | string | N | N | N | the target contact frequency for the client |
| 22 | REVIEW FREQUENCY | string | N | N | N | the target number of reviews conducted annually with the client |
| 23 | LAST REVIEW | date | Y | N | N | the date of the last review meeting with the client |
| 24 | count_ID_NAME_ALG | string | N | N | N | a check for duplicates (the field should always be = 1 for all rows) |

## ref_ACCOUNT
| Order | Field | Type | Nullable | Primary Key | Unique | Note |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | Account Number | string | N | Y | Y | The account number + 3 digit JPN number at the end of the string |
| 2 | Contact | string | N | N | N | Contact display name |
| 3 | Sort Name | string | N | N | N | this is the ID_NAME_ALG aka client group ID |
| 4 | Short Name | string | N | N | N | Short name / alias |
| 5 | Status | string | Y | N | N | the client tier (should not be blank for any rows) |
| 6 | Account Type | string | N | N | N | e.g., IRA, AAA |
| 7 | Account Category | string | N | N | N | e.g., Investments, Trust |
| 8 | Account Taxable / Non-Taxable | string | N | N | N | Taxability classification |
| 9 | Opened Date | date | N | N | N | Account open date |
| 10 | SS / Tax ID | string | N | N | N | Masked in file (still treat as sensitive) |
| 11 | First Name | string | Y | N | N |  |
| 12 | Middle Name | string | Y | N | N |  |
| 13 | Last Name | string | Y | N | N |  |
| 14 | Name Suffix | string | Y | N | N |  |
| 15 | Date of Birth | date | Y | N | N | client date of birth |
| 16 | Tax Bracket | number | Y | N | N | Percent-like values observed (e.g., 25) |
| 17 | Total Net Worth | number | Y | N | N | Currency-like numeric |
| 18 | Liquid Net Worth | number | Y | N | N | Currency-like numeric |
| 19 | Total Annual Income | number | Y | N | N | Currency-like numeric |
| 20 | Risk Tolerance | string | Y | N | N | e.g., Moderate |
| 21 | Retirement Date | date | Y | N | N | Nullable |
| 22 | Review Month | string | Y | N | N | Month label/value (nullable) |
| 23 | Total Assets($) | number | N | N | N | Currency |
| 24 | SLG Total Assets($) | number | N | N | N | Currency |
| 25 | CG Fee Based Assets | number | N | N | N | Currency |
| 26 | SLG Prior Year Total Assets($) | number | N | N | N | Currency |
| 27 | Total Liabilities($) | number | N | N | N | Currency |
| 28 | Home Loans Balance($) | number | N | N | N | Currency |
| 29 | Asset Allocation - Annuities & Insurance ($) | number | N | N | N | Currency |
| 30 | Investment Assets from adopted Plan/Analysis | number | N | N | N | Currency |
| 31 | T12 Revenue($) | number | N | N | N | Currency |
| 32 | SLG T12 Total Revenue($) | number | N | N | N | Currency |
| 33 | SLG Prior Year Total Revenue($) | number | N | N | N | Currency |
| 34 | CG Fee-based T12 Revenue | number | N | N | N | Currency |
| 35 | RL Gain/Loss (Current Year)-($) | number | N | N | N | Currency (can be negative) |
| 36 | Net Real G/L Prior Year($) | number | N | N | N | Currency |
| 37 | Unrl Gain/Loss($) | number | N | N | N | Currency |
| 38 | Tax Year | number | N | N | N | e.g., 2026 |
| 39 | Checks | string | Y | N | N | Enrollment/eligibility status |
| 40 | Debit Cards | string | Y | N | N | Enrollment/eligibility status |
| 41 | American Express Card | string | Y | N | N | Enrollment/eligibility status |
| 42 | Bill Pay | string | Y | N | N | Enrollment/eligibility status |
| 43 | Funds Transfer Service | string | Y | N | N | Enrollment/eligibility status |
| 44 | Online Transfers | string | Y | N | N | Enrollment/eligibility status |
| 45 | Account All Enrolled in eDelivery | string | Y | N | N | Yes/No |
| 46 | FPL Potential Annual Income($) | number | Y | N | N | Currency (nullable) |
| 47 | Auto RMD | string | Y | N | N | Yes/No (nullable) |
| 48 | VDA on File | string | Y | N | N | Yes/No |
| 49 | Last Traded | date | Y | N | N | Nullable |
| 50 | Account Red Flagged | string | Y | N | N | Yes/No |

## ref_SECURITY
| Order | Field | Type | Nullable | Primary Key | Unique | Note |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | ID_MSDW_SECURITY | string | N | Y | Y | internal unique ID |
| 2 | ID_CUSIP | string | N | Y | Y | external/public unique ID |
| 3 | Pref Iss Nme | string | N | N | N | description of the security |
| 4 | STRATEGY | string | N | N | N | clean description of the strategy |
| 5 | isALT | string | Y | N | N | indicates if the security is an alternative investment |
| 6 | isESCROW | string | Y | N | N | indicates if the security is in escrow |
| 7 | isCASH | string | Y | N | N | indicates if the security is classified as cash & equivalents |
| 8 | isBDPS | string | Y | N | N | indicates if the security is pure cash |
| 9 | isTAXFREE | string | Y | N | N | indicates if income from the security is subject to income taxes |
| 10 | count_ID_CUSIP | number | N | N | N | a check to ensure there are no duplicates (all rows should always = 1) |
| 11 | count_ID_MSDW_SECURITY | number | N | N | N | a check to ensure there are no duplicates (all rows should always = 1) |

## ref_PRODUCT
| Order | Field | Type | Nullable | Primary Key | Unique | Note |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | Cde Prd | string | N | Y | Y | the product code |
| 2 | PRODUCT_L1 | string | N | N | N | Morgan Stanley grouping level 1 |
| 3 | PRODUCT_L2 | string | N | N | N | Morgan Stanley grouping level 2 |
| 4 | PRODUCT_L3 | string | N | N | N | Morgan Stanley grouping level 3 |
| 5 | RECURRING | string | N | N | N | indicates whether or not the revenue stream is recurring for the product |
| 6 | GROUP | string | N | N | N | custom grouping |
| 7 | FREQUENCY | string | N | N | N | indicates how often the revenue is received if recurring |
| 8 | isREVENUESHARE | string | Y | N | N | indicates whether or not the revenue from this product is included in revenue share agreements |

## ref_WEALTHDESK
| Order | Field | Type | Nullable | Primary Key | Unique | Note |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | CRG Name | string | Y | N | N | ignore |
| 2 | Client Name | string | N | N | N | the client name in the system |
| 3 | Client type | string | N | N | N | ignore |
| 4 | Account Number | string | N | Y | Y | the account number + the 3 digit JPN number included at the end of the string |
| 5 | Account Type | string | N | N | N | a long description of the account type |
| 6 | Account Restrictions | string | N | N | N | ignore |
| 7 | Account Tax Status | string | N | N | N | indicates whether or not the account is subject to taxes |
| 8 | Inception Date | date | N | N | N | the date the account was enrolled in the current strategy |
| 9 | Termination Date | date | Y | N | N | the date the account was terminated from the program |
| 10 | Inception Status | string | N | N | N | ignore |
| 11 | Program | string | N | N | N | The program the strategy was created in |
| 12 | Portfolio Name | string | Y | N | N | ignore |
| 13 | Investment Type | string | N | N | N | the name of the strategy the account is enrolled in |
| 14 | Available Cash | number | N | N | N | ignore |
| 15 | Account Mkt Value | number | N | N | N | ignore |
| 16 | Account Risk profile | string | N | N | N | ignore |
| 17 | Account Volatility Score | number | N | N | N | the calculated risk associated with the strategy/strategies the account is invested in |
| 18 | Custodian | string | N | N | N | ignore |
| 19 | UMA Program Discretion | string | Y | N | N | ignore |
| 20 | UMA Tax Managed Ind | string | Y | N | N | indicates whether or not the account is enrolled in UMA tax management |
| 21 | Client Rate | number | N | N | N | the total percent fee the client pays |
| 22 | Advisory Fee | number | N | N | N | the percent fee the advisor receives (subtract 7 bps if the Program = UMA) |
| 23 | Manager Fee | number | N | N | N | the percent fee the SMA manager receives (if applicable) |
| 24 | Other Fee | number | N | N | N | ignore |
| 25 | YTD Fee | number | N | N | N | ignore |

## data_REVENUE
| Order | Field | Type | Nullable | Primary Key | Unique | Note |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | Num Ofc | string | N | N | N | the office code for the advisor associated with the transaction |
| 2 | Num Fa | string | N | N | N | the individual FA code for the advisor associated with the transaction |
| 3 | Num Acc | string | N | N | N | the account number |
| 4 | Key Acct | string | N | N | N | a unique identifier for the account |
| 5 | Num Alt Ofc | string | N | N | N | the office code for the account |
| 6 | Num Alt Fa | string | N | N | N | the JPN number for the account |
| 7 | Cde Prd | string | N | N | N | the product code associated with the revenue transaction |
| 8 | Cde Tnd Cntl Bill Comb | string | N | N | N | ignore |
| 9 | Cde Choice Acct | string | N | N | N | ignore |
| 10 | Dt2 Trd Entr | datetime | N | N | N | the date of the transaction |
| 11 | Cde Pyout Basis | string | N | N | N | ignore |
| 12 | Txt Pyout Basis Desc | string | N | N | N | ignore |
| 13 | Amt Cmp Grs | number | N | N | N | Gross amount of revenue credited to the advisor for the transaction |
| 14 | Amt Gl Grs | number | N | N | N | Ignore |
| 15 | Amt Pyout | number | N | N | N | Gross amount paid to the advisor for the revenue transaction |
| 16 | Amt Agg Asst Cmsn Alloc | number | N | N | N | Ignore |
| 17 | Amt Cr Rvnu | number | N | N | N | Ignore |
| 18 | Amt Disc | number | N | N | N | Ignore |
| 19 | Amt Net | number | N | N | N | Ignore |
| 20 | Num Emp Fa | string | N | N | N | Unique ID indicating the advisor receiving credit for the transaction. |

## data_POSITIONS
| Order | Field | Type | Nullable | Primary Key | Unique | Note |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | Key Acct | string | N | N | N | Unique ID that can be used to link account reference files |
| 2 | Branch Number | string | N | N | N | the 3 digit office code |
| 3 | Account Number | string | N | N | N | the 6 digit account number |
| 4 | FA Number | string | N | N | N | the 3 digit JPN number |
| 5 | Name Account | string | N | N | N | the name listed on the account |
| 6 | Managed | string | N | N | N | "Y" indicates this account is enrolled in an advisory program.  "N" indicates that the account is unmanaged. |
| 7 | Dt2 End | string | N | N | N | this is the as of date for the market values.  If the data is for the current month, this field will always show the last business day of the month. |
| 8 | Nme Sym | string | Y | N | N | the security symbol |
| 9 | Num Cusip | string | Y | N | N | the security CUSIP |
| 10 | Cde Msdw Sec | string | N | N | N | the 9-character unique ID for the security.  This is an internal identfier. |
| 11 | Pref Iss Nme | string | Y | N | N | the long description for the security |
| 12 | Txt Lvl1 Ast Desc | string | N | N | N | the most granular asset class/category for the security |
| 13 | Txt Lvl2 Ast Desc | string | N | N | N | the second most granular asset class/category for the security |
| 14 | Txt Lvl3 Ast Desc | string | N | N | N | the highest level asset class/category grouping for the security |
| 15 | Txt Lvl4 Ast Desc | string | N | N | N | Ignore |
| 16 | Cde Ms Rtg | string | Y | N | N | Morgan Stanley Research Analyst rating for the security (if applicable) |
| 17 | Nme Iss Typ | string | Y | N | N | product groupings |
| 18 | Nme Iss Sub Typ | string | Y | N | N | product groupings |
| 19 | Nme Asset Typ | string | Y | N | N | product groupings |
| 20 | Num Face Val | number | Y | N | N | Ignore |
| 21 | Dt2 Call | date | Y | N | N | the next date the security can by called by the issuer (if applicable) |
| 22 | Mat Exp Date | date | Y | N | N | the date the security is scheduled to mature (if applicable) |
| 23 | Pct Coupon | number | Y | N | N | the annual coupon percent for the security (if applicable) |
| 24 | Amt Avg Dly Price | number | Y | N | N | the average daily price for the security |
| 25 | Num Qty | number | N | N | N | the number of shares/units owned |
| 26 | Amt Mkt Val | number | N | N | N | the market value of the security held in the account |

## data_ACTIVITY
| Order | Field | Type | Nullable | Primary Key | Unique | Note |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | Client | string | N | N | N | This is the system/database name for the client that owns the account |
| 2 | Account Number | string | N | N | N | This is the account number |
| 3 | Sort Name | string | N | N | N | This is the ID_NAME_ALG |
| 4 | Activity Type | string | N | N | N | This the system defined category for the activity |
| 5 | Symbol | string | Y | N | N | This is the security symbol (if applicable) |
| 6 | CUSIP | string | Y | N | N | This is the security CUSIP (if applicable) |
| 7 | Activity Description | string | N | N | N | This is the long description for the activity |
| 8 | Quantity | number | N | N | N | This is the number of shares/units included in the transaction (if applicable) |
| 9 | Price($) | number | N | N | N | This is the price of the security at the time of the activity (if applicable) |
| 10 | Amount($) | number | N | N | N | This is the dollar amount associated with the activity |
| 11 | Activity Date | date | N | N | N | This is the date the activity occurred |
| 12 | Settlement Date | date | Y | N | N | This is the date the activity settled |
| 13 | Account Type | string | N | N | N | This is the type of account the activity took place in |
| 14 | Entity Type | string | N | N | N | Ignore |
| 15 | MSID | string | Y | N | N | This is the unique ID of the employee who posted the activity |