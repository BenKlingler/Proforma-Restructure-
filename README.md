# Proforma-Restructure-
## Overview

This document provides an overview of VBA macros designed to transfer and process data within an Excel workbook containing multiple sheets, specifically "UPCOMING PROJECTS", "Data", and "Revenue Summary".

## Macros Summary

### Macro: `FindMatchingDataAndInsertRow`

**Purpose:** 

Automates the process of transferring project details from "UPCOMING PROJECTS" to the "Data" sheet. It locates projects based on a 5-character and a 2-character code, then inserts a new row in "Data" with the project's details.

**Steps:**

1. Fetches codes from "Dummy" and finds corresponding rows in "Data".
2. Inserts a new row below the last match in "Data".
3. Copies data and formulas down from the row above, with special handling for certain columns.
4. Utilizes `AutoFill` for formula columns P to AC.
5. Forces a sheet recalculation to update formulas.

### Macro: `FindMatchingDataAndInsertRowInProportion`

**Purpose:** 

Transfers specific project details from "UPCOMING PROJECTS" to "Revenue Summary" based on matching codes. It also handles non-numeric or error-containing cells from "Data" and propagates associated details to "Revenue Summary".

**Steps:**

1. Matches project codes between "Dummy" and "Proportion".
2. Inserts a new row in "Proportion" under the matched project.
3. Searches "Data" for non-numeric/error cells and transfers adjacent details to "Proportion".
4. Drags down formulas in specified columns, if the above row contains formulas.

### Macro: `ForecastMonthlySpending`

**Purpose:** 

Calculates and forecasts monthly spending based on project start and end dates within "Revenue Summary". It distributes the total contract value over the project duration.

**Steps:**

1. Identifies the appropriate row in "Revenue Summary" based on contract type.
2. Calculates the number of months between project start and end dates.
3. Distributes the monthly budget across the identified time frame.
4. Transfers "SCOPE NOTES" from "UPCOMING PROJECTS" to "Revenue Summary".

### Macro: `DistributeEstimatedHours`

**Purpose:** 

Distributes estimated hours for a project across multiple categories in "Revenue Summary".

**Steps:**

1. Splits the estimated total hours evenly across categories A to E.
2. Copies formulas for categories F and G from the row above to the current row.

### Macro: `UpdateTotalFormulas`

**Purpose:** 

Updates the "GRAND TOTAL" formulas in "Revenue Summary" to ensure accurate summation based on current data.

**Steps:**

1. Identifies the "GRAND TOTAL" row.
2. Updates SUM formulas to include only the rows above the summary section.
3. Prints debug information to the Immediate Window in the VBA editor for verification.
