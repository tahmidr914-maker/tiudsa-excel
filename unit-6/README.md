# Unit 6 - Lookup Functions (VLOOKUP, XLOOKUP, INDEX/MATCH)

## Overview

This project demonstrates the use of Excel lookup functions to retrieve customer information from a customer dataset. The objective was to perform lookups using VLOOKUP, XLOOKUP, and INDEX/MATCH while understanding the advantages and limitations of each method.

## Dataset

* Customer Dataset: Contains customer IDs, customer names, balances, purchase information, and other customer-related attributes.
* Orders Sheet: Created to simulate customer orders using Customer IDs from the customer dataset.

## Tasks Completed

### 1. VLOOKUP

Used the VLOOKUP function to retrieve customer information from the Customer Dataset based on Customer ID.

Example:
=VLOOKUP(Customer_ID, Customer_Dataset_Range, Column_Number, FALSE)

### 2. XLOOKUP

Used the XLOOKUP function to perform the same lookup and verify that the results matched those returned by VLOOKUP.

Example:
=XLOOKUP(Customer_ID, Lookup_Array, Return_Array)

### 3. INDEX/MATCH

Used the INDEX and MATCH functions together to perform a left-side lookup, a task that cannot be performed directly using VLOOKUP.

Example:
=INDEX(Return_Range, MATCH(Lookup_Value, Lookup_Range, 0))

### 4. XLOOKUP Advantage Comment

Added a comment explaining that XLOOKUP is more flexible because it:

* Can search both left and right
* Does not require a column index number
* Is easier to maintain when columns are inserted, removed, or rearranged

## Files Included

Customer DataSet_For LOOKUP Function.xlsx

Contains:

* Customer Dataset sheet
* Orders sheet
* VLOOKUP implementation
* XLOOKUP implementation
* INDEX/MATCH implementation
* XLOOKUP explanation comment

## Learning Outcomes

Through this exercise, I learned how different lookup functions work in Excel, when to use each function, and why XLOOKUP and INDEX/MATCH are often preferred over VLOOKUP for more advanced lookup tasks.
