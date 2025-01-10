Macro Automation for NRIC Processing
This project contains a set of Excel VBA macros designed for managing and processing NRIC data, with additional user access control through a whitelist. The macros assist with tasks such as finding missing NRICs, cleaning up data, and outputting results in an organized manner, while restricting access to only authorized users.

Macros Included
1. Whitelist Check
Macro Name: WhitelistCheckAndRun

Function: Prompts the user to enter their name and checks if it exists in a predefined whitelist.
Whitelist Sheet: The whitelist of names should be stored in a sheet called whitelist with the names starting from A2.
Access Control: If the name is not found in the whitelist, the macro will prevent execution and display a notification. If the name is found, the user will be authorized to run the other macros.
Usage:

Run this macro before executing any other macros to ensure that only authorized users can proceed.
2. Find Missing NRICs and Output to Column C
Macro Name: FindMissingNRICsToColumnC

Function: This macro compares the NRICs listed in the current sheet (Column A) against a separate data set in the rawNR sheet, then outputs any missing NRICs into Column C.
Input: The current sheet should have a list of NRICs in Column A.
Output: Missing NRICs will be displayed starting from cell C2 in the current sheet.
Usage:

Run this macro to identify and display missing NRICs in the active sheet.
3. Output Missing NRICs with Associated Data
Macro Name: OutputMissingNRICsWithData

Function: This macro outputs missing NRICs (identified from Column C) along with associated data (Name, Email, Address) from the rawNR sheet to the output sheet.
Input: Missing NRICs should be listed in Column C of the current sheet.
Output: The macro will populate the output sheet with the missing NRICs and corresponding details.
Usage:

Run this macro after identifying missing NRICs in Column C to output the relevant data to the output sheet.
4. Clean Up Output Sheet
Macro Name: CleanUpOutputSheet

Function: This macro clears the data in the output sheet, keeping only the header row intact.
Output: All data from row 2 onwards (in columns A-D) will be cleared.
Usage:

Run this macro after completing a task to clean up the output sheet and prepare it for new data.
How to Use the Macros
Ensure the whitelist sheet is set up:

Create a sheet called whitelist.
Enter the names of authorized users in Column A, starting from A2.
Run the Whitelist Check:

Run the WhitelistCheckAndRun macro to verify user access. The user will be prompted to enter their name.
Run the Other Macros:

After passing the whitelist check, you can run any of the other macros:
FindMissingNRICsToColumnC to find missing NRICs.
OutputMissingNRICsWithData to output the missing NRICs with associated data to the output sheet.
CleanUpOutputSheet to clear the output sheet.
Dependencies
Excel (VBA-enabled) required for running the macros.
The following sheets must be present:
whitelist: For storing authorized names.
rawNR: For storing NRIC and associated data (NRIC, Name, Address, Email).
output: For displaying output results.
Customization
Modify the whitelist sheet to add or remove authorized users.
The columns in the rawNR sheet (for NRIC, Name, Address, Email) can be customized based on your specific data structure, but the macro assumes they are in the same format.
The output sheet is customizable for other types of data display or analysis.
Notes
Ensure the whitelist sheet has the header Names in cell A1.
Column references (like A, C, etc.) in the macros are based on the assumption that the data structure in your sheets follows the given format.
You can easily add more macros or adjust the existing ones as per your requirements.
