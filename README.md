Here’s a **README** for the collection of VBA macros you’ve been working with:

---

# VBA Macros for Managing Data in Excel

This repository contains a collection of VBA macros designed to manage and manipulate data in Excel sheets. These macros streamline processes like comparing and cleaning data, filtering, and exporting results, helping automate data processing tasks for efficient workflow management.

## Table of Contents

- [Macros Overview](#macros-overview)
- [Setup and Usage](#setup-and-usage)
- [Individual Macros](#individual-macros)
  - [FindMissingNRICsToColumnC](#findmissingnricstocolumnc)
  - [CleanUpNRICsInColumnA](#cleanupnricsincolumna)
  - [ExportMissingDataToOutput](#exportmissingdatatooutput)
  - [Whitelist Check](#whitelist-check)
  - [AddColumnDataFromPd](#addcolumndatafrompd)
  - [CleanupOutputSheet](#cleanupoutputsheet)
- [Testing Data](#testing-data)

---

## Macros Overview

These macros are intended to work on sheets that contain personnel data, typically formatted with columns for NRIC, name, email, status, and more. They allow for easy searching, matching, cleaning, and exporting of data between sheets.

## Setup and Usage

1. **Add Macros to Excel**:
   - Open Excel.
   - Press `Alt + F11` to open the VBA editor.
   - Insert a new module (`Insert > Module`).
   - Paste the relevant macro code into the module.

2. **Running Macros**:
   - You can run macros from the VBA editor (`F5` or right-click and choose `Run`).
   - Alternatively, you can assign these macros to buttons in the worksheet for ease of use.

---

## Individual Macros

### **FindMissingNRICsToColumnC**
This macro checks for NRICs in the `rawNR` sheet that are missing from column A of the active sheet. It outputs the missing NRICs in **column C**.

#### Features:
- Compares NRICs in the active sheet and `rawNR` sheet.
- Outputs missing NRICs to column C.

---

### **CleanUpNRICsInColumnA**
This macro clears all the NRIC data in **column A** from row 2 to the last row with data, leaving the header intact.

#### Features:
- Clears only the data in column A (NRICs) while keeping the header in row 1.

---

### **ExportMissingDataToOutput**
This macro finds missing NRICs from the active sheet and exports the relevant data (NRIC, name, rank, etc.) to a sheet called **`output`**.

#### Features:
- Exports missing NRICs from the active sheet.
- Maps related data from `rawNR` to the `output` sheet.

---

### **Whitelist Check**
This macro prompts the user to enter their name, checks if the name exists in the `whitelist` sheet, and ensures only authorized users can run other macros.

#### Features:
- Verifies user name before allowing access to macro functionalities.
- Can be integrated into other macros to control access.

---

### **AddColumnDataFromPd**
This macro appends data from the **`pd`** sheet (columns Q and S) to the **`output`** sheet (columns M and N), based on matching NRICs.

#### Features:
- Appends relevant data from the `pd` sheet.
- Ensures columns Q and S from `pd` are mapped to the correct columns in `output`.

---

### **CleanupOutputSheet**
This macro prompts the user with a confirmation message asking if they want to clean up the `output` sheet, clearing data in the NRIC, name, rank, and other columns while preserving the headers.

#### Features:
- Cleans up the `output` sheet after processing.
- Clears all data except for the header row.

---

## Testing Data

You can use the following test data to ensure the macros work as expected:

### **Test Data for `rawNR` Sheet**

| **NRIC**   | **NAME**  | **ADDRESS**       | **EMAIL**            | **RANK** |
|------------|-----------|-------------------|----------------------|----------|
| S1234567A  | John Doe  | 123 Example St    | johndoe@hr.com       | HR Specialist |
| T7654321B  | Jane Smith| 456 Sample Road   | janesmith@eng.com    | Engineer |
| G8765432D  | Alex Brown| 789 Placeholder Ln| alexbrown@int.com    | Intern |
| F2345678C  | Emily White| 789 Mock Ave      | emilywhite@ld.com    | Lead |

---

### **Test Data for `output` Sheet**

| **NRIC**   | **NAME**      | **RANK**         | **EMAIL**            | **ADDRESS**       |
|------------|---------------|------------------|----------------------|-------------------|
| S1234567A  | John Doe      | HR Specialist    | johndoe@hr.com       | 123 Example St    |
| T7654321B  | Jane Smith    | Engineer         | janesmith@eng.com    | 456 Sample Road   |

---

### **Test Data for `pd` Sheet**

| **NRIC**   | **Q (Output Column N)**  | **S (Output Column M)** |
|------------|--------------------------|-------------------------|
| S1234567A  | Completed Training       | Eligible for Promotion |
| T7654321B  | Exceeds Expectations     | Consider for Bonus     |
| G8765432D  | Pending Evaluation       | On Hold                |
| F2345678C  | Leadership Certified     | Eligible for Award     |
| S9876543E  | Top Performer            | Recommended for Raise  |

---

### License

This project is open-source. Feel free to use, modify, and distribute the code under the terms of the MIT License.

---

This **README** provides an overview of the macros, their functionality, and test data to ensure everything works as expected. Let me know if any further details are needed!