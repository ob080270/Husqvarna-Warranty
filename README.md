# Husqvarna-Warranty

Welcome to the **Husqvarna-Warranty** repository. This project is a Microsoft Access-based system designed to work with warranty claims for Husqvarna motorcycles. Below is an overview of the main features and the project structure.

## Table of Contents

- [Description](#description)
- [Main Features](#main-features)
- [Screenshots](#screenshots)
- [Project Structure](#project-structure)
  - [Forms Folder](#forms-folder)
  - [Modules Folder](#modules-folder)
- [How to Run](#how-to-run)

---

## Description

**Husqvarna-Warranty** is an Access application that implements forms, queries, reports, and VBA modules for:

1. Tracking warranty claims.
2. Automating processes for handling warranty requests.
3. Generating various printable forms (labels, reports).

The project contains form files (with `.cls` and `.txt` extensions) and VBA modules (with `.bas` files), as well as additional objects required for the application to work correctly.

---

## Main Features

1. **Creating and managing warranty claims**  
   - The **frAP** form (the main form) includes fields for entering and editing warranty cases.
   - Subforms (**sfClm**, **sfItm**, etc.) allow you to detail information such as repairs, parts, and operations.

2. **Defect code selection**  
   - **frDfSelect** is a form for selecting defect codes (Defect Codes).
   - It can switch languages (Russian/English) for structural groups and symptom descriptions, and also filter lists by pattern matching.
   - By clicking a button, you can **transfer** the selected code to the main form (see screenshot HQ_17_TransferToClaim.jpg).

3. **Warranty label printing**  
   - The project provides reports (for example, **Report_rpBirka.cls**) that can be generated for printing:
     - You need to check the "causal part" box.
     - A warranty label is automatically displayed in preview mode. You can then print it (see HQ_20_WarrantyLabelGet.jpg and HQ_21_WarrantyLabel.jpg).

4. **Managing reference data**  
   - **frCust** — reference for customers/buyers.
   - VBA modules (**GlobalFn.bas**, **modWarntLR.bas**, **SumProp.bas**) — global functions, settings, and business logic.

---

## Screenshots

1. **HQ_000MainForm.jpg** — the main application screen.
2. **HQ_001.jpg, HQ_002.jpg, HQ_003.jpg** — interacting with the warranty case form.
3. **HQ_010EditMode.jpg** — the form displayed in edit mode.
4. **HQ_011_DefectCodeSelect.jpg** — the defect code selection window.
5. **HQ_012_DefectCodeSelectionRu.jpg** — selecting a defect code (Russian localization).
6. **HQ_013_DefectCodeSelectionEn.jpg** — selecting a defect code (English localization).
7. **HQ_014_SymptomFilterRu.jpg** — symptom filtering (Russian localization).
8. **HQ_015_SymptomFilterEn.jpg** — symptom filtering (English localization).
9. **HQ_16_ClearFilter.jpg** — example of clearing the symptom search box.
10. **HQ_17_TransferToClaim.jpg** — transferring the selected code to the warranty claim form.
11. **HQ_20_WarrantyLabelGet.jpg** — switching to edit mode and printing the label.
12. **HQ_21_WarrantyLabel.jpg** — example of a printed warranty label.

---

## Project Structure

### Forms Folder

This folder holds forms (and/or their text versions), as well as screenshots:
- Examples: `frWA`, `frCust`, `sfClm`, `sfItm`, `sfVhc`, and other forms, each corresponding to a particular interface section.

### Modules Folder

This folder contains the VBA modules and other files related to program logic:
- **GlobalFn.bas** — global functions used throughout the project.
- **modWarntLR.bas** — modules for calculations, settings, or auxiliary functions.
- **SumProp.bas** — conversion of a numeric amount into words (for Ukrainian accounting, etc.).

---

## How to Run

1. **Open** the Husqvarna-Warranty project in **Microsoft Access**.
2. If necessary, enable VBA macros (in Security Settings).
3. Open the **frAP** form (or another main form, if specified in the documentation).

