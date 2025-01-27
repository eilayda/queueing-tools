# Queuing Model Tool
## Overview
This project is a Queuing Model Tool (QMT) designed as part of the IE 301 - Operations Research II course. The tool allows users to perform queuing model analyses using Excel VBA. It provides a user-friendly interface for selecting queuing models, entering parameters, and calculating performance measures. 

## Features
- **Queuing Model Selection**: Choose from multiple queuing models via a drop-down menu:
  - M/M/1:GD/∞/∞
  - M/M/1:GD/N/∞
  - M/M/c:GD/∞/∞
  - M/M/c:GD/N/∞
  - M/M/∞:GD/∞/∞ (self-service model)
  - M/M/R:GD/K/K (repair-shop model)
  - M/G/1:GD/∞/∞ (P-K formula)

- **Dynamic Input Validation**: Prompts for model-specific input fields and ensures parameters like \(\lambda\), \(\mu\), and others are positive.

- **Performance Metrics**: Calculates steady-state measures including:
  - (p_0, p_n) (for (n < 20))
  - For limited N models: (p_N, lambda_eff, lambda_lost)
  - (L_s, L_q, W_s, W_q)

## **Error Handling**

Our tool includes robust error-checking mechanisms to ensure accurate data input and prevent calculation issues:

Model Selection Validation: If no model is selected from the drop-down menu, an error message box displays:
"Please select a model."

Input Validation: Parameters such as 𝜆 (arrival rate) and  𝜇 (service rate) are checked for validity. If negative or invalid values are entered, the tool warns with messages like:
"Please enter positive values for Lambda and Mu."

Parameter-Specific Warnings: The error messages adapt to the specific input requirements of each model.
Resetting Data: The "Clear Data" macro ensures that all input fields are reset, preventing conflicts between calculations for different models.

- **Clear Data Macro**: Allows resetting inputs and outputs with a single click.

## Files in Repository
- **IE301_Report.docx**: Project documentation with details on implementation, user forms, macros, and error-handling mechanisms.
- **QM_Model.xlsm**: The main Excel file containing VBA macros and user forms for the queuing tool.

## How to Use
1. Open `QM_Model.xlsm` in Microsoft Excel.
2. Enable macros to activate VBA functionality.
3. Select a queuing model from the drop-down menu.
4. Enter the required parameters based on the selected model.
5. Click "Calculate" to view performance metrics in the results section.
6. Use the "Clear Data" button to reset the form for a new analysis.

## Acknowledgments
This project was developed under the guidance of our IE 301 instructor, who provided the framework and scope for the assignment.

---
