
---

# Compression Test Analysis Application

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![Open Source](https://img.shields.io/badge/Open%20Source-Yes-brightgreen)

A Python-based desktop application for analyzing compression test data. This tool automates the calculation of stress, strain, compression modulus, maximum stress, and energy up to 40% strain. It also generates a stress-strain curve and saves the results in an updated Excel file.
##Note : Follow "for test.xlsx" file basic require format
---

## Features

- **Load Excel Data**: Load compression test data from an Excel file.
- **Calculate Stress and Strain**: Automatically calculate stress and strain values based on input data.
- **Maximum Stress**: Identify and save the maximum stress value.
- **Compression Modulus**: Calculate and save the compression modulus (Ec).
- **Energy Calculation**: Calculate the energy under the stress-strain curve up to 40% strain.
- **Plot Stress-Strain Curve**: Generate and save a stress-strain curve as an image.
- **Save Results**: Save all calculations and plots back to the Excel file in the same directory as the input file.

---

## Requirements

- Python 3.8+
- Libraries:
  - `openpyxl`
  - `matplotlib`
  - `numpy`
  - `scipy`
  - `tkinter`

Install the required libraries using:
```bash
pip install openpyxl matplotlib numpy scipy
```

---

## How to Use

1. **Run the Application**:
   - Clone the repository:
     ```bash
     git clone https://github.com/AliAlAsif/Stress-Strain-Analysis-Using-Python-Excel/
     ```
   - Navigate to the project directory:
     ```bash
     cd compression-test-analysis
     ```
   - Run the script:
     ```bash
     python compression_test_app.py
     ```

2. **Load Excel File**:
   - Click the "Load Excel File" button to load your compression test data.

3. **Perform Calculations**:
   - Use the buttons to calculate stress and strain, maximum stress, compression modulus, and energy up to 40% strain.

4. **Plot Stress-Strain Curve**:
   - Click the "Plot Stress-Strain Curve" button to generate and save the plot.

5. **Save Results**:
   - All results are automatically saved in an updated Excel file in the same directory as the input file.

---

## Example Input File

Your Excel file should include the following columns:
- **Force**: Applied force values.
- **Length, L**: Length of the sample.
- **Width, W**: Width of the sample.
- **Stroke**: Displacement values.
- **Thickness, T**: Thickness of the sample.

---

### 📊 **Core Calculations**  
- **Stress (σ) = Force / (Length × Width)**  
- **Strain (ε) = (Stroke / Thickness) × 100**  
---

This version avoids formula formatting issues. You can now **copy-paste it smoothly** into LinkedIn. Let me know if you need any more fixes! 😊🚀

## Packaging the Application

To create a standalone executable:
1. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```
2. Package the application:
   ```bash
   pyinstaller --onefile --windowed compression_test_app.py
   ```
3. The executable will be located in the `dist` folder.

---


## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.

---

## Author

ALI AL ASIF
GitHub: (https://github.com/AliAlAsif/) 
Email: asif142636@gmail.com

---
