
**PediaZ** (Pediatric Z-score Calculator) is a Python-based tool for automated calculation of WHO-based anthropometric z-scores (weight-for-age, height-for-age, weight-for-height, and BMI-for-age) for epidemiological and clinical research.

## Features
- Calculates HFA, WFA, WFH, and BFA z-scores based on WHO LMS parameters
- Processes large datasets quickly using vectorized operations (NumPy/Pandas)
- Generates prevalence tables with 95% confidence intervals (Wald and Wilson)
- Produces data quality reports
- User-friendly GUI with Tkinter

## Requirements
- Windows 7/8/10/11, macOS, or Linux (via Python source)
- Python 3.13 or higher
- Required libraries: NumPy, Pandas, SciPy, OpenPyXL, XlsxWriter

## Installation
pip install -r requirements.txt
Usage
Run the main script:
python "PediaZ v.01.py"

Input Format
The software expects an Excel file with the following columns:
sex (f/m)
age_months (numeric)
weight_kg (numeric)
height_cm (numeric)

Output
A single Excel file containing:

Individual z-scores with interpretations
Prevalence tables stratified by age and sex
Data quality metrics

License
This project is licensed under the GNU General Public License v3.0.

Citation
If you use PediaZ in your research, please cite. search in google scholar for: Sharifi A. Development and Validation of WHO-Based Software for Automated Pediatric Growth Assessment in Epidemiological Research.
