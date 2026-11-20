# PediaZ

**PediaZ** (Pediatric Z-score Calculator) is a free, open-source, Python-based desktop tool for automated calculation of WHO-based anthropometric z-scores (height-for-age, weight-for-age, weight-for-height and BMI-for-age) and prevalence reporting for epidemiological and clinical research, from birth to 19 years of age.

## Features

* Calculates HFA, WFA, WFH and BFA z-scores from the WHO Child Growth Standards (0-5 years) and the WHO 2007 Growth Reference (5-19 years), using the LMS method with linear interpolation by age
* Applies the WHO adjustment for z-scores beyond +/-3 (WFA, WFH and BFA) and indicator-specific WHO plausibility limits; z-scores are rounded to two decimals before flagging and classification, as in WHO Anthro and AnthroPlus
* Smart column matching: if your column headings differ from the defaults, PediaZ proposes a match and shows it in four combo boxes (sex, age, weight, height). Each box lists all headings in your file, so you can confirm or correct the proposal
* Warns when age or height appear to be in years or meters and offers to convert them
* Processes large datasets quickly using vectorized operations (NumPy/pandas)
* Generates prevalence tables for the total sample and the clean sample (excluding missing and flagged records), with the effective sample size and 95% confidence intervals (Wald and Wilson), by sex and age group
* Compares the z-score distribution of the clean sample with the WHO standard, N(0,1): charts (overall, by sex, <61 months and >=61 months) and one-sample tests (mean, SD, overall shape, and tail percentages) with raw and Holm-adjusted p-values
* Produces data-quality and input-check reports
* User-friendly GUI with Tkinter

## Requirements

* Windows 7/8/10/11, macOS, or Linux (via Python source)
* Python 3.13 or higher (with Tkinter)
* Required libraries: NumPy, pandas, OpenPyXL, XlsxWriter (SciPy is not required; xlrd is needed only for old .xls files)
* The `WHO\_LMS` folder (WHO LMS tables, 20 files) must be in the same folder as the script

## Installation

```
pip install -r requirements.txt
```

## Usage

Run the main script:

```
python "PediaZ v.01.py"
```

The program asks you to select the input file, confirm the variables, and choose where to save the output.

## Input format

An Excel (.xlsx, .xls), CSV or tab-delimited file with one row per child and four variables. Column names do not have to follow a fixed convention (you confirm them in the program), but the defaults are:

* `sex` (f/m; also female/male, girl/boy, 1 = boy, 2 = girl)
* `age\_months` (numeric, completed or decimal months)
* `weight\_kg` (numeric)
* `height\_cm` (numeric; length for children younger than 24 months)

Age ranges: weight-for-age is calculated up to 119 months, weight-for-height up to 60 months, and the other indicators up to 228 months.

## Output

A single Excel file containing:

* `z\_scores`: individual BMI and z-scores with interpretations
* `HFA\_prevalence`, `WFA\_prevalence`, `WFH\_prevalence`, `BFA\_prevalence`: prevalence tables for the total and clean samples, stratified by sex and age, with the comparison to the WHO standard (charts and tests) beside each block
* `data\_quality`: missing, flagged and valid records for each indicator
* `input\_checks`: results of the input checks
* `readme`: software version, column assignment and methods

Note: the p-values of the distribution tests assume simple random sampling. For clustered or weighted survey data they should be read descriptively.

## Validation

In NHANES 2021-2023 (children and adolescents aged 24 months to 18 years), PediaZ matched the WHO Anthro and AnthroPlus reports in all compared statistics (valid N, prevalence at the standard cut-offs, and mean and SD of z-scores, by sex and age group). Infants younger than 24 months and individual-level z-scores have not been compared yet.

## License

This project is licensed under the GNU General Public License v3.0.

## Citation

If you use PediaZ in your research, for citation please search for "PediaZ" in Google Scholar.

