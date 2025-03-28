# Forecasting Techniques Implementation

This repository provides Python code for implementing standard time series forecasting models:

- Single Exponential Smoothing (SES)
- Double Exponential Smoothing (DES)
- Triple Exponential Smoothing (Holt-Winters Method)
- Linear Regression (LR)

## Files Included

1. `forecastModels.py` - Python script implementing the forecasting models with camelCase variables.
2. `Results.xlsx` - Excel file containing the forecast results and model performance metrics.

## Methodology

The code manually implements the forecasting algorithms using only core Python libraries like NumPy, SciPy, and Pandas.

For each model, the following error metrics are calculated:
- Mean Absolute Deviation (MAD)
- Mean Squared Error (MSE)
- Mean Absolute Percentage Error (MAPE)

The forecasted values and metrics are saved in `Results.xlsx`.

## How to Run

1. Install the required libraries:
```
pip install pandas numpy scipy openpyxl
```

2. Make sure the input file `HW 5 - ES and LR.xlsx` is placed in the same directory.

3. Run the script:
```
python forecastModels.py
```

## GitHub Repository

Access the full implementation here:  
https://github.com/ac733s/HWForecast

## Author

Developed by ac733s for educational use. Contributions are welcome.
