# Import necessary libraries
import pandas as pd
import numpy as np
from scipy.stats import linregress

# --- Single Exponential Smoothing ---
def singleES(inputData, alpha):
    if not 0 < alpha <= 1:
        raise ValueError("Alpha must be between 0 and 1!")
    forecastValues = [inputData[0]]
    for i in range(1, len(inputData) + 1):
        newForecast = alpha * inputData[i - 1] + (1 - alpha) * forecastValues[i - 1]
        forecastValues.append(newForecast)
    return forecastValues

# --- Double Exponential Smoothing ---
def doubleES(inputData, alpha, beta):
    dataLength = len(inputData)
    if dataLength < 2:
        raise ValueError("Data must contain at least two periods' data")
    level = [inputData[0]]
    trend = [inputData[1] - inputData[0]]
    forecastValues = ["N/A"] * 2
    for t in range(1, dataLength):
        updatedLevel = alpha * inputData[t] + (1 - alpha) * (level[t - 1] + trend[t - 1])
        level.append(updatedLevel)
        updatedTrend = beta * (level[t] - level[t - 1]) + (1 - beta) * trend[t - 1]
        trend.append(updatedTrend)
        forecastValues.append(updatedLevel + updatedTrend)
    return level, trend, forecastValues

# --- Triple Exponential Smoothing (Holt-Winters) ---
def tripleES(inputData, alpha, beta, gamma, seasonLength):
    trend = ["N/A"] * (2 * seasonLength - 1)
    trend.append((sum(inputData[seasonLength:2 * seasonLength]) - sum(inputData[0:seasonLength])) / seasonLength ** 2)

    seasonality = ["N/A"] * seasonLength
    for t in range(seasonLength):
        seasonAvg = inputData[t] / (sum(inputData[0:seasonLength]) / seasonLength) + inputData[seasonLength + t] / (sum(inputData[seasonLength:2 * seasonLength]) / seasonLength)
        seasonAvg /= 2
        seasonality.append(seasonAvg)

    level = ["N/A"] * (2 * seasonLength - 1)
    level.append(inputData[2 * seasonLength - 1] / seasonality[2 * seasonLength - 1])

    forecastValues = ["N/A"] * (2 * seasonLength)
    forecastValues.append((level[2 * seasonLength - 1] + trend[2 * seasonLength - 1]) * seasonality[seasonLength])

    for t in range(2 * seasonLength, len(inputData)):
        level.append(alpha * inputData[t] / seasonality[t - seasonLength] + (1 - alpha) * (level[t - 1] + trend[t - 1]))
        trend.append(beta * (level[t] - level[t - 1]) + (1 - beta) * trend[t - 1])
        seasonality.append(gamma * inputData[t] / level[t] + (1 - gamma) * seasonality[t - seasonLength])
        forecastValues.append((level[t] + trend[t]) * seasonality[t + 1 - seasonLength])

    return level, trend, seasonality, forecastValues

# --- Linear Regression ---
def linearRegression(xValues, yValues):
    x = np.array(xValues)
    y = np.array(yValues)
    slope, intercept, rValue, pValue, stdErr = linregress(x, y)
    fitLine = slope * x + intercept
    return fitLine, slope, intercept

# --- Performance Metrics ---
def calculateMetrics(predictions, actuals):
    dataLength = len(actuals)
    sumAbsolute = sumSquared = sumPercentage = 0.0
    for i in range(dataLength):
        if predictions[i] == "N/A":
            dataLength -= 1
        else:
            error = predictions[i] - actuals[i]
            sumAbsolute += abs(error)
            sumSquared += error ** 2
            sumPercentage += abs(error) / actuals[i]
    mad = sumAbsolute / dataLength
    mse = sumSquared / dataLength
    mape = sumPercentage / dataLength
    return mad, mse, mape

# --- Main Execution Block ---
dataFrame = pd.read_excel("HW 5 - ES and LR.xlsx")
monthList = dataFrame['Month'].tolist()
demandList = dataFrame['Demand'].tolist()

timePeriods = [i + 1 for i in range(len(monthList) + 1)] + ["MAD", "MSE", "MAPE"]
resultsFrame = pd.DataFrame()
resultsFrame['Month'] = timePeriods

forecastLR, slope, intercept = linearRegression(monthList, demandList)
forecastLR = np.append(forecastLR, intercept + slope * (len(monthList) + 1))
forecastLR = np.append(forecastLR, list(calculateMetrics(forecastLR, demandList)))
resultsFrame['LR'] = forecastLR

alpha = 0.5
forecastSES = singleES(demandList, alpha)
forecastSES += list(calculateMetrics(forecastSES, demandList))
resultsFrame['Single ES'] = forecastSES

beta = 0.35
levelDES, trendDES, forecastDES = doubleES(demandList, alpha, beta)
forecastDES += list(calculateMetrics(forecastDES, demandList))
resultsFrame['Double ES'] = forecastDES

gamma = 0.6
seasonLength = 4
levelTES, trendTES, seasonalityTES, forecastTES = tripleES(demandList, alpha, beta, gamma, seasonLength)
forecastTES += list(calculateMetrics(forecastTES, demandList))
resultsFrame['Triple ES'] = forecastTES

resultsFrame.to_excel("Results.xlsx", index=False)
