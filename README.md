# sales_moving_average
Sales Data Analysis Tool
Description

This Python script reads sales data from an Excel file, processes it, and generates a plot showing sales trends along with moving averages. The script allows users to input the filename (excluding the extension) and handles missing data by filling gaps with the mean of nearby values.
Prerequisites

    Python 3.x
    Required Python packages: openpyxl, matplotlib, numpy

Usage

    Run the script by executing the following command in your terminal:

    bash

    python sales_analysis.py

    Enter the filename (excluding the extension) when prompted.

Input File Format

The script expects an Excel file with the following structure:

    Column A: Date
    Column D: Sales data

Handling Missing Data

The script identifies zero sales values and fills them using a mean regression over a 7-day window.
Moving Averages

The script calculates and plots 10-day, 30-day, 90-day, 180-day, and 360-day moving averages.
Troubleshooting

    If the script encounters a FileNotFoundError, double-check the file path.

Note

This script assumes a specific file structure and may need adjustments based on the format of your data.
