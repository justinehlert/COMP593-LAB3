""" 
Description: 
  Divides sales data CSV file into individual order data Excel files.

Usage:
  python process_sales_data.py sales_csv_path

Parameters:
  sales_csv_path = Full path of the sales data CSV file
"""
import pandas as pd
from sys import argv
import re
import os
from datetime import datetime
import xlsxwriter
from openpyxl import load_workbook

def main():
    sales_csv_path = get_sales_csv_path()
    orders_dir_path = create_orders_dir(sales_csv_path)
    process_sales_data(sales_csv_path, orders_dir_path)

def get_sales_csv_path():    
    """Gets the path of sales data CSV file from the command line

    Returns:
        str: Path of sales data CSV file
    """
    # TODO: Check whether command line parameter provided
    try:
        file_path = argv[1]
    except IndexError:
        print("Error: Please provide the file path")
        quit()
    # TODO: Check whether provide parameter is valid path of file
    try:
        open(file_path, 'r')
    except:
        print("Error: File not found")
        quit()
    # TODO: Return path of sales data CSV file
    return file_path

def create_orders_dir(sales_csv_path):
    """Creates the directory to hold the individual order Excel sheets

    Args:
        sales_csv_path (str): Path of sales data CSV file

    Returns:
        str: Path of orders directory
    """
    # TODO: Get directory in which sales data CSV file resides
    salesParentDirectory = os.path.abspath(os.path.join(sales_csv_path, os.pardir))
    # TODO: Determine the path of the directory to hold the order data files
    match = re.search(r".*^\S+", str(datetime.now()))
    today = match.group()
    # TODO: Create the orders directory if it does not already exist
    orderDir = f"{salesParentDirectory}\\Orders_{today}"
    try:
        os.mkdir(orderDir)
    except OSError:
        pass
    # TODO: Return path of orders directory
    return orderDir

def process_sales_data(sales_csv_path, orders_dir_path):
    """Splits the sales data into individual orders and save to Excel sheets

    Args: 
        sales_csv_path (str): Path of sales data CSV file
        orders_dir_path (str): Path of orders directory
    """
    # TODO: Import the sales data from the CSV file into a DataFrame
    salesDataFrame = pd.read_csv(sales_csv_path)
    # TODO: Insert a new "TOTAL PRICE" column into the DataFrame
    salesDataFrame.insert(7, 'TOTAL PRICE', salesDataFrame['ITEM QUANTITY'] * salesDataFrame['ITEM PRICE'])
    # TODO: Remove columns from the DataFrame that are not needed
    salesDataFrame = salesDataFrame.drop(['ADDRESS', 'CITY', 'STATE', 'POSTAL CODE', 'COUNTRY'], axis=1)

    # TODO: Groups orders by ID and iterate 
    salesDataFrame = salesDataFrame.drop_duplicates(subset=['ORDER ID'])

        # TODO: Remove the 'ORDER ID' column
    salesDataFrame = salesDataFrame.drop(['ORDER ID'], axis=1)
        # TODO: Sort the items by item number
    salesDataFrame = salesDataFrame.sort_values('ITEM NUMBER')
        # TODO: Append a "GRAND TOTAL" row
    salesDataFrame.at['GRAND TOTAL', 'TOTAL PRICE'] = salesDataFrame['TOTAL PRICE'].sum()
    print(salesDataFrame)
        # TODO: Determine the file name and full path of the Excel sheet
    excelFilePath = f'{orders_dir_path}\\sales_csv.xlsx'
    worksheet = 'Sales Info'
        # TODO: Export the data to an Excel sheet
    salesDataFrame.to_excel(excelFilePath, index=False, sheet_name=worksheet)
        # TODO: Format the Excel sheet
    workbook = load_workbook(excelFilePath)
    worksheet = workbook.active
        # TODO: Define format for the money columns
    for col in worksheet.iter_cols(min_row=1, min_col=6, max_col=7):
        for cell in col:
            cell.number_format = '$#,##0.00' 
        # TODO: Format each colunm
    worksheet.column_dimensions.width = 80
    workbook.save(excelFilePath)
        # TODO: Close the Excelwriter 
    return

if __name__ == '__main__':
    main()