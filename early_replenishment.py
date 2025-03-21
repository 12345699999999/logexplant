# -*- coding: utf-8 -*-
"""Early Replenishment.ipynb

Automatically generated by Colab.

Original file is located at
    https://colab.research.google.com/drive/1vIG3cFXpycSZ5H5l-DTN6vDwUFEyB-HS
"""

from google.colab import files
import pandas as pd
import math
import os

def determine_storage_area(storage_bin):
    if storage_bin.startswith("BKT"):
        return "BAKTI"
    elif storage_bin.startswith("ARG"):
        return "ARGO"
    else:
        return "TAS"

def process_shipments(shipments, stock, master):
    # Convert NaN values to empty string
    stock['S. Cat'] = stock['S. Cat'].fillna("")

    # Fill in the storage area of each line of stock
    stock['Storage Area'] = stock['S. Bin'].apply(determine_storage_area)

    # Filter the data only for stocks with stock category UU & Q, and in valid storage types
    stock = stock[stock['S. Cat'].isin(["", "Q"]) & stock['S. Type'].isin(["Z0A", "Z0C", "ZBF", "ZFR"])]

    # Ensure that the data type for stock quantity is float
    stock['Case Qty'] = pd.to_numeric(stock['Case Qty'], errors='coerce') # Transform the Case Qty field intu numeric value
    stock['Case Qty'] = stock['Case Qty'].astype(float) # Transform the Case Qty data type into float

    # Create a grouped stock DataFrame before iteration & merge with UPP master data
    grouped_stock = stock.groupby(['Material', 'Material Description', 'Storage Area', 'S. Cat'])[['Case Qty']].sum().reset_index()
    grouped_stock = grouped_stock.merge(master, on='Material Description', how='left', suffixes=('','_master'))

    replenishment_list = [] # Create an empty list for replenishment list

    for _, shipment in shipments.iterrows():
        material_id = shipment['Material']
        required_qty = shipment['Delivery quantity']

        available_stock = grouped_stock[(grouped_stock['Material'] == material_id) & (grouped_stock['Case Qty'] > 0)]

        tas_stock = available_stock[(available_stock['Storage Area'] == "TAS") & (available_stock['S. Cat'] == "")]
        in_plant_qty = tas_stock['Case Qty'].sum()

        if in_plant_qty >= required_qty:
            grouped_stock.loc[(grouped_stock['Material'] == material_id) & (grouped_stock['Storage Area'] == "TAS") & (grouped_stock['S. Cat'] == ""), 'Case Qty'] -= required_qty
            continue

        required_qty -= in_plant_qty
        external_stock = available_stock[(available_stock['Storage Area'].isin(["ARGO", "BAKTI"])) & (available_stock['S. Cat'] == "")]

        for _, row in external_stock.iterrows():
            if required_qty <= 0:
                break
            qty_to_move = min(required_qty, row['Case Qty'])
            replenishment_list.append([row['Storage Area'], material_id, row['Material Description'], row['S. Cat'], qty_to_move, math.ceil(qty_to_move/row['UPP']*100)/100])
            grouped_stock.loc[(grouped_stock['Material'] == material_id) & (grouped_stock['Storage Area'] == row['Storage Area']) & (grouped_stock['S. Cat'] == ""), 'Case Qty'] -= qty_to_move
            required_qty -= qty_to_move

        if required_qty > 0:
            q_stock = available_stock[(available_stock['S. Cat'] == "Q")].sort_values(by=['Case Qty'])
            for _, row in q_stock.iterrows():
                if required_qty <= 0:
                    break
                qty_to_move = min(required_qty, row['Case Qty'])
                replenishment_list.append([row['Storage Area'], material_id, row['Material Description'], row['S. Cat'], qty_to_move, math.ceil(qty_to_move/row['UPP']*100)/100])
                grouped_stock.loc[(grouped_stock['Material'] == material_id) & (grouped_stock['Storage Area'] == row['Storage Area']) & (grouped_stock['S. Cat'] == "Q"), 'Case Qty'] -= qty_to_move
                required_qty -= qty_to_move

    output_df = pd.DataFrame(replenishment_list, columns=["Storage Area", "Material", "Material Descriptiion", "S. Cat", "Replenishment Quantity (in Box)", "Replenishment Quantity (in Pallet)"])
    return output_df, stock

def main():
    # Define default file path
    print(f"Please upload the required file:")
    file_path = "/content/Sample Data"

    # Delete existing file
    if os.path.exists(file_path):
      os.remove(file_path)
      print(f"{file_path} has been deleted.")
    else:
      print(f"{file_path} does not exist.")

    input_file = files.upload() # Upload required excel file
    file_name = list(input_file.keys())[0]  # Get file name
    os.rename(file_name, "Sample Data")
    print("File renamed successfully!")
    print(file_name)


    xls = pd.ExcelFile(file_path)
    shipments = pd.read_excel(xls, sheet_name="Shipments")
    stock = pd.read_excel(xls, sheet_name="Stock")
    master = pd.read_excel(xls, sheet_name="Master")

    result_df, updated_stock = process_shipments(shipments, stock, master)
    result_df = result_df.groupby(['Storage Area', 'Material', 'Material Descriptiion', 'S. Cat'])[['Replenishment Quantity (in Box)', 'Replenishment Quantity (in Pallet)']].sum().reset_index()

    if not result_df.empty:
        output_file = "replenishment_plan.csv"
        result_df.to_csv(output_file, index=False)
        print(f"Replenishment plan saved as {output_file}")

        stock_output_file = "updated_stock.csv"
        updated_stock.to_csv(stock_output_file, index=False)
        print(f"Updated stock saved as {stock_output_file}")
    else:
        print("No replenishment needed.")


if __name__ == "__main__":
    main()