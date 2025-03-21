# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import math
from io import BytesIO

def determine_storage_area(storage_bin):
    if storage_bin.startswith("BKT"):
        return "BAKTI"
    elif storage_bin.startswith("ARG"):
        return "ARGO"
    else:
        return "TAS"

def process_shipments(shipments, stock, master):
    stock['S. Cat'] = stock['S. Cat'].fillna("")

    stock['Storage Area'] = stock['S. Bin'].apply(determine_storage_area)
    stock = stock[stock['S. Cat'].isin(["", "Q"]) & stock['S. Type'].isin(["Z0A", "Z0C", "ZBF", "ZFR"])]

    stock['Case Qty'] = pd.to_numeric(stock['Case Qty'], errors='coerce')
    stock['Case Qty'] = stock['Case Qty'].astype(float)

    grouped_stock = stock.groupby(['Material', 'Material Description', 'Storage Area', 'S. Cat'])[['Case Qty']].sum().reset_index()
    grouped_stock = grouped_stock.merge(master, on='Material Description', how='left', suffixes=('','_master'))

    replenishment_list = []

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

def convert_df_to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Replenishment Plan')
    writer.close()  # Correct method to close the writer
    output.seek(0)  # Move to the beginning of the BytesIO object
    return output

def main():
    st.title("Replenishment Plan Generator")

    # File uploader for user to upload their Excel file
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

    if uploaded_file is not None:
        xls = pd.ExcelFile(uploaded_file)
        shipments = pd.read_excel(xls, sheet_name="Shipments")
        stock = pd.read_excel(xls, sheet_name="Stock")
        master = pd.read_excel(xls, sheet_name="Master")

        result_df, updated_stock = process_shipments(shipments, stock, master)

        result_df = result_df.groupby(['Storage Area', 'Material', 'Material Descriptiion', 'S. Cat'])[['Replenishment Quantity (in Box)', 'Replenishment Quantity (in Pallet)']].sum().reset_index()

        if not result_df.empty:
            st.write("Replenishment Plan")
            st.dataframe(result_df)

            # Convert dataframe to Excel
            excel_data = convert_df_to_excel(result_df)
            
            # Provide a download button for the Excel file
            st.download_button(
                label="Download Replenishment Plan as Excel",
                data=excel_data,
                file_name="replenishment_plan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.write("Updated Stock")
            st.dataframe(updated_stock)

            stock_csv = updated_stock.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="Download Updated Stock as CSV",
                data=stock_csv,
                file_name='updated_stock.csv',
                mime='text/csv',
            )
        else:
            st.write("No replenishment needed.")

if __name__ == "__main__":
    main()
