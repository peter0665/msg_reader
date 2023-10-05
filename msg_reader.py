import extract_msg
import pandas as pd
import os
import re
import sys

def extract_table_from_msg(msg_files, excel_file_path):
    df = None
    error_list = []
    for msg_path in msg_files:
        # Load the .msg file
        msg = extract_msg.Message(msg_path)
        
        # Extract body from the .msg file
        body = msg.htmlBody.decode()
        
        # Split the body into sections based on empty lines to find the table
        table = re.search(r"<table.*?</table>", body)
        if not table:
            error_list.append([os.path.basename(msg_path), "table is not found"])
            continue
        html_data = pd.read_html(table[0], header=0)
        if not html_data:
            error_list.append([os.path.basename(msg_path), "read table fail"])
            continue
        if df is None:
            df = html_data[0]
        else:
            df = pd.concat([df, html_data[0]], axis=0)
        # print("OK")
    # Reset index
    # df = df.reset_index()
    # Drop index
    # df = df.drop("index", errors="ignore")
    
    # # Save the DataFrame to an Excel file
    if df is None:
        return None

    df.to_excel(excel_file_path, index=False, engine="openpyxl")

    return df, error_list
