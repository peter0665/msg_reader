import pandas as pd


def combine_all_excel(file_paths, out_file_path):
    error_list = []
    # Create an empty DataFrame
    merged_df = pd.DataFrame()

    # Loop through all files and merge them in order
    for i, file_path in enumerate(file_paths):
        try:
            df = pd.read_excel(file_path, "MP")
            # Make all date to MM/DD/YYYY format
            for col in df.columns:
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    df[col] = df[col].dt.strftime("%m/%d/%Y")
            if i == 0:
                # Use columns of the first file as column items
                merged_df = df
            else:

                # For other files, merge according to the column order of the first file
                merged_df = pd.concat([merged_df, df], ignore_index=True)
        except Exception as e:
            error_list.append([file_path, str(e)])

    # Save the merged DataFrame to an Excel file
    merged_df.to_excel(out_file_path, index=False, engine="openpyxl")

    return merged_df, error_list
