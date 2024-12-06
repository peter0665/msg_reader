import gradio as gr
from msg_reader import extract_table_from_msg
from xlsx_combine import combine_all_excel
import os
from datetime import datetime as dt
from pathlib import Path


def merge_excel(files):
    file_list = []
    error_list = []
    out_file = f'{dt.now().strftime("%Y%m%d%H%M%S")}.xlsx'
    for idx, file in enumerate(files):
        abs_path = Path(str(file.name))
        try:
            if os.path.splitext(abs_path)[-1] != ".msg":
                raise Exception("ext")
            file_list.append(abs_path)
        except Exception:
            error_list.append([os.path.basename(abs_path), "Not msg file"])

    df, ext_error = extract_table_from_msg(file_list, out_file)

    if ext_error:
        error_list += ext_error

    if not error_list:
        error_list = [["None"]]

    if df is None:
        return None, None, error_list

    return out_file, df, error_list


def combine_excel(files):
    file_list = []
    error_list = []
    out_file = f'combine_xlsx_{dt.now().strftime("%Y%m%d%H%M%S")}.xlsx'
    # Get files
    for idx, file in enumerate(files):
        abs_path = Path(str(file.name))
        if os.path.splitext(abs_path)[-1].lower() not in [".xlsx", ".xls"]:
            error_list.append([os.path.basename(abs_path), "Not excel file"])
            continue
        file_list.append(abs_path)

    df, cae_errors = combine_all_excel(file_list, out_file)

    if cae_errors:
        error_list += cae_errors

    if not error_list:
        error_list = [["None"]]

    return out_file, error_list


msg_output = [
    gr.File(label="Output Excel"),
    gr.DataFrame(label="Result Table", interactive=True),
    gr.List(label="Error files", headers=["Name"], interactive=True),
]

xlsx_output = [
    gr.File(label="Output Excel"),
    # gr.DataFrame(label="Result Table", interactive=True),
    gr.List(label="Error files", headers=["Name"], interactive=True),
]

with gr.Blocks() as app:
    with gr.Tab(label="MSG Reader"):
        gr.Interface(
            merge_excel,
            gr.File(file_count="multiple", file_types=[".msg"]),
            msg_output,
            flagging_mode="never",
        )
    with gr.Tab(label="Excel combiner(Beta)"):
        gr.Interface(
            combine_excel,
            gr.File(file_count="multiple", file_types=[".xlsx", "xls"]),
            xlsx_output,
            flagging_mode="never",
        )

# demo = gr.Interface(
#     merge_excel,
#     gr.File(file_count="multiple", file_types=["msg"]),
#     demo_output,
#     allow_flagging="never",
# )

app.launch(server_port=7990, server_name="0.0.0.0")
