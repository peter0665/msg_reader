import gradio as gr
from msg_reader import extract_table_from_msg
import os
from datetime import datetime as dt


def merge_excel(files):
    file_list = []
    error_list = []
    out_file = f'{dt.now().strftime("%Y%m%d%H%M%S")}.xlsx'
    for idx, file in enumerate(files):
        try:
            abs_path = file.orig_name
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

demo_output = [
    gr.File(label="Output Excel"),
    gr.DataFrame(label="Result Table", interactive=True),
    gr.List(label="Error files", interactive=True),
]

demo = gr.Interface(
    merge_excel,
    gr.File(file_count="multiple", file_types=["msg"]),
    demo_output,
    allow_flagging="never",
)
    
demo.launch(server_port=7990, server_name="0.0.0.0") 