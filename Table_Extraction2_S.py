import streamlit as st
import camelot
import numpy as np
import pandas as pd
import base64
import io

def create_download_link(data, filename):
    b64 = base64.b64encode(data).decode()  # 將檔案數據轉換為 base64 編碼
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}" download="{filename}">點此下載</a>'
    return href

def get_text_after_keyword(input_string, keyword):
    index = input_string.find(keyword)
    
    if index != -1:
        return input_string[index + len(keyword):]
    else:
        return ""
    
def swap_columns(row):
    if row['里村'] is None:
        row['區鄉'], row['里村'] = row['里村'], row['區鄉']
    return row

def extract_tables_from_pdf(pdf_file, output_excel):
    #with open("input.pdf", "wb") as f:
    #    base64_pdf = base64.b64encode(pdf_file.getbuffer()).decode()
    #    f.write(base64.b64decode(base64_pdf))
    #f.close()

    # Use BytesIO to create a file-like object in memory
    buffer = io.BytesIO()
    buffer.write(pdf_file.read())

    # Now you can manipulate the file data in memory
    # For example, you can save it to a file on your local machine
    with open("input.pdf", "wb") as f:
        f.write(buffer.getbuffer())

    # 使用camelot-py從pdf中讀取表格數據
    tables = camelot.read_pdf("input.pdf", flavor='stream', pages='all')

    # 合併所有表格到一個DataFrame中
    all_tables_data = pd.DataFrame()
    
    for table in tables:
        df = table.df
        
        extracted_city = df.iloc[1][0].replace(" ", "")
        
        if len(extracted_city) > 0:
            extracted_city = get_text_after_keyword(extracted_city, "縣市別：")
            field_name = list(df.iloc[2])
            field_name[0] = '區里'
            df = df.iloc[3:]
            df.columns = field_name
        else:
            field_name = list(df.iloc[1])
            field_name[0] = '區里'
            df = df.iloc[2:]
            df.columns = field_name
        
        all_tables_data = pd.concat([all_tables_data, df], ignore_index=True, axis=0)

    # 將Column1根據'\n'分隔後轉為兩個欄位
    split_data = all_tables_data['區里'].str.split('\n', expand=True)

    # 複製整個Data Frame
    df_copy = pd.DataFrame()
    df_copy['區鄉'] = split_data[0]
    df_copy['里村'] = split_data[1]
    
    all_tables_data = all_tables_data.drop('區里', axis=1)
    
    # 將兩個Data Frame合併
    result_df = pd.concat([df_copy, all_tables_data], ignore_index=True, axis=1)
    result_df.columns = df_copy.columns.append(all_tables_data.columns)
    
    result_df = result_df.apply(swap_columns, axis=1)
    result_df.iloc[:, 0] = result_df.iloc[:, 0].fillna(method='ffill', axis=0)

    # 創建一個新的Excel文件
    writer = pd.ExcelWriter(output_excel+".xlsx", engine='xlsxwriter')

    # 將合併的表格存儲到Excel文件的Data Sheet中
    result_df.to_excel(writer, header=True, sheet_name=extracted_city, index=False)
    result_df.to_csv(output_excel+".csv", index=False, header=True)

    writer.close()
    
    print(f"All tables extracted from {pdf_file} and saved to {output_excel}")

if __name__ == '__main__':
    st.title('資料轉換系統(pdf->xlsx)')

    file = st.file_uploader('請選擇要上傳的文件:', type=['pdf'])
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    if col1.button('轉換檔案') and file is not None:
        extract_tables_from_pdf(file, "converted")
        
        #with open('converted.xlsx', 'rb') as f:
        #    data = f.read()

        #顯示下載連結
        #st.markdown(create_download_link(data, 'converted.xlsx'), unsafe_allow_html=True)
        
        with open('converted.xlsx', 'rb') as my_file:
            col2.download_button(label = '📥點此下載', data = my_file, file_name = 'converted.xlsx')    