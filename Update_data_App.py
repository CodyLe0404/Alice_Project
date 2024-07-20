import pandas as pd
import time
import streamlit as st
import os

st.header("Auto update data in STD Calculation file")

upload_file = st.file_uploader("Upload file raw", type=['.xlsx'])

outputFile_name = st.text_input("Enter output file name")

export_output = st.button("EXPORT")

if upload_file is not None:
    try:
        df = pd.read_excel(upload_file)
        if export_output:
            st.write("Please wait file generating...")
            start_time = time.time()
            df_columns = df[['Sub Package Group', 'Material', 'Material Description', 'Operation Longer Name', 'Formula Key', 'Standard text key', 'New Machine Time']]
            df_columns['Formulas'] = ''
            df_columns['Manual Calculation'] = ''
            df_columns['Factor'] = ''
            df_columns['STD (UPH not xout)'] = ''
            df_columns['SAP UPH'] = ''
            df_columns['excel UPH']= ''
            df_columns['Var(%)'] = ''
            df_columns['Remark'] = ''

            UPH = []
            for index, row in df_columns.iterrows():
                #Check UPH number in Monaco
                if "2277" in row['Material Description']:
                    df_monaco = pd.read_excel("./Monaco.xlsx")
                    for _, row_monaco in df_monaco.iterrows():
                        row_standard = row['Standard text key']
                        row_monaco_standard = row_monaco['Standard text key']
                        UPH_monaco = row_monaco['UPH']
                        if row['Standard text key'] == row_monaco['Standard text key']:
                            UPH.append(row_monaco['UPH'])
                    if len(UPH) == 2:
                        if "interposer" in row['Material Description']:
                            df_columns.loc[index, 'excel UPH'] = UPH[1]
                            UPH1 = UPH[1]
                            UPH.clear()
                        else: 
                            df_columns.loc[index, 'excel UPH'] = UPH[0]
                            UPH0 = UPH[0]
                            UPH.clear()
                    elif len(UPH) == 1:
                        df_columns.loc[index, 'excel UPH'] = UPH[0]
                        UPH0 = UPH[0]
                        UPH.clear()
                    elif len(UPH) == 0:
                        df_columns.loc[index, 'excel UPH'] = ""
                    row_UPH = row['excel UPH']
                #Check UPH number in Qorvo
                elif "948" in row['Material Description']:
                    df_qorvo = pd.read_excel("./Qorvo.xlsx")
                    for _, row_qorvo in df_qorvo.iterrows():
                        if row['Standard text key'] == row_qorvo['Standard text key']:
                            if "76300" in row['Material Description']:
                                df_columns.loc[index, 'excel UPH'] = row_qorvo['QM76300']
                            elif "76065" in row['Material Description']:
                                df_columns.loc[index, 'excel UPH'] = row_qorvo['QM76065']
                            elif "76092" in row['Material Description']:
                                df_columns.loc[index, 'excel UPH'] = row_qorvo['QM76092']
                            elif "76308" in row['Material Description']:
                                df_columns.loc[index, 'excel UPH'] = row_qorvo['QM76308']
                            elif "76309" in row['Material Description']:
                                df_columns.loc[index, 'excel UPH'] = row_qorvo['QM76309']
                            elif "76095" in row['Material Description']:
                                df_columns.loc[index, 'excel UPH'] = row_qorvo['QM76095']
                df_choose = pd.read_excel('./Choose.xlsx')
                standard_key_not_found = []
                factor_list = []
                for _, row_choose in df_choose.iterrows():
                    row_standard = row['Standard text key']
                    row_choose_standard = row_choose['Standard text key']
                    if row['Standard text key'] == row_choose['Standard text key']:
                        factor_list.append(row_choose['Factor'])
                    if len(factor_list) != 0:
                        df_columns.loc[index, 'Factor'] = factor_list[0]
                    else:
                        df_columns.loc[index, 'Factor'] = ''
                    row_factor = row['Factor']
                if len(factor_list) == 0:
                        df_columns.loc[index, 'Factor'] = ""
                if df_columns.loc[index,'Factor'] != "":   
                    df_columns.loc[index, 'STD (UPH not xout)'] = df_columns.loc[index,'New Machine Time'] / df_columns.loc[index,'Factor']
                else: df_columns.loc[index, 'STD (UPH not xout)'] = ""
                if df_columns.loc[index, 'STD (UPH not xout)'] != "":
                    df_columns.loc[index, 'SAP UPH'] = 3600 / df_columns.loc[index, 'STD (UPH not xout)']
                else: df_columns.loc[index, 'SAP UPH'] = ""
                excel_UPH = df_columns.loc[index, 'excel UPH']
                var_SAP_UPH = df_columns.loc[index, 'SAP UPH'] 
                if df_columns.loc[index, 'excel UPH'] != '':
                    if df_columns.loc[index, 'SAP UPH'] != "":
                        df_columns.loc[index, 'Var(%)'] = (df_columns.loc[index, 'excel UPH'] - df_columns.loc[index, 'SAP UPH']) / df_columns.loc[index, 'excel UPH']
                    var_percent = df_columns.loc[index, 'Var(%)']
            if outputFile_name:
                outputFile_name = outputFile_name + ".xlsx"
                df_columns.to_excel(outputFile_name, index=False)
                st.write("Export output successfully. Please check output file!\n", os.getcwd())
                print("\nPlease output file!\n", os.getcwd())
                end_time = time.time()
                execution_time = end_time - start_time
                st.write("\nExecution time: ", execution_time, "seconds")
    except FileNotFoundError:
        st.write("File not found!")
    except pd.errors.ParserError:
        st.write("Error reading the Excel file.")
    except Exception as e:
        st.write("An error occurred:", e)



