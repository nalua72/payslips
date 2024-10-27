import os
import re
import pandas as pd
from tabula import read_pdf


def main():
    data = []
    # thisdir = r"Z:\Jose\Nominas\Cognizant"
    thisdir = os.getcwd()

    for file in os.listdir(thisdir):
        # For every payslip file, extracts the main table, adds a column with the date and stores it in a list
        if "pdf" in file:
            if file.split("_")[0] != "20220430":
                dfs1 = read_pdf(file, stream=True, pages=1, relative_area=True, relative_columns=True, area=[
                                38, 0, 63, 100], columns=[15, 25, 31, 69, 80])
                df1 = pd.DataFrame(dfs1[0])
                df1.insert(0, "FECHA", [file.split("_")[0]
                                        for _ in range(len(df1))])
                data.append(df1)

                # For march months (except for 2016 and 2017) there are 2 pages in the excel because of the BONUS. Table is extracted from the second page,
                # added Fecha columnd and stored it in the list
                if re.search("[0-9][0-9][0-9][0-9]03[0-9][0-9]", file) and file.split("_")[0] != "20160331" and file.split("_")[0] != "20170331":
                    dfs1 = read_pdf(file, stream=True, pages=2, relative_area=True, relative_columns=True, area=[
                                    38, 0, 63, 100], columns=[15, 25, 31, 69, 80])
                    df1 = pd.DataFrame(dfs1[0])
                    df1.insert(0, "FECHA", [file.split("_")[0]
                                            for _ in range(len(df1))])
                    data.append(df1)
            else:
                dfs1 = read_pdf(file, stream=True, pages=1, relative_area=True, relative_columns=True, area=[
                    36, 0, 63, 100], columns=[15, 25, 30, 67, 78])
                df1 = pd.DataFrame(dfs1[0])
                df1.insert(0, "FECHA", [file.split("_")[0]
                                        for _ in range(len(df1))])
                data.append(df1)

    # A DataFrame is created with all the dataframes stored in the list
    df = pd.concat(data)

    df.columns = ["FECHA", "CUANTIA", "PRECIO", "CODIGO",
                  "CONCEPTO", "DEVENGOS", "DEDUCCIONES"]
    df = df.reset_index(drop=True)
    df = df.fillna(0)

    df['CUANTIA'] = df['CUANTIA'].str.replace(',', '.')
    df['PRECIO'] = df['PRECIO'].str.replace(',', '.')
    df['DEVENGOS'] = df['DEVENGOS'].str.replace('.', '')
    df['DEVENGOS'] = df['DEVENGOS'].str.replace(',', '.')
    df['DEDUCCIONES'] = df['DEDUCCIONES'].str.replace(',', '.')

    column_modify = ["CUANTIA", "PRECIO", "DEVENGOS", "DEDUCCIONES"]
    df[column_modify] = df[column_modify].astype(float)
    df['FECHA'] = pd.to_datetime(df['FECHA'], format='%Y%m%d').dt.date
    df['CODIGO'] = df['CODIGO'].astype(int)
    df = df.fillna(0)

    df['NOMINA'] = df['DEVENGOS'] - df['DEDUCCIONES']
    # Drop empty rows

    filt = df['CONCEPTO'] == 0
    df = df.drop(index=df[filt].index)

    # DataFrame is stored in excel file
    df.to_excel('payslip.xlsx', index=False)


if __name__ == "__main__":
    main()
