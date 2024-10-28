import os
import re
import pandas as pd
from pathlib import Path
from tabula import read_pdf

# Function to extract a table from a file a return a dataframe. A DATA column is added.


def pdftable_to_dataframe(file, area, columns, page):
    dfs = read_pdf(file, stream=True, pages=page, relative_area=True,
                   relative_columns=True, area=area, columns=columns)
    df = pd.DataFrame(dfs[0])
    df.insert(0, "FECHA", [file.name.split("_")[0] for _ in range(len(df))])
    return df


def main():
    data = []
    # thisdir = r"Z:\Jose\Nominas\Cognizant"
    # thisdir = os.getcwd()

    base_path = Path(r"Z:\Jose\Nominas\Cognizant")
    files = list(base_path.glob("*.pdf"))

    for file in files:

        if file.name.split("_")[0] != "20220430":
            area = [38, 0, 63, 100]
            columns = [15, 25, 31, 69, 80]

            data.append(pdftable_to_dataframe(file, area, columns, 1))

            # For march months (except for 2016 and 2017) there are 2 pages in the excel because of the BONUS. Table is extracted from the second page.
            if re.search("[0-9][0-9][0-9][0-9]03[0-9][0-9]", file.name) and file.name.split("_")[0] != "20160331" and file.name.split("_")[0] != "20170331":

                data.append(pdftable_to_dataframe(file, area, columns, 2))
        else:
            area = [36, 0, 63, 100]
            columns = [15, 25, 30, 67, 78]

            data.append(data.append(
                pdftable_to_dataframe(file, area, columns, 1)))

    # for file in os.listdir(thisdir):
    #     # For every payslip file, a dataframe with the table is stored in a list
    #     if "pdf" in file:
    #         if file.split("_")[0] != "20220430":
    #             area = [38, 0, 63, 100]
    #             columns = [15, 25, 31, 69, 80]

    #             data.append(pdftable_to_dataframe(file, area, columns, 1))

    #             # For march months (except for 2016 and 2017) there are 2 pages in the excel because of the BONUS. Table is extracted from the second page.
    #             if re.search("[0-9][0-9][0-9][0-9]03[0-9][0-9]", file) and file.split("_")[0] != "20160331" and file.split("_")[0] != "20170331":

    #                 data.append(pdftable_to_dataframe(file, area, columns, 2))
    #         else:
    #             area = [36, 0, 63, 100]
    #             columns = [15, 25, 30, 67, 78]

    #             data.append(data.append(
    #                 pdftable_to_dataframe(file, area, columns, 1)))

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
