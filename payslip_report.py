import pandas as pd
from pathlib import Path
from tabula import read_pdf
import re

# Function to extract a table from a file a return a dataframe. A DATA column is added.


def pdftable_to_dataframe(file, area, columns, page):
    dfs = read_pdf(file, stream=True, pages=page, relative_area=True,
                   relative_columns=True, area=area, columns=columns)
    df = pd.DataFrame(dfs[0])
    df.insert(0, "FECHA", [file.name.split("_")[0] for _ in range(len(df))])
    return df

# Function that takes a df with payslips and formats it


def dataframe_format(df):
    # Name the columns
    df.columns = ["FECHA", "CUANTIA", "PRECIO",
                  "CODIGO", "CONCEPTO", "DEVENGOS", "DEDUCCIONES"]
    df = df.reset_index(drop=True)
    df = df.fillna(0)

    # Cambia el formato de strin a float en las columnas ("CUANTIA", "PRECIO", "DEVENGOS", "DEDUCCIONES"). Previamente se cambia el caracter ',' por '.'
    df['DEVENGOS'] = df['DEVENGOS'].str.replace('.', '')
    column_to_modify = ["CUANTIA", "PRECIO", "DEVENGOS", "DEDUCCIONES"]
    for col in column_to_modify:
        df[col] = df[col].str.replace(',', '.', regex=False).astype(float)

    # Formatea el campo'FECHA' a la forma deseada
    df['FECHA'] = pd.to_datetime(df['FECHA'], format='%Y%m%d').dt.date

    # Formatea el campo'CODIGO' a int
    df = df.fillna(0)

    # Adds column Balance
    df['BALANCE'] = df['DEVENGOS'] - df['DEDUCCIONES']

    # Drop empty rows
    filt = df['CONCEPTO'] == 0
    df = df.drop(index=df[filt].index)
    return df


def main():
    data = []

    payslips_path = Path(r"Z:\Jose\Nominas\Cognizant")
    dest_file = Path(r"Z:\Jose\Nominas\Cognizant\payslip_report.xlsx")

    payslips = list(payslips_path.glob("*.pdf"))

    for payslip in payslips:
        if payslip.name.split("_")[0] != "20220430":
            area = [38, 0, 63, 100]
            columns = [15, 25, 31, 69, 80]

            data.append(pdftable_to_dataframe(payslip, area, columns, 1))

            # For march months (except for 2016 and 2017) there are 2 pages in the excel because of the BONUS. Table is extracted from the second page.
            if re.search("\d{4}03\d{2}", payslip.name) and payslip.name.split("_")[0] != "20160331" and payslip.name.split("_")[0] != "20170331":
                data.append(pdftable_to_dataframe(payslip, area, columns, 2))
        else:
            area = [36, 0, 63, 100]
            columns = [15, 25, 30, 67, 78]

            data.append(data.append(
                pdftable_to_dataframe(payslip, area, columns, 1)))

    # A DataFrame is created with all the dataframes stored in the list
    df = pd.concat(data)

    # Dataframe cleaning
    df = dataframe_format(df)

    # DataFrame is stored in excel file
    df.to_excel(dest_file, sheet_name='Master', index=False)


if __name__ == "__main__":
    main()
