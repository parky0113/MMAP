import pandas as pd
from datetime import datetime

# Function to highlight rows based on conditions
def highlight_row(row):
    # Convert "Invoice Date" to datetime and calculate the difference in days
    inv_date = pd.to_datetime(row.loc["Invoice Date"], dayfirst=True)
    diff = (pd.to_datetime(datetime.now().strftime("%d/%m/%Y")) - inv_date).days
    
    # Set background color based on conditions
    if row.loc["IsCreditMemo"] == True:
        color = "#ff91a4"  # Red color for credit memos
    elif diff > 10:
        color = "#ffffcc"  # Light yellow color for older invoices
    else:
        color = "#FFFFFF"  # White color for other rows
    return [f'background-color: {color}' for r in row]

# Function to convert Excel data to dictionaries
def excel_to_dict(supp_df):
    supp_list = []
    discriminant_list = []
    special_list = []

    for col in supp_df.columns:
        key = col
        values = list(supp_df[col].dropna())
        discriminant_list.append(values)
        supp_list.append(key)
        if len(values[0]) > 1:
            special_list.append(values[0])

    return supp_list, discriminant_list, special_list

# Function to export data to Excel pages
def export_pages(sheet, page, entity, ind):
    with pd.ExcelWriter(f"reports/{datetime.now().strftime('%d-%m-%Y')} {page} {entity} - {ind}.xlsx") as writer:
        sheet.to_excel(writer, sheet_name=page, index=False)
        wb = writer.book
        for column in sheet:
            column_width = max(sheet[column].astype(str).map(len).max(), len(column))
            col_idx = sheet.columns.get_loc(column)
            writer.sheets[page].set_column(col_idx, col_idx, column_width + 1)
        format = wb.add_format({'text_wrap': True})
        writer.sheets[page].set_column(sheet.columns.get_loc("Comments"), sheet.columns.get_loc("Comments"), 35, format)
        sheet = sheet.style.apply(highlight_row, axis=1)
        sheet.to_excel(writer, sheet_name=page, index=False)

# Main function
def main(data_df, supp_df):

    # Get Supplier dictionary
    supp_list, discriminant_list, special_list = excel_to_dict(supp_df)

    # Clean the Master Data
    data_df.drop(columns=["SC_Invoice_UniqueId", "Invoice Type"], inplace=True)
    data_df = data_df.loc[(data_df['Status'] == "Pending") | (data_df['Status'] == "Approved")]
    data_df["Entity"].fillna("BLANK", inplace=True)

    # Save it for later reference
    column_list = data_df.columns

    # Get Entity and Dictionary for each page
    entity_list = data_df['Entity'].unique()
    entity_dict = {entity: {supp: [] for supp in supp_list} for entity in entity_list}

    # Sorting every row of the data into appropriate pages
    for row in data_df.itertuples():
        supp = row[3]
        entity = row[1]

        if supp in special_list:
            location = discriminant_list.index([supp])
            entity_dict[entity][supp_list[location]].append(row)
        else:
            ind = 0
            skip = 0
            while skip == 0:
                if supp[0].upper() in discriminant_list[ind]:
                    entity_dict[entity][supp_list[ind]].append(row)
                    skip = 1
                ind += 1

    # Export each entity/supplier
    for entity in entity_list:
        if len(entity_dict[entity]) != 0:
            for page in supp_list:
                if len(entity_dict[entity][page]) > 1:
                    ind = 1
                    sheet = pd.DataFrame(entity_dict[entity][page])
                    sheet.drop(columns=sheet.columns[0], axis=1, inplace=True)
                    sheet.columns = column_list
                    sheet.sort_values(by=["Supplier Name", "Invoice Date", "PO #"], inplace=True)
                    sheet["Invoice Date"] = sheet["Invoice Date"].dt.strftime("%d/%m/%Y")
                    sheet["ReceivedDate"] = sheet["ReceivedDate"].dt.strftime("%d/%m/%Y")
                    if len(sheet) < 50:
                        export_pages(sheet, page, entity, ind)
                    else:
                        while len(sheet) >= 50:
                            sheet1 = sheet.iloc[:50]
                            export_pages(sheet1, page, entity, ind)
                            sheet = sheet.iloc[50:]
                            ind += 1
                        if len(sheet) != 0:
                            export_pages(sheet, page, entity, ind)