import pandas as pd
import openpyxl
import locale

def format_currency(amount):
    return f"¥{amount:,}"

def process_excel_data(file_path, start_row, end_row):
    df = pd.read_excel(file_path)
    workbook = openpyxl.load_workbook(file_path, data_only=True)
    sheet = workbook.active
    result_lines = []
    
    header_lines = [
        "田中様，大八木様",
        "cc：藤澤先生",
        "",
        "情報システム工学科 藤澤研究室 ４年",
        "NAMEです．",
        "",
        "以下の物品の手配をお願いいたします．",
        ""
    ]
    result_lines.extend(header_lines)
    
    for row in range(start_row, end_row + 1):
        if row >= len(df):
            break
            
        budget = sheet.cell(row=row+1, column=7).value
        item_name = df.iloc[row, 7]
        model_number = df.iloc[row, 8]
        vendor = df.iloc[row, 9]
        quantity = int(df.iloc[row, 10]) if pd.notna(df.iloc[row, 10]) else 0
        unit_price = int(df.iloc[row, 11]) if pd.notna(df.iloc[row, 11]) else 0
        subtotal = int(df.iloc[row, 12]) if pd.notna(df.iloc[row, 12]) else 0
        note = df.iloc[row, 13]
        
        line_data = [
            f"支出予算: {budget}",
            f"品名: {item_name}",
            f"販売元: {vendor}",
            f"数量: {quantity}",
            f"単価: {format_currency(unit_price)}",
            f"小計: {format_currency(subtotal)}"
        ]
        
        if pd.notna(model_number):
            line_data.insert(2, f"型番: {model_number}")
        
        if pd.notna(note):
            line_data.append(str(note))
        
        result_lines.extend(line_data)
        result_lines.append("")
    
    footer_lines = [
        "以上よろしくお願いたします.",
        "",
        "---------------------------------------------------",
        "北九州市立大学　国際環境工学部 情報システム工学科 藤澤研究室",
        "NAME",
        "E-mail: EMAIL",
        "---------------------------------------------------"
    ]
    result_lines.extend(footer_lines)
            
    workbook.close()
    return result_lines

def main():
    file_path = "temp.xlsx"
    start_row = 23
    end_row = 23
    
    results = process_excel_data(file_path, start_row, end_row)

    for line in results:
        print(line)

if __name__ == "__main__":
    main()