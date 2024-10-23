import pandas as pd
import os
import sys

def process_csv_data(file_path, start_row, end_row):
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"ファイルが見つかりません: {file_path}")
            
        df = pd.read_csv(file_path, encoding='utf-8')
        
        if df.empty:
            raise ValueError("CSVファイルにデータが含まれていません")
            
        result_lines = []
        
        header_lines = [
            "田中様，大八木様",
            "cc：藤澤先生",
            "",
            "情報システム工学科 藤澤研究室 ４年",
            "",
            "渡辺 太貴です．",
            "",
            "以下の物品の手配をお願いいたします．",
            ""
        ]
        result_lines.extend(header_lines)
        
        for row in range(start_row, end_row + 1):
            if row >= len(df):
                break
                
            try:
                budget = df.iloc[row, 0]
                item_name = df.iloc[row, 1]
                model_number = df.iloc[row, 2]
                vendor = df.iloc[row, 3]
                quantity = int(df.iloc[row, 4]) if pd.notna(df.iloc[row, 4]) else 0
                unit_price = df.iloc[row, 5]
                subtotal = df.iloc[row, 6]
                note = df.iloc[row, 7] if 7 < len(df.columns) else None
                
                line_data = [
                    f"支出予算: {budget}",
                    f"品名: {item_name}",
                    f"販売元: {vendor}",
                    f"数量: {quantity}",
                    f"単価: {unit_price}",
                    f"小計: {subtotal}"
                ]
                
                if pd.notna(model_number):
                    line_data.insert(2, f"型番: {model_number}")
                
                if pd.notna(note):
                    line_data.append(str(note))
                
                result_lines.extend(line_data)
                result_lines.append("")
                
            except Exception as e:
                print(f"行 {row + 1} の処理中にエラーが発生しました: {str(e)}")
                continue
        
        footer_lines = [
            "-------------------------------------------------",
            "北九州市立大学　国際環境工学部 情報システム工学科 藤澤研究室",
            "渡辺 太貴",
            "E-mail: c1531069@eng.kitakyu-u.ac.jp",
            "-------------------------------------------------"
        ]
        result_lines.extend(footer_lines)
                
        return result_lines
        
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        sys.exit(1)

def main():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(current_dir, "data.csv")
    start_row = 2
    end_row = 3
    
    try:
        results = process_csv_data(file_path, start_row - 2, end_row - 2)
        
        for line in results:
            print(line)
            
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()