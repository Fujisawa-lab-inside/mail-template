import pandas as pd
import os
import sys

def process_csv_data(file_path, name, email, grade):
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"ファイルが見つかりません: {file_path}")
            
        df = pd.read_csv(file_path, encoding='utf-8')
        
        if df.empty:
            raise ValueError("CSVファイルにデータが含まれていません")
            
        result_lines = [
            "田中様，大八木様",
            "cc：藤澤先生",
            "",
            f"情報システム工学科 藤澤研究室 {grade}",
            "",
            f"{name}です．",
            "",
            "以下の物品の手配をお願いいたします．",
            ""
        ]

        for index, row in df.iterrows():
                
            try:
                line_data = [
                    f"支出予算: {row['支出予算']}",
                    f"品名: {row['商品名']}",
                    f"販売元: {row['販売店']}",
                    f"数量: {int(float(str(row['数量']).replace(',', '')))}",
                    f"単価: {row['単価']}",
                    f"小計: {row['小計']}"
                ]
                
                if pd.notna(row['型番']):
                    line_data.insert(2, f"型番: {row['型番']}")
                
                if pd.notna(row['URL']):
                    line_data.append(str(row['URL']))
                
                result_lines.extend(line_data)
                result_lines.append("")
                
            except Exception as e:
                print(f"行 {index + 1} の処理中にエラーが発生しました: {str(e)}")
                continue

        footer_lines = [
            "-------------------------------------------------",
            "北九州市立大学　国際環境工学部 情報システム工学科 藤澤研究室",
            f"{name}",
            f"E-mail: {email}",
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
    name = "渡辺 太貴"
    email = "c1531069@eng.kitakyu-u.ac.jp"
    grade = "B4"
    
    try:
        results = process_csv_data(file_path, name, email, grade)
        for line in results:
            print(line)
            
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()