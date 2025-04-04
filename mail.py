import pandas as pd
import streamlit as st
from io import StringIO
import json


def process_csv_data(csv_text, name, email, grade):
    try:
        csv_file_like = StringIO("支出予算,商品名,型番,販売店,数量,単価,小計,URL\n" + csv_text)
        df = pd.read_csv(csv_file_like)

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


def load_data():
    try:
        with open("data.json", "r") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}

def save_data(data):
    with open("data.json", "w") as f:
        json.dump(data, f)
        
def main():
    st.set_page_config(page_title="mail-generator", layout="centered")
    st.title(" メール生成")
    
    data = load_data()

    if "name" not in st.session_state:
        st.session_state.name = data.get("name", "")
    if "email" not in st.session_state:
        st.session_state.email = data.get("email", "")
    if "grade" not in st.session_state:
        st.session_state.grade = data.get("grade", "")

    name = st.text_input("名前", value=st.session_state.name)
    email = st.text_input("メールアドレス", value=st.session_state.email)
    grade_suggestions = ["B1", "B2", "B3", "B4", "M1", "M2", "D1", "D2", "D3"]
    grade = st.selectbox("学年", grade_suggestions, index=grade_suggestions.index(st.session_state.grade) if st.session_state.grade in grade_suggestions else 0)

    st.session_state.name = name
    st.session_state.email = email
    st.session_state.grade = grade

    save_data({"name": name, "email": email, "grade": grade})
    
    csv_text_input = st.text_area("ここにCSVデータを貼り付けてください", height=200)

    dataframe_manual = None
    if csv_text_input:
        try:
            csv_file_like = StringIO("支出予算,商品名,型番,販売店,数量,単価,小計,URL\n" + csv_text_input)
            dataframe_manual = pd.read_csv(csv_file_like)
            st.success("テキストデータの読み込みに成功しました．")
            st.dataframe(dataframe_manual)
        except Exception as e:
            st.error(f"CSVデータの解析中にエラーが発生しました: {str(e)}")

    if st.button("✅ 決定"):
        if name and email and grade and csv_text_input:
            try:
                results = process_csv_data(csv_text_input, name, email, grade)                    
                if st.button("コピー"):
                        pyperclip.copy("\n".join(results))
                        st.success("クリップボードにコピーしました．")
                st.text_area("生成されたメール", value="\n".join(results), height=400)
            except Exception as e:
                print(f"エラーが発生しました: {str(e)}")
                sys.exit(1)
        else:
            if not name:
                st.warning("名前を入力してください")
            if not email:
                st.warning("メールアドレスを入力してください")
            if not grade:
                st.warning("学年を入力してください")
            if not csv_text_input:
                st.warning("CSVデータを貼り付けてください")

if __name__ == "__main__":
    main()