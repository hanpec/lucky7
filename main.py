import streamlit as st
import pandas as pd
# import matplotlib.pyplot as plt
from docx import Document
from collections import Counter
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO
import base64


# タイトルの表示
st.title("Word文書解析アプリ")

# ファイルのアップロード
uploaded_file = st.file_uploader("ファイルをアップロードしてください。", type=["docx"])

# ファイルがアップロードされた場合
if uploaded_file is not None:
    document = Document(uploaded_file)
    text = ""
    for para in document.paragraphs:
        text += para.text

    # テキストの単語分割
    words = text.split()

    # 単語と品詞のカウント
    counter = Counter()
    for word in words:
        counter[word] += 1

    # データフレームに変換
    data = {"単語/品詞": list(counter.keys()), "出現回数": list(counter.values())}
    df = pd.DataFrame(data)

    # 円グラフの表示
    # fig, ax = plt.subplots()
    # ax.pie(df["出現回数"], labels=df["単語/品詞"], startangle=90, counterclock=False, autopct="%1.1f%%")
    # ax.set_title("単語/品詞出現回数")
    # st.pyplot(fig)

    # データフレームの表示
    st.subheader("単語/品詞出現回数一覧")
    st.dataframe(df.sort_values("出現回数", ascending=False))

    # # Excelダウンロード用のデータフレームの作成
    # df_download = pd.DataFrame(columns=["単語/品詞", "出現回数"])
    # for row in dataframe_to_rows(df):
    #     df_download = df_download.append({"単語/品詞": row[0], "出現回数": row[1]}, ignore_index=True)

    # Excelダウンロード用のデータフレームの作成
    # download_data = []
    # for row in dataframe_to_rows(df):
    #     download_data.append({"単語/品詞": row[1], "出現回数": row[2]})

    # df_download = pd.DataFrame(download_data, columns=["単語/品詞", "出現回数"])
    # for row in dataframe_to_rows(df, index=False, header=True):
    #     sheet.append(row)


    df_download = df.sort_values("出現回数", ascending=False)

    # # ダウンロードボタンの表示
    # if st.button("Excelにダウンロード"):
    #     output = BytesIO()
    #     writer = pd.ExcelWriter(output, engine="xlsxwriter")
    #     df_download.to_excel(writer, sheet_name="Sheet1", index=False)
    #     writer.save()
    #     processed_data = output.getvalue()
    #     output.seek(0)
    #     b64 = base64.b64encode(processed_data).decode()
    #     href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="word_analysis.xlsx">Excelファイルのダウンロード</a>'
    #     st.markdown(href, unsafe_allow_html=True)
    # else:
    #     st.warning("ファイルがアップロードされていません。")

    # ダウンロードボタンの表示
    if st.button("Excelにダウンロード"):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_download.to_excel(writer, sheet_name="Sheet1", index=False)
        processed_data = output.getvalue()
        output.seek(0)
        b64 = base64.b64encode(processed_data).decode()
        href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="word_analysis.xlsx">Excelファイルのダウンロード</a>'
        st.markdown(href, unsafe_allow_html=True)
    # else:
    #     st.warning("ファイルがアップロードされていません。")
