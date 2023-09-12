import streamlit as st
import pandas as pd
from docx import Document
from collections import Counter
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO
import base64
from janome.tokenizer import Tokenizer

st.title("Word文書解析")
uploaded_file = st.file_uploader("ファイルをアップロードしてください。", type=["docx"])

# ファイルがアップロードされた場合
if uploaded_file is not None:
    document = Document(uploaded_file)
    text = ""
    for para in document.paragraphs:
        text += para.text

    # テキストの単語分割
    tokenizer = Tokenizer()
    tokens = tokenizer.tokenize(text)

    # 名詞、動詞、形容詞の品詞を抽出し、出現回数をカウント
    selected_pos = ["名詞", "動詞", "形容詞"]
    counter = Counter()
    for token in tokens:
        pos = token.part_of_speech.split(',')[0]
        if pos in selected_pos:
            counter[token.surface] += 1

    # データフレームに変換
    data = {"単語/品詞": list(counter.keys()), "出現回数": list(counter.values())}
    df = pd.DataFrame(data)

    # 円グラフの表示
    import matplotlib.pyplot as plt
    import japanize_matplotlib
    from matplotlib.font_manager import FontProperties
    # 日本語フォントがイマイチ
    plt.rcParams['font.family'] = 'IPAexGothic'
    fig, ax = plt.subplots(figsize=(20, 20))
    ax.pie(df["出現回数"], labels=df["単語/品詞"], startangle=90, counterclock=False, autopct="%1.1f%%")
    ax.set_title("単語/品詞出現回数")
    st.pyplot(fig)

    from wordcloud import WordCloud
    wordcloud = WordCloud(width=800, height=400, background_color='white', font_path='./ipaexg.ttf').generate(text)
    # ワードクラウドをMatplotlibのプロットとして表示
    # st.pyplot(plt.figure(figsize=(10, 5)))
    plt.figure(figsize=(10, 5))
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis("off")
    st.pyplot(plt)

    # データフレームの表示
    st.subheader("単語/品詞出現回数一覧")
    st.dataframe(df.sort_values("出現回数", ascending=False))

    df_download = df.sort_values("出現回数", ascending=False)

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
