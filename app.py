import streamlit as st
import pandas as pd
from scoring_script import process_scores
import os

st.set_page_config(page_title="学生体测评分系统", layout="wide")
st.title("🏃‍♂️ 学生体测评分系统")

uploaded_file = st.file_uploader("请上传原始 Excel 文件（.xlsx）", type=["xlsx"])

if uploaded_file is not None:
    # 保存上传的文件
    with open("raw_scores.xlsx", "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.success("✅ 文件上传成功，正在评分中...")

    # 调用评分函数并接收返回值
    try:
        total_file = process_scores("raw_scores.xlsx")
    except Exception as e:
        st.error(f"❌ 评分过程中发生错误：{e}")
        st.stop()

    if total_file is None or not os.path.exists(total_file):
        st.error("❌ 没有找到评分结果文件，请确认表格内容是否符合要求。")
        st.stop()

    # 读取评分结果
    result_df = pd.read_excel(total_file)

    st.subheader("📊 总表评分结果预览（前 30 行）")
    st.dataframe(result_df.head(30), use_container_width=True)

    # 下载总表按钮
    with open(total_file, "rb") as f:
        st.download_button(
            label="⬇️ 下载总评分结果 Excel 文件",
            data=f,
            file_name=total_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # 分班下载按钮
    st.subheader("📁 分班评分结果下载")

    class_files = [
        f for f in os.listdir()
        if f.endswith(".xlsx") and f.startswith("_") is False and "总表" not in f
    ]

    if class_files:
        for file in sorted(class_files):
            with open(file, "rb") as f:
                st.download_button(
                    label=f"⬇️ 下载：{file}",
                    data=f,
                    file_name=file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.info("暂无分班文件，请确认评分已完成并包含班级字段。")

