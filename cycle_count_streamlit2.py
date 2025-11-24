# cycle_count_streamlit_fixed.py
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import datetime
import math
import os
from io import BytesIO

# PDF libs (with Chinese font registration)
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase import pdfmetrics

# Frontend QR/barcode scanner (works in Streamlit Cloud, mobile camera)
from streamlit_qrcode_scanner import qrcode_scanner
from PIL import Image

# ---------------- Helper functions ----------------

# Register Chinese font for ReportLab so Chinese text doesn't show as squares
pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))

@st.cache_data
def load_inventory(file_path="inventory.xlsx", sheet_name="BRITA"):
    
    # 读取指定 sheet（BRITA），并从表中取出第 C(索引2)、G(索引6)、K(索引10) 列，
    # 并重命名为 SKU, Location, SystemQty，做基本清洗。
    
    # 读取原表（保留原有列头）
    df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
    # 确认至少有 11 列
    if df.shape[1] < 11:
        raise ValueError("BRITA 表列数小于 11，无法按 C/G/K 列抽取，请检查文件格式。")
    # 取列：C (index 2), G (index 6), K (index 10)
    cleaned = df.iloc[:, [2, 6, 10]].copy()
    cleaned.columns = ["SKU", "Location", "SystemQty"]
    # 清洗
    cleaned["SKU"] = cleaned["SKU"].astype(str).str.strip()
    cleaned["Location"] = cleaned["Location"].astype(str).str.strip()
    # SystemQty 转为数字（若不能转换设为 0）
    cleaned["SystemQty"] = pd.to_numeric(cleaned["SystemQty"], errors="coerce").fillna(0).astype(int)
    # 去掉 SKU 或 Location 为空的行
    cleaned = cleaned.dropna(subset=["SKU", "Location"])
    # 合并相同 (Location, SKU) 的库存（求和）
    cleaned = cleaned.groupby(["Location", "SKU"], as_index=False)["SystemQty"].sum()
    return cleaned

def generate_cycle_plan(inventory, days=30):
    """
    生成每天的盘点清单（按库位+SKU为行）
    """
    plan = {}
    total = len(inventory)
    per_day = max(1, math.ceil(total / days))
    shuffled = inventory.sample(frac=1, random_state=42).reset_index(drop=True)
    for d in range(days):
        start = per_day * d
        end = start + per_day
        plan[d+1] = shuffled.iloc[start:end].reset_index(drop=True)
    return plan

def save_results(df, suffix="results", name_prefix="cycle_count"):
    today = datetime.date.today().strftime("%Y-%m-%d")
    file_name = f"{name_prefix}_{suffix}_{today}.xlsx"
    df.to_excel(file_name, index=False)
    return file_name


import cv2
from pyzxing import BarCodeReader
from PIL import Image

qr_reader = BarCodeReader()  # ZXing 解码器
def decode_image(image):
    """识别二维码 + 条形码（OpenCV + ZXing）"""
    # 转换成 OpenCV 格式
    img = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)

    # ---------- 识别二维码（ZXing） ----------
    qr_result = qr_reader.decode_array(img)
    if qr_result:
        return qr_result[0].get("raw", None)

    # ---------- 识别条形码（OpenCV） ----------
    detector = cv2.QRCodeDetector()
    data, bbox, _ = detector.detectAndDecode(img)
    if data:
        return data

    return None

def scan_code(label, key):
    """
    Cloud：使用上传图片
    本地：摄像头 + 上传
    """
    st.subheader(label)

    # 判断是否在 Streamlit Cloud
    is_cloud = "STREAMLIT_SERVER_DEployment_TYPE" in os.environ

    if is_cloud:
        img_file = st.file_uploader("上传二维码或条形码图片", type=["jpg", "jpeg", "png"], key=key)
        if img_file:
            img = Image.open(img_file)
            result = decode_image(img)
            if result:
                st.success(f"识别成功：{result}")
                return result
            st.error("未识别到任何二维码或条形码")
        return None

    # ---------------- 本地摄像头模式 ----------------
    cam = st.camera_input("点击拍照扫码", key=key)
    if cam:
        img = Image.open(cam)
        result = decode_image(img)
        if result:
            st.success(f"识别成功：{result}")
            return result
        st.error("未识别到任何二维码或条形码")
        return None

    return None


def create_inventory_report(df, accuracy, shortage_df, overage_df):
    """
    生成 PDF 报告，包含：准确率、差异图、缺货Top、 多货Top。
    文件名：盘点报告_YYYY-MM-DD.pdf
    """
    # 生成差异图（保存为 png）
    fig, ax = plt.subplots(figsize=(8, 4))
    plot_df = df.copy().sort_values("Variance", ascending=False)
    if len(plot_df) > 50:
        plot_df = plot_df.head(50)
    # use pandas plotting for convenience
    plot_df.plot(kind='bar', x='SKU', y='Variance', ax=ax, legend=False, color='steelblue')
    ax.set_title("Inventory Variance Distribution")
    ax.set_xlabel("SKU")
    ax.set_ylabel("Variance")
    plt.tight_layout()
    chart_path = "inventory_chart.png"
    fig.savefig(chart_path, dpi=150)
    plt.close(fig)

    # PDF 名称按中文要求
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    pdf_path = f"盘点报告_{today_str}.pdf"

    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    styles = getSampleStyleSheet()

    # Force styles to use Chinese CID font
    for key in ["Normal", "Title", "Heading2", "Italic"]:
        if key in styles:
            styles[key].fontName = 'STSong-Light'

    story = []

    # 标题
    story.append(Paragraph("<b>📦 盘点分析报告</b>", styles["Title"]))
    story.append(Spacer(1, 12))

    # 基本信息
    story.append(Paragraph(f"生成日期：{today_str}", styles["Normal"]))
    story.append(Spacer(1, 8))
    story.append(Paragraph(f"总体盘点准确率： <b>{accuracy:.2f}%</b>", styles["Normal"]))
    story.append(Spacer(1, 12))

    # 插入差异图
    story.append(Paragraph("<b>差异分布(Variance Distribution)</b>", styles["Heading2"]))
    story.append(Spacer(1, 6))
    story.append(Image(chart_path, width=450, height=250))
    story.append(Spacer(1, 12))

    # 缺货 Top 表
    story.append(Paragraph("<b>📉 缺货 Top 5(Shortage)</b>", styles["Heading2"]))
    story.append(Spacer(1, 6))
    shortage_data = [["Location", "SKU", "SystemQty", "CountedQty", "Variance"]] + shortage_df[["Location","SKU","SystemQty","CountedQty","Variance"]].values.tolist()
    table1 = Table(shortage_data, repeatRows=1)
    table1.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
        ("BACKGROUND", (0,0), (-1,0), colors.lightblue),
        ("FONTNAME", (0,0), (-1,-1), "STSong-Light"),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
    ]))
    story.append(table1)
    story.append(Spacer(1, 12))

    # 多货 Top 表
    story.append(Paragraph("<b>📈 多货 Top 5（Overage）</b>", styles["Heading2"]))
    story.append(Spacer(1, 6))
    overage_data = [["Location", "SKU", "SystemQty", "CountedQty", "Variance"]] + overage_df[["Location","SKU","SystemQty","CountedQty","Variance"]].values.tolist()
    table2 = Table(overage_data, repeatRows=1)
    table2.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgreen),
        ("FONTNAME", (0,0), (-1,-1), "STSong-Light"),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
    ]))
    story.append(table2)
    story.append(Spacer(1, 18))

    story.append(Paragraph("报告说明：本报告由系统自动生成，包含当前盘点结果的差异分析及 Top SKU 列表。", styles["Italic"]))
    story.append(Spacer(1, 6))

    doc.build(story)
    return pdf_path


# ---------------- Streamlit 页面 ----------------
st.set_page_config(page_title="Cycle Count 盘点系统", layout="wide")
st.title("📦 Cycle Count 盘点系统")
st.write("每日自动生成盘点任务（按库位+SKU），支持手机摄像头扫码（前端 JS），导出 Excel 与 PDF 报表。")

# 加载并清洗库存（BRITA）
try:
    inventory = load_inventory()
except Exception as e:
    st.error(f"读取 inventory.xlsx 出错：{e}")
    st.stop()

# 生成 30 天盘点计划（按 Location+SKU 行）
plan = generate_cycle_plan(inventory, days=30)

# 今天盘点清单
today = datetime.date.today()
day_index = (today.day % 30) or 30
daily_list = plan[day_index]
st.subheader(f"📅 今日盘点任务 (Day {day_index}/30)")
st.dataframe(daily_list)

# 保存当天清单（并提供下载）
list_file = save_results(daily_list, "list")
with open(list_file, "rb") as f:
    st.download_button(
        label="📥 下载今日盘点清单（Excel）",
        data=f,
        file_name=list_file,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ---------------- 扫库位 -> 扫SKU -> 输入数量 的交互逻辑 --------------
st.subheader("📲 盘点录入（先扫库位，再扫 SKU）")

# session_state 初始化
if "scanner_id" not in st.session_state:
    st.session_state.scanner_id = 0
if "current_location" not in st.session_state:
    st.session_state.current_location = ""
if "last_scanned_code" not in st.session_state:
    st.session_state.last_scanned_code = ""
if "results" not in st.session_state:
    st.session_state.results = pd.DataFrame(columns=["Location", "SKU", "CountedQty"])

# ---------- 使用前端扫码（streamlit-qrcode-scanner） ----------
st.markdown("**扫码说明**：点击下方“打开摄像头扫描”会请求浏览器相机权限，手机可直接使用摄像头扫码；若无法调用摄像头，请使用下方手动输入。")

# 扫库位
loc_scan = scan_code("📌 扫描库位二维码", "loc_scanner")
if loc_scan:
    st.info(f"检测到库位条码：{loc_scan}")
    if st.button("确认库位", key="confirm_loc"):
        st.session_state.current_location = str(loc_scan).strip()
        st.success(f"当前库位设为：{st.session_state.current_location}")

# 扫 SKU
sku_scan = scan_code("📦 扫描 SKU 条码 / 二维码", "sku_scanner")
if sku_scan:
    st.info(f"检测到 SKU：{sku_scan}")
    if st.button("确认 SKU", key="confirm_sku"):
        st.session_state.last_scanned_code = str(sku_scan).strip()
        st.success(f"当前 SKU：{st.session_state.last_scanned_code}")

# 手动备用输入
st.subheader("或手动输入（若摄像头不可用）")
loc_manual = st.text_input("手动输入/编辑库位：", value=st.session_state.get("current_location",""), key="manual_loc")
sku_manual = st.text_input("手动输入/编辑 SKU：", value=st.session_state.get("last_scanned_code",""), key="manual_sku")

# choose final location & sku for this record (camera or manual)
final_location = (loc_manual or st.session_state.get("current_location","")).strip()
final_sku = (sku_manual or st.session_state.get("last_scanned_code","")).strip()

qty = st.number_input("实盘数量：", min_value=0, step=1)
if st.button("提交记录（保存）"):
    if not final_location or not final_sku:
        st.error("请先填写或扫码库位与 SKU！")
    else:
        df = st.session_state.results
        mask = (df["Location"] == final_location) & (df["SKU"] == final_sku)
        if mask.any():
            # 累加实盘数量
            st.session_state.results.loc[mask, "CountedQty"] = st.session_state.results.loc[mask, "CountedQty"] + int(qty)
        else:
            new_row = pd.DataFrame({"Location":[final_location],"SKU":[final_sku],"CountedQty":[int(qty)]})
            st.session_state.results = pd.concat([st.session_state.results, new_row], ignore_index=True)
        st.success(f"已记录：库位 {final_location} - SKU {final_sku} - 数量 {qty}")
        # 清空扫码缓存（保留库位）
        st.session_state.last_scanned_code = ""
        st.session_state.current_location = final_location

# show temp results
st.subheader("📋 已录入盘点数据（临时）")
st.dataframe(st.session_state.results)

# --------------- Generate final merged report ---------------
if not st.session_state.results.empty:
    if st.button("📊 生成并导出盘点结果（Excel & PDF）"):
        # merge on Location + SKU
        merged = pd.merge(daily_list, st.session_state.results, on=["Location","SKU"], how="left")
        merged["CountedQty"] = merged["CountedQty"].fillna(0).astype(int)
        merged["Variance"] = merged["CountedQty"] - merged["SystemQty"]

        # save excel
        excel_name = save_results(merged, "final", name_prefix="盘点结果")
        with open(excel_name, "rb") as f:
            st.download_button(
                label="📥 下载盘点结果（Excel）",
                data=f,
                file_name=excel_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # analysis
        counted_mask = merged["CountedQty"] != 0
        total_counted = counted_mask.sum()
        correct_counted = ((merged["Variance"] == 0) & counted_mask).sum()
        accuracy = correct_counted / total_counted * 100 if total_counted > 0 else 0
        st.subheader("📈 盘点分析")
        st.metric("盘点准确率", f"{accuracy:.2f}%")

        shortage = merged[merged["Variance"] < 0].sort_values("Variance").head(5)
        overage = merged[merged["Variance"] > 0].sort_values("Variance", ascending=False).head(5)

        col1, col2 = st.columns(2)
        with col1:
            st.write("📉 缺货 Top SKU")
            st.dataframe(shortage[["Location","SKU","SystemQty","CountedQty","Variance"]])
        with col2:
            st.write("📈 多货 Top SKU")
            st.dataframe(overage[["Location","SKU","SystemQty","CountedQty","Variance"]])

        # 差异可视化（页面展示）
        st.subheader("📊 库存差异分布（示意）")
        fig, ax = plt.subplots(figsize=(8,4))
        plot_df = merged.copy().sort_values("Variance", ascending=False)
        if len(plot_df) > 50:
            plot_df = plot_df.head(50)
        plot_df.set_index("SKU")["Variance"].plot(kind="bar", ax=ax)
        ax.set_ylabel("Variance")
        ax.set_title("Variance of each SKU")
        st.pyplot(fig)

        # 生成 PDF 报告并提供下载 (文件名为中文形式)
        pdf_path = create_inventory_report(merged, accuracy, shortage, overage)
        with open(pdf_path, "rb") as f:
            st.download_button(
                label="📄 下载盘点报告 PDF",
                data=f,
                file_name=os.path.basename(pdf_path),
                mime="application/pdf"
            )





