from pydoc import doc
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import datetime
import math
import os
from io import BytesIO

# PDF libs
from reportlab.lib.pagesizes import A4 # 纸张的大小
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle #字体，表格，图片，段落文件常用组件，PDF中可放的元素
from reportlab.lib.styles import getSampleStyleSheet #获取预定义的文字样式
from reportlab.lib import colors

# webcam + barcode
from streamlit_webrtc import webrtc_streamer, VideoProcessorBase, WebRtcMode #streamlit_webrtc库是用来调用streamlit网页中摄像头视频流，webrtc_streamer打开摄像头视频流组件，VideoProcessorBase定义如何处理每一帧视频，比如识别二维码， WebRtcMode设置WebRTC通信模式，比如发送视频，接收视频
from pyzbar import pyzbar #是用来识别条形码和二维码的Python库，可读取摄像头捕获的内容并识别
import av #用于从摄像头读取实时的视频帧，可以逐帧分析画面
import cv2 # 用于图像处理，把识别的结果显示在视频画面上
import numpy as np

#----------------------------Helper Functions----------------------------------------------------
@st.cache_data #缓冲存储，不用每次刷新页面都要重新加载
def load_inventory(file_path="inventory.xlsx", sheet_name="BRITA"):
    #从BRITA sheet提取C/G/K（索引2,6,10）列并清洗数据
    df = pd.read_excel(file_path,sheet_name=sheet_name, dtype=str)
    if df.shape[1] < 11: # df.shape【0】代表行数， df.shape[1]代表列数
        raise ValueError("BRITA 表列数小于11，无法抽取")
    cleaned= df.iloc[:, [2, 6,10]].copy()
    cleaned.columns = ["SKU", "Location", "SystemQty"]
    #清洗
    cleaned["SKU"] = cleaned["SKU"].astype(str).str.strip()#将SKU列数据转换为字符串，然后再去除空格
    cleaned["Location"]=cleaned["Location"].astype(str).str.strip()
    cleaned["Location"] = pd.to_numeric(cleaned["SystemQty"], errors="coerce").fillna(0).astype(int)#先将SystemQty都转化为数值，如遇报错，非数值转化时会报错，报错均转化为Nan,再将Nan转化为0，再将全部转为为整数
    cleaned = cleaned.dropna(subset=["SKU", "Location"]) #将SKU与Location里面的空值去掉
    cleaned = cleaned.groupby(["SKU", "Location"], as_index=False)["SystemQty"].sum()
    return cleaned 
def generate_cycle_plan(inventory, days=30):
    plan = {}
    total = len(inventory)
    per_day = math.ceil(total / days)
    shuffled = inventory.sample(frac=1, random_state=42).reset_index(drop=True)
    for d in range(days):
        start = d * per_day
        end = start + per_day
        plan[d+1] = shuffled.iloc[start:end]
    return plan
def save_results(df, suffix="results", name_prefix="cycle_count"):
    today = datetime.date.today().strftime("%Y-%m-%d")
    file_name = f"{name_prefix}_{suffix}_{today}.xlsx"#f字符串格式化，将{}中内容插入字符串中
    df.to_excel(file_name, index=False) # pandas生成文件后会自动带索引，index=false将索引去掉
    return file_name
def create_inventory_report(df, accuracy, shortage_df, overage_df):
    fig, ax = plt.subplots(figsize=(8,4)) # fig,ax 定图纸和坐标轴， 8,4 单位是英寸，1英寸=2.53厘米
    plot_df = df.sort_values("Variance", ascending=False)
    if len(plot_df) >  50:
        plot_df = plot_df.head(50)
    plot_df.plot(kind="bar", x="SKU", y="Variance", ax=ax, legend=False, color="steelblue")
    ax.set_title("Variance Distribution")
    ax.set_xlabel("SKU")
    ax.set_ylabel("Variance")
    plt.tight_layout()
    chart_path = "inventory_chart.png"
    fig.savefig(chart_path, dpi=150)
    plt.close(fig)

    today_str = datetime.date.today().strftime("%Y-%m-%d")
    pdf_filename = f"盘点报告_{today_str}.pdf"
    styles = getSampleStyleSheet()#获取PDF所需所有的字体格式
    story= [] # 创建空列表，为后续填充内容用

    #报告标题
    story.append(Paragraph("<b>📦 Inventory Cycle Count Report<b>", styles["Title"]))# <b> 字体加粗
    story.append(Spacer(1,20)) #添加空格，1dot，0.23mm，20个dot高, 7cm

    # 添加总体盘点差异率
    story.append(Paragraph(f"✅ Overall Accuracy: <b>{accuracy:.2f}%</b>", styles["Normal"]))
    story.append(Spacer(1,15))

    #添加差异图
    story.append(Paragraph("<b>Variance Distribution<b>", styles["Heading2"]))
    story.append(Image(chart_path, width=400, height=300))
    story.append(Spacer(1,20))

    #添加缺货Top SKU 表格
    story.append(Paragraph("<b>Shortage Top 5 SKUs<b>", styles["Heading2"]))
    shortage_data =  [["SKU", "SystemQty", "CountedQty", "Variance"]] + shortage_df.values.tolist() # shortafe_df是数据结构，不能直接读取，.values先转换为二维数组，包含里面的数据和数据类型，.tolist转为数列
    table1 = Table(shortage_data)
    table1.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
        ("BACKGROUND", (0,0), (-1,0), colors.lightblue),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
    ]))      
    story.append(table1)
    story.append(Spacer(1, 20))

     # 4️⃣ 添加多货Top SKU表格
    story.append(Paragraph("📈 <b>Overage Top 5 SKUs</b>", styles["Heading2"]))
    overage_data = [["SKU", "SystemQty", "CountedQty", "Variance"]] + overage_df.values.tolist()
    table2 = Table(overage_data)
    table2.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgreen),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
    ]))
    story.append(table2)
    story.append(Spacer(1, 30))

    # 5️⃣ 添加底部备注
    story.append(Paragraph("Report automatically generated by Cycle Count System.", styles["Italic"]))

    # 生成 PDF
    doc.build(story)
    return pdf_filename

class BarcodeProcessor(VideoProcessorBase):
    def __init__(self):
        self.last_code = None
        self.last_time = None
    
    def recv(self, frame: av.VideoFrame) -> av.VideoFrame:
        img = frame.to_ndarray(format="bgr24")
        # convert to grayscale for better barcode detection
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        barcodes = pyzbar.decode(gray)
        if barcodes:
            # take first barcoed
            barcode = barcodes[0]
            data = barcode.data.decode("utf-8")
            self.last_code = data
            self.last_time = datetime.datetime.now().isoformat()
            # drawm rectangle and text on image for visual feedback
            (x, y, w, h) = barcode.rect
            cv2.rectangle(img, (x,y), (x+w, y+h), (0, 255, 0), 2)
            cv2.putText(img, data, (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0,255,0), 2)
            return av.VideoFrame.from_ndarray(img, format="bgr24")
        
# -----------------------Streamlit 页面-----------------------------------------------------------
st.set_page_config(page_title="Cycle Count 盘点系统", layout="wide")
st.title("📦 Cycle Count 盘点系统(支持一维/二维扫码)")
st.write("可使用手机摄像头扫码： 先扫库位， 再扫SKU(可连续多个), 如摄像头不可用, 可手动输入")

# load inventory
try:
    inventory = load_inventory()
except Exception as e:
    st.error(f"读取 inventory.xlsx 出错: {e}")
    st.stop()

plan = generate_cycle_plan(inventory, days=30)
today = datetime.date.today()
day_index = (today.day % 30) or 30
daily_list = plan[day_index]

st.subheader(f"📅 今日盘点任务 Day {day_index}/30")
st.dataframe(daily_list)

# save today's list for download
list_file = save_results(daily_list, "list")
with open(list_file, "rb") as f:
    st.download_button("下载今日盘点清单", data=f, file_name=list_file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# session state init
if "scanner_id" not in st.session_state:
    st.session_state.scanner_id = 0
if "current_location" not in st.session_state:
    st.session_state.current_location = ""
if "last_scanned_code" not in st.session_state:
    st.session_state.last_scanned_code = ""
if "results" not in st.session_state:
    st.session_state.results = pd.DataFrame(columns=["Location", "SKU", "CountQty"])

#------------------webrtc scanner----------------------------------------------------------
st.subheader("📲摄像头扫码（一维/二维）")
st.write("使用摄像头扫码时:先点击“打开摄像头并扫描库位”,扫描成功后点击“确认库位”,然后切换至SKU扫描并点击“确认SKU”")
# create 2 streamers(one can be reused)--we will use same processor but different keys
loc_col1, loc_col2, loc_col3 = st.columns([1,1,1])
with loc_col1:
    if st.button("打开摄像头并扫描库位"):
        st.session_state.scanner_id += 1
        st.session_state.loc_stream_key =f"loc_stream_{st.session_state.scanner_id}"
        st.session_state.show_loc_stream = True
# show location stream if requested
if st.session_state.get("show_loc_stream", False):
    ctx_loc = webrtc_streamer(
        key=st.session_state.get("loc_stream_key", "loc_stream"),
        video_processor_factory=BarcodeProcessor,
        media_stream_constraints={"video":True, "audio":False},
        async_processing=True,
        mode=WebRtcMode.SENDRECV,
        video_html_attrs={"style":"width:320ox; height:auto;"}
    )
    # fetch detected code(if any)
    if ctx_loc and ctx_loc.video_processor:
        code = ctx_loc.video_processor.last_code
        if code:
            st.info(f"摄像头检测到条码：{code}")
            if st.button("确认库位", key="confirm_loc"):
                st.session_state.current_location = str(code).strip()
                st.session_state.show_loc_stream = False
                st.success(f"当前库位设为：{st.session_state.current_location}")
    # SKU scanning
    sku_col1, sku_col2 = st.columns([1,1])
    with sku_col1:
        if st.button("打开摄像头并扫描SKU"):
            st.session_state.scanner_id += 1
            st.session_state.sku_stream_key = f"sku_stream_{st.session_state.scanner_id}"
            st.session_state.show_sku_stream = True
    if st.session_state.get("show_sku_stream", False):
        ctx_sku = webrtc_streamer(
            key=st.session_state.get("sku_stream_key","sku_stream"),
            video_processor_factory=BarcodeProcessor,
            media_stream_constraints={"video":True, "audio":False},
            async_processing=True,
            mode=WebRtcMode.SENDRECV,
            video_html_attrs={"style":"width:320ox; height:auto;"}
        )
        if ctx_sku and ctx_sku.video_processor:
            code = ctx_sku.video_processor.last_code
            if code:
                st.info(f"摄像头检测到SKU: {code}")
                if st.button("确认SKU", key=f"confirm_sku_{st.session_state.scanner_id}"):
                    st.session_state.last_scanned_code = str(code).strip()
                    st.session_state.show_sku_stream = False
                    st.success(f"当前SKU: {st.session_state.last_scanne_codeZ}")
#----------Manual inputs------------------------------------------------------------
st.subheader("或手动输入（若摄像头不可用）")
loc_manual = st.text_input("手动输入/编辑库位:", value=st.session_state.get("current_location",""), key="manual_loc")
sku_manual = st.text_input("手动输入/编辑SKU:", value=st.session_state.get("last_scanned_code", ""), key="manual_sku")
#choose final location & sku for this record(camera or manual)
final_location = loc_manual.strip()
final_sku = sku_manual.strip()
qty =st.number_input("实盘数量: ", min_value=0, step=1)
if st.button("提交记录(保存)"):
    if not final_location or not final_sku:
        st.error("请先填写或扫码库位与SKU!")
    else:
        #add or aggregate if same location+SKU exists in session results
        df = st.session_state.results
        mask = (df["Location"] == final_location) & (df["SKU"] == final_sku)
        if mask.any():
            st.session_state.resultd.loc[mask, "CountedQty"] = st.session_state.results.loc[mask, "CountedQty"] + int(qty)
        else:
            new_row = pd.DataFrame({"Location":[final_location], "SKU":[final_sku], "CountedQty":[int(qty)]})
            st.session_state.resultd = pd.concat([st.session_state.results, new_row], ignore_index=True)
        st.success(f"已记录: 库位{final_location} - SKU{final_sku} - 实盘{qty}")       
        #清空扫码缓存（保留手动输入设计）
        st.session_state.last_scanned_code = ""
        st.session_state.current_location = final_location
    #show temp results
    st.subheader("已记录盘点数据（临时）")
    st.dataframe(st.session_state.results)
#----------generate final merged report-------------------------------------------
if not st.session_state.results.empty:
    if st.button("📊 生成盘点结果(excel & PDF)"):
        # merged on location+SKU
        # ensure daily_list has location+SKU
        # daily_list in earlier part is based on inventory rows
        merged = pd.merge(daily_list, st.session_state.results, on=["location", "SKU"], how="left")
        merged["CountedQty"] = merged["CountedQty"].fillna(0).astype(int)
        merged["Variance"] = merged["CountedQty"] - merged["SystemQty"]
        
        #save excel
        excel_name = save_results(merged, "final", name_prefix="盘点结果")
        with open(excel_name, "rb") as f:
            st.download_button("📥 点击下载盘点报表", data=f, file_name=excel_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        # analysis
        counted_mask = merged["CountedQty"] != 0
        total_counted = counted_mask.sum()
        correct_counted = ((merged["Variance"] == 0) & counted_mask).sum()
        accuracy = correct_counted / total_counted * 100 if total_counted > 0 else 0
        st.metric("盘点准确率", f"{accuracy:.2f}%")

        shortage = merged[merged["Variance"] < 0].sort_values("Variance").head(5)
        overage = merged[merged["Variance"] > 0].sort_values("Variance", ascending=False).head(5)

        col1, col2 = st.columns(2)
        with col1:
            st.write("📉 缺货 Top SKU")
            st.dataframe(shortage[["SKU", "SystemQty", "CountedQty", "Variance"]])
        with col2:
            st.write("📈 多货 Top SKU")
            st.dataframe(overage[["SKU", "SystemQty", "CountedQty", "Variance"]])
        
        #差异可视化
        st.subheader("库存差异分布")
        fig, ax = plt.subplots()
        merged.set_index("SKU")["Variance"].plot(kind="bar", ax=ax)
        ax.set_ylabel("Variance")
        ax.set_title("Variance of each SKU")
        st.pyplot(fig)

        # -----生成PDF文件并添加下载按钮------
        pdf_path = create_inventory_report(merged, accuracy, shortage, overage)
        with open(pdf_path, "rb") as f:
            st.download_button(
                label="📄 下载盘点报告 PDF",
                data=f,
                file_name="inventory_report.pdf",
                mime="application/pdf"
            )
