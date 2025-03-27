import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
import zipfile
import os
import tempfile
from openpyxl import load_workbook
from openpyxl.chart import LineChart, BarChart, Reference

# === 仓库容量配置 ===
warehouse_capacities = {
    'Singapore': 16,
    'Tokyo': 15,
    'Osaka': 5,
    'Nagoya': 7,
    'Fukuoka': 5,
    'Default': 5
}

# 设置页面
st.set_page_config(page_title="IJOOZ 仓库模拟器", page_icon="🍊", layout="centered")
st.markdown('<p class="title-text">🍊 IJOOZ 仓库模拟器</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle-text">上传仓库使用计划 Excel 文件，自动计算库存及生命周期，并生成图表。</p>', unsafe_allow_html=True)
st.markdown("---")

# 选择仓库
warehouse_options = list(warehouse_capacities.keys())
warehouse_options.insert(0, '全部仓库')
warehouse_name = st.selectbox("📍 选择仓库地点", warehouse_options, index=0)
uploaded_file = st.file_uploader("📤 上传 Excel 文件", type=["xlsx", "xls"])

# 图表函数
def add_charts_to_workbook(wb):
    if "Daily Inventory" not in wb.sheetnames or "Container Schedule" not in wb.sheetnames:
        return
    chart_sheet = wb.create_sheet("Charts")

    inv_ws = wb["Daily Inventory"]
    chart1 = LineChart()
    chart1.title = "每日库存趋势"
    chart1.y_axis.title = "库存单位数"
    chart1.x_axis.title = "日期"
    chart1.height = 10
    chart1.width = 20
    chart1.x_axis.majorTickMark = "out"
    chart1.x_axis.tickLblPos = "low"
    chart1.y_axis.tickLblPos = "low"
    chart1.y_axis.majorGridlines = None
    data = Reference(inv_ws, min_col=2, max_col=3, min_row=1, max_row=inv_ws.max_row)
    categories = Reference(inv_ws, min_col=1, min_row=2, max_row=inv_ws.max_row)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(categories)
    chart_sheet.add_chart(chart1, "A1")

    sched_ws = wb["Container Schedule"]
    chart2 = BarChart()
    chart2.title = "每柜生命周期（天）"
    chart2.y_axis.title = "生命周期（天）"
    chart2.x_axis.title = "PO"
    chart2.height = 10
    chart2.width = 20
    chart2.y_axis.tickLblPos = "low"
    chart2.y_axis.majorTickMark = "out"
    chart2.y_axis.majorGridlines = None

    po_col, life_col = 1, 1
    for col in range(1, sched_ws.max_column + 1):
        if sched_ws.cell(1, col).value == "PO":
            po_col = col
        if sched_ws.cell(1, col).value == "生命周期（天）":
            life_col = col

    data2 = Reference(sched_ws, min_col=life_col, min_row=1, max_row=sched_ws.max_row)
    categories2 = Reference(sched_ws, min_col=po_col, min_row=2, max_row=sched_ws.max_row)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(categories2)
    chart_sheet.add_chart(chart2, "A20")

# 单仓库模拟函数（粘贴原来的 run_simulation 函数）
def run_simulation(file, warehouse_name):
    from warehouse_simulator import simulate_warehouse  # 替换为你的模拟逻辑函数
    output = simulate_warehouse(file, warehouse_name)  # 获取 BytesIO
    wb = load_workbook(output)
    add_charts_to_workbook(wb)
    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output

# 批量生成 + 打包 zip
def run_all_simulations(file):
    xls = pd.ExcelFile(file)
    available_warehouses = [name.replace("Container-", "") 
                            for name in xls.sheet_names 
                            if name.startswith("Container-")]

    with tempfile.TemporaryDirectory() as tmpdirname:
        excel_paths = []
        for wh in available_warehouses:
            try:
                sim_output = run_simulation(file, wh)
                filename = f"{wh}_simulation.xlsx"
                file_path = os.path.join(tmpdirname, filename)
                with open(file_path, "wb") as f:
                    f.write(sim_output.read())
                excel_paths.append(file_path)
            except Exception as e:
                st.warning(f"⚠️ 仓库 {wh} 模拟失败：{e}")

        zip_output = BytesIO()
        with zipfile.ZipFile(zip_output, "w") as zipf:
            for path in excel_paths:
                arcname = os.path.basename(path)
                zipf.write(path, arcname=arcname)
        zip_output.seek(0)
        return zip_output

# 主入口逻辑
if uploaded_file and st.button("🚀 运行模拟"):
    try:
        today_str = datetime.date.today().strftime('%Y-%m-%d')
        with st.spinner("模拟进行中，请稍候..."):
            if warehouse_name == '全部仓库':
                output_zip = run_all_simulations(uploaded_file)
                filename = f"IJOOZ_Simulation_ALL_{today_str}.zip"
                st.success("✅ 所有仓库模拟完成！点击下方按钮下载所有结果：")
                st.download_button("📦 下载 ZIP 文件", data=output_zip, file_name=filename, mime="application/zip")
            else:
                output_excel = run_simulation(uploaded_file, warehouse_name)
                filename = f"IJOOZ_Simulation_{warehouse_name}_{today_str}.xlsx"
                st.success("✅ 模拟完成！点击下方按钮下载结果：")
                st.download_button("📥 下载 Excel 文件", data=output_excel, file_name=filename)
    except Exception as e:
        st.error(f"❌ 出错了：{str(e)}")
