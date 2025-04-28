# 必须放最上面
import streamlit as st
st.set_page_config(page_title="IJOOZ 仓库模拟器", page_icon="🍊", layout="centered")

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
    'Tokyo': 16,
    'Osaka': 4.5,
    'Nagoya': 5,
    'Fukuoka': 5
}

# === 图表功能 ===
def add_charts_to_workbook(wb):
    if "Daily Inventory" not in wb.sheetnames or "Container Schedule" not in wb.sheetnames:
        return

    sched_ws = wb["Container Schedule"]
    inv_ws = wb["Daily Inventory"]

    chart_sheet = wb.create_sheet("Charts")

    chart1 = LineChart()
    chart1.title = "每日库存趋势"
    chart1.height = 10
    chart1.width = 20
    chart1.x_axis.title = "日期"
    chart1.x_axis.numFmt = "yyyy-mm-dd"
    chart1.y_axis.title = "库存单位数"

    data = Reference(inv_ws, min_col=2, max_col=3, min_row=1, max_row=inv_ws.max_row)
    cats = Reference(inv_ws, min_col=1, min_row=2, max_row=inv_ws.max_row)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart_sheet.add_chart(chart1, "A1")

    chart2 = BarChart()
    chart2.title = "每柜生命周期（天）"
    chart2.height = 10
    chart2.width = 20
    chart2.x_axis.title = "PO"
    chart2.y_axis.title = "生命周期（天）"

    po_col = life_col = None
    for i, cell in enumerate(sched_ws[1], 1):
        if cell.value == "PO":
            po_col = i
        if cell.value == "生命周期（天）":
            life_col = i
    if po_col and life_col:
        data2 = Reference(sched_ws, min_col=life_col, min_row=1, max_row=sched_ws.max_row)
        cats2 = Reference(sched_ws, min_col=po_col, min_row=2, max_row=sched_ws.max_row)
        chart2.add_data(data2, titles_from_data=True)
        chart2.set_categories(cats2)
        chart_sheet.add_chart(chart2, "A20")

# === 单仓库模拟 ===
def run_simulation(file, warehouse_name):
    if warehouse_name not in warehouse_capacities:
        raise ValueError(f"未定义仓库容量：{warehouse_name}")

    capacity = warehouse_capacities[warehouse_name]
    xls = pd.ExcelFile(file)
    container_sheet = f"Container-{warehouse_name}"
    usage_sheet = f"weekly usage-{warehouse_name}"
    if container_sheet not in xls.sheet_names or usage_sheet not in xls.sheet_names:
        raise ValueError(f"缺少 sheet：{container_sheet} 或 {usage_sheet}")

    container_df = xls.parse(container_sheet)
    usage_df = xls.parse(usage_sheet)

    container_df['HARVEST DAY'] = pd.to_datetime(container_df['HARVEST DAY'])
    container_df['ETA'] = pd.to_datetime(container_df['ETA'])
    container_df['ETD'] = pd.to_datetime(container_df['ETD']) if 'ETD' in container_df else container_df['ETA'] - pd.Timedelta(days=30)

    usage_df[['year', 'week_number']] = usage_df['week'].str.extract(r'(\d{4})WK(\d{2})').astype(int)
    usage_df['monday'] = pd.to_datetime(usage_df['year'].astype(str) + usage_df['week_number'].astype(str) + '1', format='%G%V%u')

    daily_usage = []
    for _, row in usage_df.iterrows():
        for i in range(7):
            daily_usage.append({
                'date': row['monday'] + pd.Timedelta(days=i),
                'daily_usage': row['用量'] / 7
            })
    daily_usage_df = pd.DataFrame(daily_usage)

    today = pd.Timestamp(datetime.date.today())
    containers = []
    for _, row in container_df.iterrows():
        containers.append({
            'PO': row['PO'],
            'Vessel': row.get('Vessel', ''),
            'harvest_day': row['HARVEST DAY'],
            'eta': row['ETA'],
            'etd': row['ETD'],
            'unit': float(row['单位']),
            'in_ext_date': None,
            'in_ijooz_date': today if pd.isnull(row['ETA']) else None,
            'start_use': None,
            'end_use': None,
            'used': 0,
            'can_enter_date': row['ETA'] + pd.Timedelta(days=3) if pd.notnull(row['ETA']) else today
        })

    ijooz_storage = [c for c in containers if c['in_ijooz_date'] is not None]
    external_storage = []
    inventory_log = []
    date_range = pd.date_range(start=today, end=daily_usage_df['date'].max())

    for day in date_range:
        used_today = []
        day_usage = daily_usage_df.loc[daily_usage_df['date'] == day, 'daily_usage'].sum()
        day_usage_original = day_usage

        while day_usage > 0 and ijooz_storage:
            c = ijooz_storage[0]
            if day < c['in_ijooz_date']:
                break
            remain = c['unit'] - c['used']
            if c['start_use'] is None:
                c['start_use'] = day
            use_now = min(day_usage, remain)
            c['used'] += use_now
            day_usage -= use_now
            if c['used'] == c['unit']:
                c['end_use'] = day
                ijooz_storage.pop(0)
            used_today.append(c['PO'])

        current_total = sum(c['unit'] - c['used'] for c in ijooz_storage)

        for c in containers:
            if c['in_ijooz_date'] is None and c['in_ext_date'] is None and c['can_enter_date'] <= day:
                if capacity - current_total >= 1:
                    c['in_ijooz_date'] = day
                    ijooz_storage.append(c)
                    current_total += c['unit']
                else:
                    c['in_ext_date'] = day
                    external_storage.append(c)

        external_storage.sort(key=lambda x: x['in_ext_date'])
        for c in external_storage[:]:
            if capacity - current_total >= 1:
                c['in_ijooz_date'] = day
                ijooz_storage.append(c)
                external_storage.remove(c)
                current_total += c['unit']
            else:
                break

        transit = sum(c['unit'] for c in containers if c['etd'] <= day < c['eta'])
        po_placed = sum(c['unit'] for c in containers if day < c['etd'])

        inventory_log.append({
            '日期': day,
            'IJOOZ 仓库库存（单位）': round(sum(c['unit'] - c['used'] for c in ijooz_storage), 1),
            '外部冷库库存（整柜数）': len(external_storage),
            '当天使用的货柜 PO': ', '.join(set(used_today)),
            '使用柜数量': len(set(used_today)),
            '总库存（单位）': round(sum(c['unit'] - c['used'] for c in ijooz_storage) + sum(c['unit'] for c in external_storage), 1),
            '运输中（单位）': round(transit, 1),
            'PO Placed（单位）': round(po_placed, 1),
            'daily_usage': round(day_usage_original, 2)
        })

    inventory_df = pd.DataFrame(inventory_log)
    inventory_df['日期'] = pd.to_datetime(inventory_df['日期']).dt.strftime('%Y-%m-%d')

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        inventory_df.to_excel(writer, index=False, sheet_name="Daily Inventory")
    output.seek(0)

    wb = load_workbook(output)
    add_charts_to_workbook(wb)

    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    if st.session_state.get("run_all_mode") and warehouse_name in ["Tokyo", "Osaka", "Nagoya", "Fukuoka"]:
        return inventory_df, final_output
    else:
        return final_output

# === 批量生成 ===
def run_all_simulations(file):
    st.session_state["run_all_mode"] = True
    xls = pd.ExcelFile(file)
    available = [n.replace("Container-", "") for n in xls.sheet_names if n.startswith("Container-")]

    all_inventory_dfs = []
    with tempfile.TemporaryDirectory() as tmpdirname:
        excel_paths = []
        for wh in available:
            try:
                result = run_simulation(file, wh)
                if isinstance(result, tuple):
                    inventory_df, sim_output = result
                    all_inventory_dfs.append((wh, inventory_df))
                else:
                    sim_output = result
                filename = f"{wh}_simulation.xlsx"
                path = os.path.join(tmpdirname, filename)
                with open(path, "wb") as f:
                    f.write(sim_output.read())
                excel_paths.append(path)
            except Exception as e:
                st.warning(f"仓库 {wh} 模拟失败：{e}")

        # === Japan 总括 ===
        japan_dfs = [df for wh, df in all_inventory_dfs if wh in ["Tokyo", "Osaka", "Nagoya", "Fukuoka"]]
        if japan_dfs:
            combined = pd.concat(japan_dfs)
            combined['日期'] = pd.to_datetime(combined['日期'])
            japan_daily = combined.groupby("日期", as_index=False).agg({
                "IJOOZ 仓库库存（单位）": "sum",
                "外部冷库库存（整柜数）": "sum",
                "使用柜数量": "sum",
                "运输中（单位）": "sum",
                "PO Placed（单位）": "sum",
                "daily_usage": "sum",
                "当天使用的货柜 PO": lambda x: ', '.join(filter(None, map(str, x)))
            })
            japan_daily['日期'] = japan_daily['日期'].dt.strftime('%Y-%m-%d')

            combined['周'] = pd.to_datetime(combined['日期']).dt.isocalendar().week
            japan_weekly = combined.groupby("周", as_index=False).agg({
                "daily_usage": "sum",
                "IJOOZ 仓库库存（单位）": "mean",
                "外部冷库库存（整柜数）": "mean",
                "运输中（单位）": "mean",
                "PO Placed（单位）": "mean"
            })
            japan_weekly.rename(columns={
                "daily_usage": "周累计用量（单位）",
                "IJOOZ 仓库库存（单位）": "周平均IJOOZ库存",
                "外部冷库库存（整柜数）": "周平均外库库存",
                "运输中（单位）": "周平均运输中单位",
                "PO Placed（单位）": "周平均PO Placed单位"
            }, inplace=True)
            japan_weekly = japan_weekly.round(1)

            japan_path = os.path.join(tmpdirname, "Japan_Daily_Inventory.xlsx")
            with pd.ExcelWriter(japan_path, engine="openpyxl") as writer:
                japan_daily.to_excel(writer, index=False, sheet_name="Japan Daily Inventory")
                japan_weekly.to_excel(writer, index=False, sheet_name="Japan Weekly Summary")
            excel_paths.append(japan_path)

        zip_output = BytesIO()
        with zipfile.ZipFile(zip_output, "w") as zipf:
            for path in excel_paths:
                zipf.write(path, arcname=os.path.basename(path))
        zip_output.seek(0)
        return zip_output

# === 界面 ===
st.title("🍊 IJOOZ 仓库模拟器")
warehouse_options = ["全部仓库"] + list(warehouse_capacities.keys())
warehouse_name = st.selectbox("📍 选择仓库地点", warehouse_options)
uploaded_file = st.file_uploader("📤 上传Excel文件", type=["xlsx", "xls"])

if uploaded_file and st.button("🚀 运行模拟"):
    try:
        today_str = datetime.date.today().strftime('%Y-%m-%d')
        with st.spinner("模拟进行中..."):
            if warehouse_name == "全部仓库":
                zip_output = run_all_simulations(uploaded_file)
                st.success("✅ 所有仓库模拟完成！")
                st.download_button("📦 下载ZIP文件", data=zip_output, file_name=f"IJOOZ_Simulation_ALL_{today_str}.zip", mime="application/zip")
            else:
                output = run_simulation(uploaded_file, warehouse_name)
                st.success("✅ 仓库模拟完成！")
                st.download_button("📥 下载Excel文件", data=output, file_name=f"IJOOZ_Simulation_{warehouse_name}_{today_str}.xlsx")
    except Exception as e:
        st.error(f"❌ 出错了：{str(e)}")
