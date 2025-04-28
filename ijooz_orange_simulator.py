# 必须放最上面
import streamlit as st
st.set_page_config(page_title="IJOOZ 仓库模拟器", page_icon="🍊", layout="centered")

# === 页面布局 ===
st.title("🍊 IJOOZ 仓库模拟器")
st.markdown("上传仓库使用计划 Excel 文件，自动计算库存及生命周期，并生成图表。")
st.markdown("---")

# === 仓库选项 ===
warehouse_capacities = {
    'Singapore': 16,
    'Tokyo': 16,
    'Osaka': 4.5,
    'Nagoya': 5,
    'Fukuoka': 5
}

warehouse_options = list(warehouse_capacities.keys())
warehouse_options.insert(0, "全部仓库")
warehouse_name = st.selectbox("📍 选择仓库地点", warehouse_options, index=0)

uploaded_file = st.file_uploader("📤 上传 Excel 文件", type=["xlsx", "xls"])

# === 主执行逻辑 ===
if uploaded_file and st.button("🚀 运行模拟"):

    # 👉 运行时再import
    import pandas as pd
    import datetime
    from io import BytesIO
    import zipfile
    import os
    import tempfile
    from openpyxl import load_workbook
    from openpyxl.chart import LineChart, BarChart, Reference

    # === 定义功能函数 ===
    def add_charts_to_workbook(wb):
        if "Daily Inventory" not in wb.sheetnames:
            return
        inv_ws = wb["Daily Inventory"]
        inv_headers = [cell.value for cell in inv_ws[1]]
        inv_data = []
        for row in inv_ws.iter_rows(min_row=2, values_only=True):
            row_dict = dict(zip(inv_headers, row))
            for col in ["IJOOZ 仓库库存（单位）", "总库存（单位）"]:
                if col in row_dict and isinstance(row_dict[col], (int, float)):
                    row_dict[col] = round(row_dict[col], 1)
            inv_data.append(row_dict)

        wb.remove(inv_ws)
        new_inv_ws = wb.create_sheet("Daily Inventory")
        for c_idx, col in enumerate(inv_headers, 1):
            new_inv_ws.cell(row=1, column=c_idx, value=col)
        for r_idx, row in enumerate(inv_data, 2):
            for c_idx, col in enumerate(inv_headers, 1):
                new_inv_ws.cell(row=r_idx, column=c_idx, value=row.get(col))

    def run_simulation(file, warehouse_name):
        ijooz_capacity = warehouse_capacities[warehouse_name]
        xls = pd.ExcelFile(file)
        container_sheet = f"Container-{warehouse_name}"
        usage_sheet = f"weekly usage-{warehouse_name}"

        container_df = xls.parse(container_sheet)
        weekly_usage_df = xls.parse(usage_sheet)

        container_df['HARVEST DAY'] = pd.to_datetime(container_df['HARVEST DAY'])
        container_df['ETA'] = pd.to_datetime(container_df['ETA'])
        if 'ETD' in container_df.columns:
            container_df['ETD'] = pd.to_datetime(container_df['ETD'])
        else:
            container_df['ETD'] = container_df['ETA'] - pd.Timedelta(days=30)

        weekly_usage_df[['year', 'week_number']] = weekly_usage_df['week'].str.extract(r'(\d{4})WK(\d{2})').astype(int)
        weekly_usage_df['monday'] = pd.to_datetime(
            weekly_usage_df['year'].astype(str) + weekly_usage_df['week_number'].astype(str) + '1', format='%G%V%u'
        )

        daily_usage_records = []
        for _, row in weekly_usage_df.iterrows():
            daily_usage = row['用量'] / 7
            for i in range(7):
                daily_usage_records.append({
                    'date': row['monday'] + pd.Timedelta(days=i),
                    'daily_usage': daily_usage
                })
        daily_usage_df = pd.DataFrame(daily_usage_records)

        today = pd.Timestamp(datetime.date.today())
        containers = []
        for idx, row in container_df.iterrows():
            eta = row['ETA']
            etd = row['ETD']
            containers.append({
                'index': idx,
                'PO': row['PO'],
                'Vessel': row.get('Vessel', ''),
                'harvest_day': row['HARVEST DAY'],
                'eta': eta,
                'etd': etd,
                'unit': float(row['单位']),
                'in_ext_date': None,
                'in_ijooz_date': today if pd.isnull(eta) else None,
                'start_use': None,
                'end_use': None,
                'used': 0,
                'can_enter_date': eta + pd.Timedelta(days=3) if pd.notnull(eta) else today
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
                if c['in_ijooz_date'] is not None and day < c['in_ijooz_date']:
                    break
                remaining = c['unit'] - c['used']
                if c['start_use'] is None:
                    c['start_use'] = day
                use_now = min(day_usage, remaining)
                c['used'] += use_now
                day_usage -= use_now
                if c['used'] == c['unit']:
                    c['end_use'] = day
                    ijooz_storage.pop(0)
                used_today.append(c['PO'])

            current_total_inventory = sum(x['unit'] - x['used'] for x in ijooz_storage)

            for c in containers:
                if c['in_ijooz_date'] is None and c['in_ext_date'] is None:
                    if c['can_enter_date'] <= day:
                        if ijooz_capacity - current_total_inventory >= 1:
                            c['in_ijooz_date'] = day
                            ijooz_storage.append(c)
                            current_total_inventory += c['unit']
                        else:
                            c['in_ext_date'] = day
                            external_storage.append(c)

            external_storage.sort(key=lambda x: x['in_ext_date'])
            for c in external_storage[:]:
                if ijooz_capacity - current_total_inventory >= 1:
                    c['in_ijooz_date'] = day
                    ijooz_storage.append(c)
                    external_storage.remove(c)
                    current_total_inventory += c['unit']
                else:
                    break

            in_transit_units = sum(c['unit'] for c in containers if c['etd'] <= day < c['eta'])
            po_placed_units = sum(c['unit'] for c in containers if day < c['etd'])

            inventory_log.append({
                '日期': day,
                'IJOOZ 仓库库存（单位）': round(sum(c['unit'] - c['used'] for c in ijooz_storage), 1),
                '外部冷库库存（整柜数）': len(external_storage),
                '当天使用的货柜 PO': ', '.join(set(used_today)),
                '使用柜数量': len(set(used_today)),
                '总库存（单位）': round(
                    sum(c['unit'] - c['used'] for c in ijooz_storage) +
                    sum(c['unit'] for c in external_storage), 1
                ),
                '运输中（单位）': round(in_transit_units, 1),
                'PO Placed（单位）': round(po_placed_units, 1),
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

    def run_all_simulations(file):
        st.session_state["run_all_mode"] = True

        xls = pd.ExcelFile(file)
        available_warehouses = [name.replace("Container-", "") for name in xls.sheet_names if name.startswith("Container-")]

        all_inventory_dfs = []
        with tempfile.TemporaryDirectory() as tmpdirname:
            excel_paths = []
            for wh in available_warehouses:
                try:
                    result = run_simulation(file, wh)

                    if isinstance(result, tuple):
                        inventory_df, sim_output = result
                        all_inventory_dfs.append((wh, inventory_df))
                    else:
                        sim_output = result

                    filename = f"{wh}_simulation.xlsx"
                    file_path = os.path.join(tmpdirname, filename)
                    with open(file_path, "wb") as f:
                        f.write(sim_output.read())
                    excel_paths.append(file_path)
                except Exception as e:
                    st.warning(f"⚠️ 仓库 {wh} 模拟失败：{e}")

            if all_inventory_dfs:
                japan_dfs = [df for wh, df in all_inventory_dfs if wh in ["Tokyo", "Osaka", "Nagoya", "Fukuoka"]]
                if japan_dfs:
                    combined = pd.concat(japan_dfs)
                    combined["日期"] = pd.to_datetime(combined["日期"])

                    japan_daily = combined.groupby("日期", as_index=False).agg({
                        "IJOOZ 仓库库存（单位）": "sum",
                        "外部冷库库存（整柜数）": "sum",
                        "使用柜数量": "sum",
                        "运输中（单位）": "sum",
                        "PO Placed（单位）": "sum",
                        "daily_usage": "sum",
                        "当天使用的货柜 PO": lambda x: ', '.join(filter(None, map(str, x)))
                    })
                    japan_daily["日期"] = japan_daily["日期"].dt.strftime("%Y-%m-%d")

                    japan_path = os.path.join(tmpdirname, "Japan_Daily_Inventory.xlsx")
                    with pd.ExcelWriter(japan_path, engine="openpyxl") as writer:
                        japan_daily.to_excel(writer, index=False, sheet_name="Japan Daily Inventory")
                    excel_paths.append(japan_path)

            zip_output = BytesIO()
            with zipfile.ZipFile(zip_output, "w") as zipf:
                for path in excel_paths:
                    arcname = os.path.basename(path)
                    zipf.write(path, arcname=arcname)
            zip_output.seek(0)
            return zip_output

    # === 执行逻辑 ===
    try:
        today_str = datetime.date.today().strftime('%Y-%m-%d')
        with st.spinner("模拟进行中，请稍候..."):

            if warehouse_name == "全部仓库":
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
