import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from io import BytesIO, StringIO

st.set_page_config(page_title="志愿学时统计工具", layout="wide")
st.title("📊 志愿学时自动统计工具")
st.caption("上传或粘贴报名表 → 自动统计学时 → 生成符合规范的 Excel")

# ==================== 侧边栏配置 ====================
with st.sidebar:
    st.header("⚙️ 配置")
    hours_per_session = st.number_input("单次活动学时", min_value=1, value=2, step=1)

    st.subheader("地点归类映射")
    st.caption("格式：地点名=分类，每行一个")
    default_mapping = "西苑一号站=宿管站\n中苑八号站=宿管站\n图书馆=图书馆"
    mapping_text = st.text_area("映射关系", value=default_mapping, height=120)

    location_map = {}
    for line in mapping_text.strip().split("\n"):
        if "=" in line:
            k, v = line.split("=", 1)
            location_map[k.strip()] = v.strip()

# ==================== 数据输入 ====================
tab1, tab2 = st.tabs(["📁 上传文件", "📋 粘贴数据"])

raw_df = None

with tab1:
    f = st.file_uploader("上传 CSV / Excel", type=["csv", "xlsx", "xls"])
    if f:
        try:
            raw_df = pd.read_csv(f) if f.name.endswith(".csv") else pd.read_excel(f)
            st.success(f"读取到 {len(raw_df)} 行")
        except Exception as e:
            st.error(f"读取失败：{e}")

with tab2:
    text = st.text_area("从在线表格复制粘贴（Tab 分隔）", height=180)
    if text.strip():
        try:
            raw_df = pd.read_csv(StringIO(text), sep="\t")
            st.success(f"解析到 {len(raw_df)} 行")
        except Exception as e:
            st.error(f"解析失败：{e}")

# ==================== 数据处理 ====================
if raw_df is not None:
    st.subheader("🔍 原始数据预览")
    st.dataframe(raw_df.head(20), use_container_width=True)

    def extract_long(df):
        """自动把各种格式提取成长表：地点、姓名、学号"""
        # 情况1：标准长表
        if "姓名" in df.columns and "学号" in df.columns:
            loc_candidates = ["地点", "活动地点", "位置", "站名", "站点"]
            loc_col = next((c for c in df.columns if c in loc_candidates), None)
            if loc_col:
                res = df[[loc_col, "姓名", "学号"]].copy()
                res.columns = ["地点", "姓名", "学号"]
                res["地点"] = res["地点"].ffill()
                return res.dropna(subset=["姓名", "学号"])

        # 情况2：宽表（多个 姓名/学号 列对，如截图那种）
        name_cols = [c for c in df.columns if "姓名" in str(c)]
        if name_cols:
            records = []
            loc_col = next((c for c in df.columns if c in ["地点", "活动地点", "位置", "站名", "站点"]), None)
            for nc in name_cols:
                prefix = str(nc).replace("姓名", "")
                ic = prefix + "学号"
                if ic in df.columns:
                    tmp = df[[nc, ic]].copy()
                    tmp.columns = ["姓名", "学号"]
                    tmp["地点"] = df[loc_col].ffill() if loc_col else "未知地点"
                    tmp = tmp.dropna(subset=["姓名", "学号"])
                    records.append(tmp)
            if records:
                return pd.concat(records, ignore_index=True)
        return None

    long_df = extract_long(raw_df)

    # 自动识别失败时，手动选列
    if long_df is None:
        st.warning("未能自动识别格式，请手动选择对应列：")
        cols = list(raw_df.columns)
        c1, c2, c3 = st.columns(3)
        loc_col = c1.selectbox("地点列", cols, index=0)
        name_col = c2.selectbox("姓名列", cols, index=min(1, len(cols) - 1))
        id_col = c3.selectbox("学号列", cols, index=min(2, len(cols) - 1))
        long_df = raw_df[[loc_col, name_col, id_col]].copy()
        long_df.columns = ["地点", "姓名", "学号"]
        long_df["地点"] = long_df["地点"].ffill()
        long_df = long_df.dropna(subset=["姓名", "学号"])
    else:
        st.success(f"✅ 自动识别成功，共 {len(long_df)} 条记录")

    # 数据清洗
    long_df["姓名"] = long_df["姓名"].astype(str).str.strip()
    long_df["学号"] = long_df["学号"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    long_df["地点"] = long_df["地点"].astype(str).str.strip()
    long_df = long_df[(long_df["姓名"] != "") & (long_df["学号"] != "")]

    st.subheader("✏️ 数据修正（可直接双击编辑）")
    edited = st.data_editor(long_df, num_rows="dynamic", use_container_width=True)

    # 映射活动分类
    edited["活动分类"] = edited["地点"].apply(lambda x: location_map.get(x, x))

    # 统计
    grouped = edited.groupby(["活动分类", "姓名", "学号"]).size().reset_index(name="次数")
    grouped["学时"] = grouped["次数"] * hours_per_session

    st.subheader("📊 统计结果")
    st.dataframe(grouped, use_container_width=True)

    # 按（分类 + 学时）分组
    groups = {}
    for _, row in grouped.iterrows():
        key = (row["活动分类"], int(row["学时"]))
        groups.setdefault(key, []).append({"姓名": row["姓名"], "学号": row["学号"]})

    # ==================== 生成规范 Excel ====================
    def make_excel(items):
        wb = Workbook()
        ws = wb.active

        # 表头：姓名在前，学号在后
        ws["A1"], ws["B1"] = "姓名", "学号"
        yellow = PatternFill(patternType="solid", start_color="FFFF00", end_color="FFFF00")
        center = Alignment(horizontal="center", vertical="center")
        for c in ["A1", "B1"]:
            ws[c].font = Font(name="宋体", size=12)
            ws[c].alignment = center
            ws[c].fill = yellow

        # 数据：宋体 12
        for i, it in enumerate(items, start=2):
            ws.cell(row=i, column=1, value=it["姓名"]).font = Font(name="宋体", size=12)
            ws.cell(row=i, column=2, value=it["学号"]).font = Font(name="宋体", size=12)

        # 全边框
        border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin")
        )
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=2):
            for cell in row:
                cell.border = border

        # 列宽，保证学号全显示
        ws.column_dimensions["A"].width = 15
        ws.column_dimensions["B"].width = 20

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    st.subheader("📥 下载规范表格")
    for (cat, hrs), items in sorted(groups.items()):
        fname = f"{cat}志愿活动{hrs}学时.xlsx"
        st.download_button(
            label=f"下载：{fname}（{len(items)} 人）",
            data=make_excel(items),
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_{cat}_{hrs}"
        )