import streamlit as st
import pandas as pd
import io
import os
import base64
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 页面配置
st.set_page_config(
    page_title="PCF Sparkletots 幼儿观察计划管理系统",
    page_icon="👶",
    layout="wide"
)

# 自定义样式
st.markdown("""
    <style>
    .main-title {
        font-size: 24px;
        font-weight: bold;
        color: #d9534f;
        text-align: center;
        margin-bottom: 5px;
    }
    .sub-title {
        font-size: 24px;
        color: #666;
        text-align: center;
        margin-bottom: 20px;
    }
    
    /* 基础字体与 Streamlit 控件强制设为楷体 */
    input, textarea, button, select, label, p, div, span {
        font-family: "KaiTi", "楷体", "STKaiti", serif !important;
    }
    div[data-testid="stDataEditor"] *, 
    div[data-testid="stDataFrame"] *,
    th, td {
        font-family: "KaiTi", "楷体", "STKAI TI", "STKaiti", serif !important;
    }
    </style>
""", unsafe_allow_html=True)

# 侧边栏导航：页面切换
st.sidebar.title("🧭 系统导航")
page = st.sidebar.radio("选择页面", ["📝 观察计划编辑", "🖨️ 打印预览与导出中心"])

# 标题
st.markdown('<div class="main-title">PCF SPARKLETOTS PRESCHOOL</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">幼儿观察计划表（成长档案）</div>', unsafe_allow_html=True)

# 头部基本信息表单
with st.container():
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        year = st.selectbox("年份 (Year)", ["", "2026", "2027", "2028"], key="year_sel")
    with col2:
        semester = st.selectbox("学期 (Semester)", ["", "第一学期", "第二学期"], key="sem_sel")
    with col3:
        term = st.selectbox("成长档案编号", ["", "第一份", "第二份", "第三份"], key="term_sel")
    with col4:
        class_name = st.selectbox("班级 (Class)", ["", "豆豆班", "幼小班", "幼儿班"], key="class_sel")
    with col5:
        teacher_name = st.text_input("教师姓名 (Teacher)", value="", key="teacher_input")

st.markdown("---")

# 初始化数据生成函数
def create_default_row(i):
    return {
        "序号": i,
        "幼儿姓名": "",
        "计划1_日期": "",
        "计划1_类型": "作品",
        "计划1_领域": "语言",
        "计划2_日期": "",
        "计划2_类型": "照片",
        "计划2_领域": "数学",
        "计划3_日期": "",
        "计划3_类型": "轶事",
        "计划3_领域": "科学",
    }

# 初始化 Session State 数据
if "df_data" not in st.session_state:
    st.session_state.df_data = pd.DataFrame([create_default_row(i) for i in range(1, 13)])

# 动态响应人数变化的回调函数
def update_student_count():
    new_count = st.session_state.student_count_input
    current_df = st.session_state.df_data.copy()
    current_len = len(current_df)
    
    if new_count > current_len:
        new_rows = pd.DataFrame([create_default_row(i) for i in range(current_len + 1, new_count + 1)])
        st.session_state.df_data = pd.concat([current_df, new_rows], ignore_index=True)
    elif new_count < current_len:
        st.session_state.df_data = current_df.iloc[:new_count].reset_index(drop=True)

# 幼儿人数输入框（绑定回调）
col_num1, col_num2 = st.columns([1, 4])
with col_num1:
    student_count = st.number_input(
        "幼儿人数 (Count)", 
        min_value=1, 
        max_value=50, 
        value=len(st.session_state.df_data), 
        step=1, 
        key="student_count_input",
        on_change=update_student_count
    )

# 列名中文化映射
column_rename = {
    "序号": "序号",
    "幼儿姓名": "幼儿姓名",
    "计划1_日期": "观察计划1日期 (Wx/日期)",
    "计划1_类型": "观察计划1类型",
    "计划1_领域": "观察计划1领域",
    "计划2_日期": "观察计划2日期 (Wx/日期)",
    "计划2_类型": "观察计划2类型",
    "计划2_领域": "观察计划2领域",
    "计划3_日期": "观察计划3日期 (Wx/日期)",
    "计划3_类型": "观察计划3类型",
    "计划3_领域": "观察计划3领域",
}

# ----------------- 页面一：观察计划编辑 -----------------
if page == "📝 观察计划编辑":
    st.markdown("<h3>📋 幼儿观察计划表</h3>", unsafe_allow_html=True)
    st.markdown("请在下方表格中输入幼儿姓名及各项观察计划（行数由上方“幼儿人数”统一控制）。")

    type_options = ["", "作品", "照片", "视频", "轶事"]
    domain_options = ["", "听", "说", "读", "写"]

    # 使用独立的 key 避免与数据 DataFrame 变量名混淆，并及时把编辑后的结果同步回 df_data
    edited_df = st.data_editor(
        st.session_state.df_data,
        use_container_width=True,
        column_config={
            "序号": st.column_config.NumberColumn("序号", disabled=True, width=50),
            "幼儿姓名": st.column_config.TextColumn("幼儿姓名", width=100),
            "计划1_日期": st.column_config.TextColumn("观察计划(TxWx/日期)", width=200),
            "计划1_类型": st.column_config.SelectboxColumn("观察计划1类型", options=type_options, width=150),
            "计划1_领域": st.column_config.SelectboxColumn("观察计划1领域", options=domain_options, width=150),
            "计划2_日期": st.column_config.TextColumn("观察计划2(TxWx/日期)", width=200),
            "计划2_类型": st.column_config.SelectboxColumn("观察计划2类型", options=type_options, width=150),
            "计划2_领域": st.column_config.SelectboxColumn("观察计划2领域", options=domain_options, width=150),
            "计划3_日期": st.column_config.TextColumn("观察计划3(Tx  Wx/日期)", width=200),
            "计划3_类型": st.column_config.SelectboxColumn("观察计划3类型", options=type_options, width=150),
            "计划3_领域": st.column_config.SelectboxColumn("观察计划3领域", options=domain_options, width=150),
        },
        key="editor_grid"
    )

    # 保持最新编辑结果更新到状态中
    st.session_state.df_data = edited_df

    st.markdown("<br>", unsafe_allow_html=True)
    
    def get_print_html():
        current_preview_df = st.session_state.df_data.copy()
        current_preview_df.rename(columns=column_rename, inplace=True)
        
        # 自动读取同目录下的 pcflandscape.png 并转换为 Base64
        img_base64 = ""
        if os.path.exists("pcflandscape.png"):
            with open("pcflandscape.png", "rb") as f:
                encoded = base64.b64encode(f.read()).decode()
                img_base64 = f"data:image/png;base64,{encoded}"

        return f"""
        <html>
            <head>
                <title>幼儿观察计划表</title>
                <style>
                    body {{ font-family: "KaiTi", "楷体", serif; padding: 20px; position: relative; }}
                    h2, h3 {{ text-align: center; margin: 5px 0; }}
                    .meta-table {{ width: 100%; margin: 20px 0; font-size: 15px; font-weight: bold; border-collapse: collapse; }}
                    .meta-table td {{ text-align: center; padding: 5px; }}
                    table.data-table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
                    th, td {{ border: 1px solid #999; padding: 8px; text-align: center; font-size: 14px; font-family: "KaiTi", "楷体", serif; }}
                    th {{ background-color: #d9534f; color: white; }}
                    
                    /* 右上角图标样式 */
                    .top-right-icon {{
                        position: absolute;
                        top: 15px;
                        right: 20px;
                        width: 50px;
                        height: auto;
                        z-index: 9999;
                    }}
                </style>
            </head>
            <body>
                <!-- 右上角图标 -->
                <img src="{img_base64}" alt="Logo" class="top-right-icon">

                <h2>PCF SPARKLETOTS PRESCHOOL</h2>
                <h3>幼儿观察计划表（成长档案）</h3>
                
                <table class="meta-table">
                    <tr>
                        <td>年份: {year or ''}</td>
                        <td>学期: {semester or ''}</td>
                        <td>成长档案编号: {term or ''}</td>
                        <td>班级: {class_name or ''}</td>
                        <td>教师: {teacher_name or ''}</td>
                    </tr>
                </table>
                
                {current_preview_df.to_html(index=False, border=0, classes="data-table")}
                <script>
                    window.print();
                </script>
            </body>
        </html>
        """
    
    st.download_button(
        label="🖨️ 下载并打印网页版 (HTML)",
        data=get_print_html(),
        file_name=f"幼儿观察计划打印版_{year}_{semester}_{class_name}.html",
        mime="text/html",
        help="点击下载网页文件，双击打开后会自动弹出打印/另存为PDF窗口"
    )

# ----------------- 页面二：打印预览与导出中心 -----------------
elif page == "🖨️ 打印预览与导出中心":
    st.markdown("### 🖨️ 导出中心")
    st.markdown("请点击下方按钮下载 Excel 观察计划表文件：")
    st.markdown("<br>", unsafe_allow_html=True)

    preview_df = st.session_state.df_data.copy()
    preview_df.rename(columns=column_rename, inplace=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        preview_df.to_excel(writer, sheet_name=semester or "观察计划", index=False)
    
    output.seek(0)
    wb = openpyxl.load_workbook(output)
    ws = wb.active

    kaiti_font_header = Font(name="楷体", size=11, bold=True, color="FFFFFF")
    kaiti_font_body = Font(name="楷体", size=11, color="000000")
    header_fill = PatternFill(start_color="D9534F", end_color="D9534F", fill_type="solid")
    align_center = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style='thin', color='CCCCCC'), right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'), bottom=Side(style='thin', color='CCCCCC')
    )
    
    for row_idx, row in enumerate(ws.iter_rows(), start=1):
        for cell in row:
            cell.border = thin_border
            cell.alignment = align_center
            if row_idx == 1:
                cell.font = kaiti_font_header
                cell.fill = header_fill
            else:
                cell.font = kaiti_font_body
                
    for col in ws.columns:
        max_len = 0
        for cell in col:
            if cell.value is not None:
                val_str = str(cell.value)
                length = sum(2 if ord(char) > 127 else 1 for char in val_str)
                max_len = max(max_len, length)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
        
    output = io.BytesIO()
    wb.save(output)
    processed_data = output.getvalue()
    
    st.download_button(
        label="📥 点击下载 Excel 观察计划表",
        data=processed_data,
        file_name=f"幼儿观察计划_{year}_{semester}_{class_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )