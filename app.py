import streamlit as st
import pandas as pd
import io
import re
from datetime import timedelta
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ================= 1. 网页基础设置 & 究极 UI 美化 =================
st.set_page_config(page_title="教师课时管理系统", page_icon="🎓", layout="wide")

st.markdown("""
<style>
    [data-testid="stHeader"] { background-color: transparent !important; }
    .stApp { background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%); font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; }
    [data-testid="block-container"] { padding-top: 0rem !important; padding-left: 2rem !important; padding-right: 2rem !important; max-width: 98% !important; }
    .custom-header {
        background: linear-gradient(90deg, #1e3c72 0%, #2a5298 100%); color: white; padding: 1.5rem; border-radius: 0 0 15px 15px; 
        text-align: center; box-shadow: 0 6px 20px rgba(30, 60, 114, 0.2); margin-bottom: 25px; margin-top: -15px; 
        font-size: 28px; font-weight: 900; letter-spacing: 2px;
    }
    [data-testid="stSidebar"] { background-color: rgba(255, 255, 255, 0.9) !important; border-right: 1px solid rgba(255,255,255,0.5); box-shadow: 2px 0 15px rgba(0,0,0,0.05); }
    div.stButton > button {
        white-space: nowrap !important; font-size: 13px !important; padding: 4px 12px !important; min-height: 32px !important; 
        height: 32px !important; width: 100% !important; background-color: rgba(255, 255, 255, 0.95) !important;      
        color: #334155 !important; border: 1px solid #cbd5e1 !important; border-radius: 20px !important; box-shadow: 0 2px 4px rgba(0,0,0,0.02) !important;
        transition: all 0.3s ease !important; 
    }
    div.stButton > button:hover {
        background: linear-gradient(90deg, #2a5298 0%, #1e3c72 100%) !important; color: white !important; border-color: #1e3c72 !important;
        transform: translateY(-2px); box-shadow: 0 6px 12px rgba(42, 82, 152, 0.25) !important;
    }
    div[data-testid="stDownloadButton"] > button {
        background: linear-gradient(135deg, #f6d365 0%, #fda085 100%) !important; color: white !important; border: none !important; letter-spacing: 1px;
        border-radius: 8px !important; box-shadow: 0 4px 15px rgba(253, 160, 133, 0.4) !important; font-size: 14px !important;
    }
    div[data-testid="stDownloadButton"] > button:hover { background: linear-gradient(135deg, #fda085 0%, #f6d365 100%) !important; transform: scale(1.02); }
    .row-title { font-size: 14px; font-weight: bold; color: #1e293b; text-align: right; padding-top: 6px; padding-right: 12px; white-space: nowrap; }
    [data-testid="stDataFrame"] { background: white; padding: 10px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.03); border: 1px solid #f1f5f9; }
    [data-testid="column"] { padding: 0 5px !important; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="custom-header">🎓 教师排课智能读取与精准统计系统</div>', unsafe_allow_html=True)

if 'all_sheets' not in st.session_state: st.session_state['all_sheets'] = None
if 'current_sheet' not in st.session_state: st.session_state['current_sheet'] = None
if 'global_mode' not in st.session_state: st.session_state['global_mode'] = False
if 'teacher_mode' not in st.session_state: st.session_state['teacher_mode'] = False
if 'search_teacher' not in st.session_state: st.session_state['search_teacher'] = ""

# ================= 辅助函数 =================
def col2num(col_str):
    expn = 0; col_num = 0
    for char in reversed(str(col_str).upper().strip()):
        if 'A' <= char <= 'Z':
            col_num += (ord(char) - ord('A') + 1) * (26 ** expn)
            expn += 1
    return col_num - 1 if col_num > 0 else 0

# (普通统计报表导出)
def convert_df_to_excel_pro(df, sheet_name, title):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        write_index = True if df.index.name or (isinstance(df.index, pd.MultiIndex)) else False
        df.to_excel(writer, sheet_name=sheet_name, startrow=2, index=write_index)
        worksheet = writer.sheets[sheet_name]
        
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        header_fill = PatternFill(start_color="1E3C72", end_color="1E3C72", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        num_index_cols = len(df.index.names) if write_index else 0
        max_col = len(df.columns) + num_index_cols
        
        cell = worksheet.cell(row=1, column=1, value=title)
        cell.font = Font(size=18, bold=True, color="000000")
        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
        cell.alignment = center_align
        worksheet.row_dimensions[1].height = 40 
        
        worksheet.row_dimensions[3].height = 30
        for col_idx in range(1, max_col + 1):
            c = worksheet.cell(row=3, column=col_idx)
            c.fill = header_fill; c.font = header_font
            c.alignment = center_align; c.border = thin_border
            
        for r_idx in range(4, len(df) + 4):
            worksheet.row_dimensions[r_idx].height = 30 
            for c_idx in range(1, max_col + 1):
                c = worksheet.cell(row=r_idx, column=c_idx)
                c.alignment = center_align; c.border = thin_border
                if c_idx <= num_index_cols: c.font = Font(bold=True) 
                    
        for i in range(1, max_col + 1):
            if i <= num_index_cols: worksheet.column_dimensions[get_column_letter(i)].width = 14 
            else: worksheet.column_dimensions[get_column_letter(i)].width = 18 

    return output.getvalue()

# 【全新黑科技】：多周纵向课表专用导出引擎
def convert_weekly_dfs_to_excel(weekly_dfs, sheet_name, main_title):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        workbook = writer.book
        worksheet = workbook.create_sheet(sheet_name)
        writer.sheets[sheet_name] = worksheet
        
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        header_fill = PatternFill(start_color="1E3C72", end_color="1E3C72", fill_type="solid")
        week_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid") # 浅蓝色周次标题
        header_font = Font(color="FFFFFF", bold=True, size=11)
        title_font = Font(size=18, bold=True, color="000000")
        week_font = Font(size=14, bold=True, color="1E3C72")
        center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        max_col = 2 + 7 # 2列索引(节次/时间) + 7天(周一到周日)
        
        # 渲染主标题
        cell = worksheet.cell(row=1, column=1, value=main_title)
        cell.font = title_font
        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
        cell.alignment = center_align
        worksheet.row_dimensions[1].height = 40
        
        current_row = 3
        
        # 循环垂直写入每一周的课表
        for week_title, df in weekly_dfs:
            export_df = df.reset_index()
            
            # 渲染周次标题条
            w_cell = worksheet.cell(row=current_row, column=1, value=week_title)
            w_cell.font = week_font; w_cell.fill = week_fill
            w_cell.alignment = Alignment(horizontal='left', vertical='center')
            for c_idx in range(1, max_col + 1):
                worksheet.cell(row=current_row, column=c_idx).border = thin_border
            worksheet.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=max_col)
            worksheet.row_dimensions[current_row].height = 30
            current_row += 1
            
            # 渲染表头 (星期一 到 星期日)
            worksheet.row_dimensions[current_row].height = 35
            headers = list(export_df.columns)
            for c_idx, col_name in enumerate(headers, 1):
                c = worksheet.cell(row=current_row, column=c_idx, value=col_name)
                c.fill = header_fill; c.font = header_font
                c.alignment = center_align; c.border = thin_border
            current_row += 1
            
            # 渲染具体课表数据
            for r_idx in range(len(export_df)):
                worksheet.row_dimensions[current_row].height = 45 # 留出行高放班级马甲
                for c_idx, col_name in enumerate(headers, 1):
                    val = export_df.iloc[r_idx][col_name]
                    c = worksheet.cell(row=current_row, column=c_idx, value=val)
                    c.alignment = center_align; c.border = thin_border
                    if c_idx <= 2: c.font = Font(bold=True) # 索引列加粗
                current_row += 1
            
            current_row += 1 # 周与周之间留一个空行
            
        # 设置完美的列宽
        worksheet.column_dimensions[get_column_letter(1)].width = 12
        worksheet.column_dimensions[get_column_letter(2)].width = 16
        for i in range(3, max_col + 1):
            worksheet.column_dimensions[get_column_letter(i)].width = 22
            
        if 'Sheet' in workbook.sheetnames and len(workbook.sheetnames) > 1:
            del workbook['Sheet']
            
    return output.getvalue()

# ================= 智能识别与清洗引擎 =================
def clean_excel_data(df):
    is_schedule = False
    for i in range(min(15, len(df))):
        row_str = " ".join(str(x) for x in df.iloc[i].values)
        if "星期" in row_str or re.search(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}', row_str):
            is_schedule = True; break
            
    if is_schedule:
        new_cols = [get_column_letter(i+1) for i in range(len(df.columns))]
        df.columns = new_cols
        return df.dropna(how='all', axis=0)
    else:
        header_idx = -1
        for i in range(min(10, len(df))):
            if any(k in str(df.iloc[i].values) for k in ["姓名", "科目", "类别", "课数"]):
                header_idx = i; break
        if header_idx != -1:
            raw_cols = df.iloc[header_idx].tolist()
            df = df.iloc[header_idx + 1:].reset_index(drop=True)
        else:
            raw_cols = df.columns.tolist() 
            
        new_cols = []
        for idx, col in enumerate(raw_cols):
            letter = get_column_letter(idx + 1)
            c = str(col).strip()
            if pd.isna(col) or c.lower() in ['nan', '', 'unnamed'] or 'unnamed' in c.lower(): c = letter
            else: c = f"{letter}_{c}"
            base = c; counter = 1
            while c in new_cols: c = f"{base}_{counter}"; counter += 1
            new_cols.append(c)
        df.columns = new_cols
        return df.dropna(how='all', axis=1).dropna(how='all', axis=0)

# ================= 核心统计算法库 =================
def parse_class_string(val_str):
    val_str = str(val_str).replace(" ", "") 
    ignore = ['0', '0.0', 'nan', 'none', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日', '体育', '班会', '国学', '美术', '音乐', '大扫除', '休息']
    if not val_str or val_str.lower() in ignore or re.search(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}', val_str) or re.search(r'^第[一二三四五六七八九十]+周', val_str):
        return None
        
    count = 1.0
    m_num = re.search(r'(\d+(?:\.\d+)?)$', val_str)
    if m_num:
        if m_num.start() == 0: return None
        count = float(m_num.group(1))
        val_str = val_str[:m_num.start()] 
        
    match = re.match(r'^([\u4e00-\u9fa5a-zA-Z]+?)(高[一二三]|初[一二三]|小[一二三四五六])(.*)$', val_str)
    if match: return {'教师姓名': match.group(1), '课程类别': match.group(2) + match.group(3), '课时数': count}
        
    known_types = ['早自', '正大', '正小', '晚自', '自大', '自小', '辅导', '正课', '早读', '晚修']
    for kt in known_types:
        if val_str.endswith(kt): return {'教师姓名': val_str[:-len(kt)], '课程类别': kt, '课时数': count}
            
    if len(val_str) >= 2: return {'教师姓名': val_str, '课程类别': '常规课', '课时数': count}
    return None

# ================= 侧边栏与核心控制台 =================
st.sidebar.markdown('<div style="text-align:center; padding-bottom:10px;"><h2 style="color:#1e3c72; font-weight:bold;">📁 数据控制台</h2></div>', unsafe_allow_html=True)
uploaded_file = st.sidebar.file_uploader("请拖拽或点击上传 Excel (.xlsm/xlsx)", type=["xlsm", "xlsx"])

if uploaded_file is not None and st.session_state['all_sheets'] is None:
    try:
        with st.spinner('正在全校数据中进行无限深潜解析，请稍候...'):
            raw_sheets = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')
            clean_sheets = {}
            for sheet_name, df in raw_sheets.items(): clean_sheets[sheet_name] = clean_excel_data(df)
            st.session_state['all_sheets'] = clean_sheets
            st.session_state['current_sheet'] = list(clean_sheets.keys())[0]
            st.sidebar.success("✅ 文件解析成功！深潜模式已启动。")
    except Exception as e:
        st.error(f"严重错误: {e}")

if st.session_state['all_sheets'] is not None:
    valid_classes = [s for s in st.session_state['all_sheets'].keys() if not any(kw in s for kw in ['总表', '分表', '汇总'])]
    
    st.sidebar.markdown("---")
    st.sidebar.markdown("<br><b>📍 全局课表坐标设置 (输入字母)</b>", unsafe_allow_html=True)
    col_c1, col_c2 = st.sidebar.columns(2)
    with col_c1: t_period = st.text_input("【节次】列", value="V")
    with col_c2: t_time = st.text_input("【时间】列", value="W")
    col_c3, col_c4 = st.sidebar.columns(2)
    with col_c3: t_start = st.text_input("【排课起】", value="Y")
    with col_c4: t_end = st.text_input("【排课止】", value="AE")
    
    st.session_state['t_period'] = t_period
    st.session_state['t_time'] = t_time
    st.session_state['t_start'] = t_start
    st.session_state['t_end'] = t_end

    st.sidebar.markdown("---")
    st.sidebar.markdown('<h4 style="color:#2a5298;">🧑‍🏫 个人课表生成器</h4>', unsafe_allow_html=True)
    search_input = st.sidebar.text_input("🔍 输入教师姓名：", placeholder="例如：聂俊生")

    if st.sidebar.button("🔎 生成纵向多周课表", use_container_width=True):
        if search_input.strip() == "": st.sidebar.warning("请先输入教师姓名！")
        else:
            st.session_state['teacher_mode'] = True
            st.session_state['global_mode'] = False
            st.session_state['search_teacher'] = search_input.strip()

    st.sidebar.markdown("---")
    st.sidebar.markdown('<h4 style="color:#2a5298;">🌐 批量课时统计</h4>', unsafe_allow_html=True)
    scope = st.sidebar.radio("📌 统计范围选择", ["所有班级 (全校)", "按年级多选", "自定义勾选班级"])
    
    target_classes = []
    if scope == "所有班级 (全校)": target_classes = valid_classes
    elif scope == "按年级多选":
        grades = st.sidebar.multiselect("挑选年级", ["高一", "高二", "高三"], default=["高三"])
        target_classes = [c for c in valid_classes if any(g in c for g in grades)]
    else:
        target_classes = st.sidebar.multiselect("勾选具体的班级", valid_classes, default=valid_classes[:2])

    g_dates = st.sidebar.date_input("🗓️ 选填：限定时间段 (留空则统计所有日期)", [])
    
    if st.sidebar.button("🚀 一键生成全局汇总", use_container_width=True, type="primary"):
        if not target_classes: st.sidebar.error("当前没有选定任何班级！")
        else:
            st.session_state['global_mode'] = True
            st.session_state['teacher_mode'] = False
            st.session_state['g_dates'] = g_dates
            st.session_state['g_targets'] = target_classes
            st.session_state['g_scope'] = scope

# ================= 动态顶部导航 =================
if st.session_state['all_sheets'] is not None:
    all_sheet_names = list(st.session_state['all_sheets'].keys())
    directory_data = {
        "总表 & 汇总": [], "高一年级": [], "高二年级": [], 
        "高三年级": [], "其他表单": []
    }
    for name in all_sheet_names:
        if "总" in name or "分表" in name or "汇总" in name: directory_data["总表 & 汇总"].append(name)
        elif "高一" in name: directory_data["高一年级"].append(name)
        elif "高二" in name: directory_data["高二年级"].append(name)
        elif "高三" in name: directory_data["高三年级"].append(name)
        else: directory_data["其他表单"].append(name)

    st.write("")
    for category, buttons in directory_data.items():
        if not buttons: continue 
        empty_space = 10 - len(buttons) if len(buttons) < 10 else 1
        cols = st.columns([1.2] + [1] * len(buttons) + [empty_space]) 
        with cols[0]: st.markdown(f'<div class="row-title">{category} :</div>', unsafe_allow_html=True)
        for i, btn_name in enumerate(buttons):
            with cols[i+1]:
                if st.button(btn_name, key=f"nav_{btn_name}"):
                    st.session_state['current_sheet'] = btn_name
                    st.session_state['global_mode'] = False 
                    st.session_state['teacher_mode'] = False
    st.markdown("<hr style='margin: 15px 0px; border: none; border-top: 1px dashed #cbd5e1;'>", unsafe_allow_html=True)

    # ================= 核心视图分支 =================
    
    # 模式一：【教师个人纵向分周课表 (完美对齐周末补白版)】 
    if st.session_state['teacher_mode']:
        target_teacher = st.session_state['search_teacher']
        st.markdown(f"<h3 style='color:#1e3c72;'>🧑‍🏫 【{target_teacher}】全周期纵向网格课表</h3>", unsafe_allow_html=True)
        
        p_idx = col2num(st.session_state['t_period'])
        t_idx = col2num(st.session_state['t_time'])
        start_idx = col2num(st.session_state['t_start'])
        end_idx = col2num(st.session_state['t_end'])
        
        teacher_schedule = []
        valid_classes_to_search = [s for s in st.session_state['all_sheets'].keys() if not any(kw in s for kw in ['总表', '分表', '汇总'])]
        
        with st.spinner('正在提取并为您进行纵向维度折叠...'):
            for s_name in valid_classes_to_search:
                s_df = st.session_state['all_sheets'][s_name]
                if len(s_df.columns) <= max(p_idx, t_idx, end_idx): continue
                    
                for col_idx in range(start_idx, end_idx + 1):
                    current_date = None
                    
                    for row_idx in range(len(s_df)):
                        val_str = str(s_df.iloc[row_idx, col_idx]).strip()
                        
                        m = re.search(r'(\d{4}[-/]\d{1,2}[-/]\d{1,2})', val_str)
                        if m:
                            try: current_date = pd.to_datetime(m.group(1)).date()
                            except: pass
                            continue 
                            
                        if not current_date: continue
                        
                        if target_teacher in val_str:
                            period = str(s_df.iloc[row_idx, p_idx]).strip()
                            time_val = str(s_df.iloc[row_idx, t_idx]).strip()
                            
                            if period.lower() in ['nan', 'none', '0', '0.0']: period = ''
                            if time_val.lower() in ['nan', 'none', '0', '0.0']: time_val = ''
                            time_val = re.sub(r'[-—~]+', '-', time_val)
                            
                            course_name = f"【{s_name}】\n{val_str}" 
                            
                            teacher_schedule.append({
                                '日期': pd.to_datetime(current_date),
                                '节次/分类': period,
                                '时间/参数': time_val,
                                '课程内容': course_name,
                                '排序辅助': time_val if time_val else "99:99" 
                            })
                            
        if teacher_schedule:
            ts_df = pd.DataFrame(teacher_schedule)
            
            # 【黑科技】：精准计算每节课属于哪一周的周一，实现数据分箱
            ts_df['周起始日'] = ts_df['日期'] - pd.to_timedelta(ts_df['日期'].dt.dayofweek, unit='D')
            
            weekly_dfs = []
            weekdays_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五', 5: '六', 6: '日'}
            
            # 拿到所有出现过的“周一”，按照时间先后顺序排序
            unique_weeks = sorted(ts_df['周起始日'].unique())
            
            for idx, monday in enumerate(unique_weeks):
                sunday = monday + pd.Timedelta(days=6)
                week_title = f"第 {idx + 1} 周 ({monday.strftime('%m月%d日')} - {sunday.strftime('%m月%d日')})"
                
                # 强制生成该周完整的 7 天表头 (无论老师这天有没有课，坑位都必须占着)
                standard_cols = []
                for i in range(7):
                    day_date = monday + pd.Timedelta(days=i)
                    col_name = f"{day_date.strftime('%Y-%m-%d')}\n星期{weekdays_map[i]}"
                    standard_cols.append(col_name)
                
                # 提取这周的所有排课
                week_data = ts_df[ts_df['周起始日'] == monday].copy()
                week_data['日期排版'] = week_data['日期'].apply(lambda d: f"{d.strftime('%Y-%m-%d')}\n星期{weekdays_map[d.dayofweek]}")
                
                # 班级合并处理
                week_data = week_data.groupby(['节次/分类', '时间/参数', '排序辅助', '日期排版'])['课程内容'].apply(lambda x: '\n\n'.join(x.unique())).reset_index()
                
                # 转化为网格
                grid_df = pd.pivot_table(
                    week_data,
                    values='课程内容',
                    index=['排序辅助', '节次/分类', '时间/参数'], 
                    columns='日期排版',
                    aggfunc=lambda x: '\n\n'.join(x.dropna().unique())
                )
                
                # 核心：重塑列名，强行塞入标准 7 天，没有课的地方就是空字符串！
                grid_df = grid_df.reindex(columns=standard_cols, fill_value='')
                
                # 扫尾工作
                grid_df = grid_df.sort_index(level='排序辅助')
                grid_df = grid_df.reset_index(level='排序辅助', drop=True)
                grid_df.index.names = ['节次', '时间/姓名'] 
                grid_df.columns.name = None
                
                weekly_dfs.append((week_title, grid_df))
            
            st.success(f"🎉 生成成功！完美实现了纵向分周，周末也整齐对齐！")
            
            # 在网页中垂直显示每一周的课表
            for title, df in weekly_dfs:
                st.markdown(f"<h5 style='color:#2a5298; margin-top:20px; border-left: 4px solid #f6d365; padding-left: 10px;'>📅 {title}</h5>", unsafe_allow_html=True)
                st.dataframe(df, use_container_width=True)
            
            formal_title = f"【{target_teacher}】全周期专属课表"
            excel_data = convert_weekly_dfs_to_excel(weekly_dfs, sheet_name="个人课表", main_title=formal_title)
            st.download_button(
                label=f"⬇️ 下载纵向多周《{target_teacher}专属课表》",
                data=excel_data, file_name=f"{target_teacher}_分周专属课表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning(f"😔 没有找到包含【{target_teacher}】的排课信息！")

    # 模式二：【全局汇总视图】 (保持不变)
    elif st.session_state['global_mode']:
        f_dates = st.session_state['g_dates']
        targets = st.session_state['g_targets']
        start_idx = col2num(st.session_state['t_start'])
        end_idx = col2num(st.session_state['t_end'])
        
        report_title_prefix = "全校" if st.session_state['g_scope'] == "所有班级 (全校)" else "选中班级"
        sub_title = f"📅 ({f_dates[0]} 至 {f_dates[1]})" if len(f_dates)==2 else (f"📅 ({f_dates[0]})" if len(f_dates)==1 else "📅 (全学期汇总)")
        st.markdown(f"<h3 style='color:#1e3c72;'>🌐 【{report_title_prefix}】课时总汇 {sub_title}</h3>", unsafe_allow_html=True)
        
        all_records = []
        with st.spinner('正在执行全盘数据深度拉取...'):
            for s_name in targets:
                if s_name not in st.session_state['all_sheets']: continue
                s_df = st.session_state['all_sheets'][s_name]
                if len(s_df.columns) <= max(start_idx, end_idx): continue
                
                for col_idx in range(start_idx, end_idx + 1):
                    current_date = None
                    for row_idx in range(len(s_df)):
                        val_str = str(s_df.iloc[row_idx, col_idx]).strip()
                        m = re.search(r'(\d{4}[-/]\d{1,2}[-/]\d{1,2})', val_str)
                        if m:
                            try: current_date = pd.to_datetime(m.group(1)).date()
                            except: pass
                            continue
                        
                        if current_date:
                            if len(f_dates) == 2:
                                if not (f_dates[0] <= current_date <= f_dates[1]): continue
                            elif len(f_dates) == 1:
                                if current_date != f_dates[0]: continue
                                
                            parsed = parse_class_string(val_str)
                            if parsed:
                                parsed['来源班级'] = s_name
                                parsed['来源日期'] = str(current_date)
                                all_records.append(parsed)
                            
        if all_records:
            stat_df = pd.DataFrame(all_records)
            pivot_df = pd.pivot_table(stat_df, values='课时数', index='教师姓名', columns='课程类别', aggfunc='sum', fill_value=0)
            pivot_df['总计'] = pivot_df.sum(axis=1)
            
            st.success(f"🎉 统计完毕！共 {len(stat_df['教师姓名'].unique())} 位老师，总计 {stat_df['课时数'].sum()} 节。")
            st.dataframe(pivot_df, use_container_width=True)
            
            formal_title = f"【{report_title_prefix}汇总】课时报表 {sub_title.replace('📅 ', '')}"
            excel_data = convert_df_to_excel_pro(pivot_df, sheet_name="数据汇总", title=formal_title)
            st.download_button(
                label=f"⬇️ 导出《{report_title_prefix}汇报表格》为 Excel",
                data=excel_data, file_name=f"{report_title_prefix}课时报表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            with st.expander("🔍 查看抓取底层明细 (用于排错)"): st.dataframe(stat_df)
        else:
            st.warning("⚠️ 在指定的范围中，未抓取到有效课时！")
            
    # 模式三：【单班级管理视图】 (保持不变)
    else:
        current = st.session_state['current_sheet']
        st.markdown(f"<h4 style='color:#1e3c72;'>👁️ 当前查看 : 【 {current} 】</h4>", unsafe_allow_html=True)
        
        df_current = st.session_state['all_sheets'][current].copy()
        display_df = df_current.astype(str).replace({' 00:00:00': ''}, regex=True).replace({'nan': '', 'None': ''})
        st.dataframe(display_df, use_container_width=True, height=350)

        st.markdown("---")
        tab1, tab2 = st.tabs(["📏 【排课表】按列字母锁定统计", "📊 【明细表】手动选列统计"])
        
        with tab1:
            start_idx = col2num(st.session_state['t_start'])
            end_idx = col2num(st.session_state['t_end'])
            
            if start_idx <= end_idx and len(display_df.columns) > end_idx:
                all_dates_in_range = set()
                for col_idx in range(start_idx, end_idx + 1):
                    for row_idx in range(len(display_df)):
                        val = display_df.iloc[row_idx, col_idx]
                        m = re.search(r'(\d{4}[-/]\d{1,2}[-/]\d{1,2})', str(val).strip())
                        if m:
                            try: all_dates_in_range.add(pd.to_datetime(m.group(1)).date())
                            except: pass
                
                if all_dates_in_range:
                    min_d, max_d = min(all_dates_in_range), max(all_dates_in_range)
                    date_range = st.date_input(f"🗓️ 选择提取区间 (默认全选)：", [min_d, max_d])
                    
                    if st.button("🚀 开始本班精准提取", type="primary"):
                        if len(date_range) >= 1:
                            f_start = date_range[0]
                            f_end = date_range[1] if len(date_range) == 2 else date_range[0]
                            
                            records = []
                            for col_idx in range(start_idx, end_idx + 1):
                                current_date = None
                                for row_idx in range(len(display_df)):
                                    val_str = str(display_df.iloc[row_idx, col_idx]).strip()
                                    m = re.search(r'(\d{4}[-/]\d{1,2}[-/]\d{1,2})', val_str)
                                    if m:
                                        try: current_date = pd.to_datetime(m.group(1)).date()
                                        except: pass
                                        continue
                                    
                                    if current_date and (f_start <= current_date <= f_end):
                                        parsed = parse_class_string(val_str)
                                        if parsed: records.append(parsed)
                                            
                            if records:
                                stat_df = pd.DataFrame(records)
                                pivot_df = pd.pivot_table(stat_df, values='课时数', index='教师姓名', columns='课程类别', aggfunc='sum', fill_value=0)
                                pivot_df['总计'] = pivot_df.sum(axis=1)
                                
                                st.success(f"🎉 统计完毕！【{current}】共计 {stat_df['课时数'].sum()} 节课时。")
                                st.dataframe(pivot_df, use_container_width=True)
                                
                                formal_title = f"【{current}】课时统计报表 ({f_start}至{f_end})"
                                excel_data = convert_df_to_excel_pro(pivot_df, sheet_name=current, title=formal_title)
                                st.download_button(
                                    label=f"⬇️ 导出《{current}报表》",
                                    data=excel_data, file_name=f"{current}_课时报表.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                                with st.expander("🔍 提取明细"): st.dataframe(stat_df)
                            else:
                                st.warning("未找到可识别的课时。")
                else:
                    st.warning("⚠️ 没有在指定的列范围内扫描到包含日期的行！")
            else:
                st.warning("您侧边栏填写的字母超出了表格范围，请检查！")

        with tab2:
            available_cols = list(display_df.columns)
            def guess_index(kw):
                for i, c in enumerate(available_cols):
                    if any(k in str(c) for k in kw): return i
                return 0
                
            col1, col2, col3 = st.columns(3)
            with col1: name_col = st.selectbox("👤 【姓名】列", available_cols, index=guess_index(['姓名','教师']))
            with col2: type_col = st.selectbox("🏷️ 【类别】列", available_cols, index=guess_index(['子类','类别']))
            with col3: count_col = st.selectbox("🔢 【数量】列", available_cols, index=guess_index(['课数','课时']))
                
            if st.button("📊 生成常规统计"):
                try:
                    stat_df = df_current.copy()
                    stat_df[count_col] = pd.to_numeric(stat_df[count_col], errors='coerce').fillna(0)
                    stat_df = stat_df[stat_df[name_col].notna()]
                    stat_df = stat_df[stat_df[name_col].astype(str).str.strip() != '']
                    pivot_df = pd.pivot_table(stat_df, values=count_col, index=name_col, columns=type_col, aggfunc='sum', fill_value=0)
                    pivot_df['总计'] = pivot_df.sum(axis=1)
                    st.dataframe(pivot_df, use_container_width=True)
                    
                    formal_title = f"【{current}】常规课时统计"
                    excel_data = convert_df_to_excel_pro(pivot_df, sheet_name=current, title=formal_title)
                    st.download_button(
                        label="⬇️ 导出常规报表", data=excel_data, file_name=f"{current}_常规课时.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except:
                    st.warning("无法生成，请确认选对了列名！")
else:
    st.info("👆 请先在左侧上传您的 Excel 文件！")
