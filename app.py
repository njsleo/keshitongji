import streamlit as st
import pandas as pd
import io
import re
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
    for char in reversed(col_str.upper().strip()):
        col_num += (ord(char) - ord('A') + 1) * (26 ** expn)
        expn += 1
    return col_num - 1

def convert_df_to_excel_pro(df, sheet_name, title):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        write_index = True if df.index.name or (isinstance(df.index, pd.MultiIndex)) else False
        export_df = df
        
        export_df.to_excel(writer, sheet_name=sheet_name, startrow=2, index=write_index)
        worksheet = writer.sheets[sheet_name]
        
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        header_fill = PatternFill(start_color="1E3C72", end_color="1E3C72", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        num_index_cols = len(df.index.names) if write_index else 0
        max_col = len(df.columns) + num_index_cols
        max_row = len(df) + 3 
        
        cell = worksheet.cell(row=1, column=1, value=title)
        cell.font = Font(size=18, bold=True, color="000000")
        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
        cell.alignment = center_align
        worksheet.row_dimensions[1].height = 40 
        
        worksheet.row_dimensions[3].height = 30
        for col_idx in range(1, max_col + 1):
            c = worksheet.cell(row=3, column=col_idx)
            c.fill = header_fill
            c.font = header_font
            c.alignment = center_align
            c.border = thin_border
            
        for r_idx in range(4, max_row + 1):
            # 因为带有换行的班级标签，增加行高让它不拥挤
            worksheet.row_dimensions[r_idx].height = 35 
            for c_idx in range(1, max_col + 1):
                c = worksheet.cell(row=r_idx, column=c_idx)
                c.alignment = center_align
                c.border = thin_border
                if c_idx <= num_index_cols: c.font = Font(bold=True) 
                    
        for i in range(1, max_col + 1):
            if i <= num_index_cols: worksheet.column_dimensions[get_column_letter(i)].width = 14 
            # 【优化】：为了显示班级小括号，将数据列宽度调宽至22
            else: worksheet.column_dimensions[get_column_letter(i)].width = 22

    return output.getvalue()

# ================= 智能识别与清洗引擎 =================
def clean_excel_data(df):
    is_schedule = False
    for i in range(min(5, len(df))):
        row_str = " ".join(str(x) for x in df.iloc[i].values)
        if "星期" in row_str or re.search(r'\d{4}[-/]\d{2}[-/]\d{2}', row_str):
            is_schedule = True; break
            
    if is_schedule:
        new_cols = []
        for idx, col in enumerate(df.columns):
            letter = get_column_letter(idx + 1)
            c = str(col).strip()
            if pd.isna(col) or c.lower() in ['nan', '', 'unnamed'] or 'unnamed' in c.lower():
                c = letter 
            else:
                c = f"{letter}_{c}"
            
            base = c; counter = 1
            while c in new_cols: c = f"{base}_{counter}"; counter += 1
            new_cols.append(c)
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
        with st.spinner('正在执行精准坐标解析，请稍候...'):
            raw_sheets = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')
            clean_sheets = {}
            for sheet_name, df in raw_sheets.items(): clean_sheets[sheet_name] = clean_excel_data(df)
            st.session_state['all_sheets'] = clean_sheets
            st.session_state['current_sheet'] = list(clean_sheets.keys())[0]
            st.sidebar.success("✅ 文件解析成功！已启用 Excel 原生字母坐标！")
    except Exception as e:
        st.error(f"严重错误: {e}")

if st.session_state['all_sheets'] is not None:
    valid_classes = [s for s in st.session_state['all_sheets'].keys() if not any(kw in s for kw in ['总表', '分表', '汇总'])]
    
    # 【教师专属课表查询】
    st.sidebar.markdown("---")
    st.sidebar.markdown('<h4 style="color:#2a5298;">🧑‍🏫 个人课表生成器</h4>', unsafe_allow_html=True)
    search_input = st.sidebar.text_input("🔍 输入教师姓名：", placeholder="例如：韩志然")
    
    with st.sidebar.expander("⚙️ 课表坐标设置 (默认 V, W, Y-AE)"):
        col_t1, col_t2 = st.columns(2)
        with col_t1: t_period = st.text_input("【节次】列", value="V")
        with col_t2: t_time = st.text_input("【时间】列", value="W")
        col_t3, col_t4 = st.columns(2)
        with col_t3: t_start = st.text_input("【课表起】", value="Y")
        with col_t4: t_end = st.text_input("【课表止】", value="AE")

    if st.sidebar.button("🔎 一键生成全周期网格课表", use_container_width=True):
        if search_input.strip() == "": st.sidebar.warning("请先输入教师姓名！")
        else:
            st.session_state['teacher_mode'] = True
            st.session_state['global_mode'] = False
            st.session_state['search_teacher'] = search_input.strip()
            st.session_state['t_period'] = t_period
            st.session_state['t_time'] = t_time
            st.session_state['t_start'] = t_start
            st.session_state['t_end'] = t_end

    # 【全局统计生成器】
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

    st.sidebar.markdown("<br><b>📍 课表区域坐标 (输入字母)</b>", unsafe_allow_html=True)
    col_g1, col_g2 = st.sidebar.columns(2)
    with col_g1: g_start_letter = st.text_input("起始列 (如 Y)", value="Y")
    with col_g2: g_end_letter = st.text_input("结束列 (如 AE)", value="AE")
    g_dates = st.sidebar.date_input("🗓️ 限定统计时间段", [])
    
    if st.sidebar.button("🚀 一键生成全局报表", use_container_width=True, type="primary"):
        if len(g_dates) < 1: st.sidebar.error("请先选择完整的时间段！")
        elif not target_classes: st.sidebar.error("当前没有选定任何班级！")
        else:
            st.session_state['global_mode'] = True
            st.session_state['teacher_mode'] = False
            st.session_state['g_start'] = g_start_letter
            st.session_state['g_end'] = g_end_letter
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
    
    # 模式一：【教师个人二维网格课表 (满血版)】 
    if st.session_state['teacher_mode']:
        target_teacher = st.session_state['search_teacher']
        st.markdown(f"<h3 style='color:#1e3c72;'>🧑‍🏫 【{target_teacher}】所有周次专属网格课表</h3>", unsafe_allow_html=True)
        
        p_idx = col2num(st.session_state['t_period'])
        t_idx = col2num(st.session_state['t_time'])
        start_idx = col2num(st.session_state['t_start'])
        end_idx = col2num(st.session_state['t_end'])
        
        teacher_schedule = []
        valid_classes_to_search = [s for s in st.session_state['all_sheets'].keys() if not any(kw in s for kw in ['总表', '分表', '汇总'])]
        
        with st.spinner('正在为您垂直下挖提取所有周次的数据，请稍候...'):
            for s_name in valid_classes_to_search:
                s_df = st.session_state['all_sheets'][s_name]
                
                if len(s_df.columns) <= max(p_idx, t_idx, end_idx):
                    continue
                    
                locked_cols = s_df.columns[start_idx : end_idx + 1]
                
                for col in locked_cols:
                    current_date = None
                    current_weekday = ""
                    col_index = s_df.columns.get_loc(col)
                    
                    # 【核心修复】：垂直向下逐行扫描，遇到新日期就实时更新
                    for row_idx in range(len(s_df)):
                        val_str = str(s_df.iloc[row_idx, col_index]).strip()
                        
                        # 1. 探测是否有新日期出现
                        m = re.search(r'(\d{4}[-/]\d{1,2}[-/]\d{1,2})', val_str)
                        if m:
                            try: current_date = pd.to_datetime(m.group(1)).date()
                            except: pass
                            current_weekday = "" # 重置星期
                            if "星期" in val_str: current_weekday = val_str
                            continue # 日期行本身不排课
                            
                        # 2. 探测独立的星期行
                        if "星期" in val_str and len(val_str) <= 6:
                            current_weekday = val_str
                            continue
                            
                        # 如果还没碰到有效日期，直接跳过当前行
                        if not current_date: continue
                        
                        # 3. 匹配目标教师并提取
                        if target_teacher in val_str:
                            period = str(s_df.iloc[row_idx, p_idx]).strip()
                            time_val = str(s_df.iloc[row_idx, t_idx]).strip()
                            
                            if period.lower() in ['nan', 'none']: period = ''
                            if time_val.lower() in ['nan', 'none']: time_val = ''
                            time_val = re.sub(r'[-—~]+', '-', time_val)
                            
                            # 【核心新增】：加上班级名称前缀并换行，比如：【高三1班】\n聂俊生高三正小
                            course_name = f"【{s_name}】\n{val_str}" 
                            
                            teacher_schedule.append({
                                '日期': current_date,
                                '星期': current_weekday,
                                '节次/分类': period,
                                '时间/参数': time_val,
                                '课程内容': course_name,
                                '排序辅助': time_val if time_val else "99:99" 
                            })
                            
        if teacher_schedule:
            ts_df = pd.DataFrame(teacher_schedule)
            
            def format_date(row):
                # 采用 YYYY-MM-DD 格式，天然保证列头按照时间先后顺序正确排序
                d_str = row['日期'].strftime('%Y-%m-%d')
                w_str = row['星期']
                if not w_str:
                    weekdays_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五', 5: '六', 6: '日'}
                    w_str = f"星期{weekdays_map[row['日期'].weekday()]}"
                return f"{d_str}\n{w_str}"
                
            ts_df['日期排版'] = ts_df.apply(format_date, axis=1)
            
            # 合并同一个时间同个老师在不同班的排课
            ts_df = ts_df.groupby(['节次/分类', '时间/参数', '排序辅助', '日期排版'])['课程内容'].apply(lambda x: '\n'.join(x.unique())).reset_index()
            
            grid_df = pd.pivot_table(
                ts_df,
                values='课程内容',
                index=['排序辅助', '节次/分类', '时间/参数'], 
                columns='日期排版',
                aggfunc=lambda x: '\n'.join(x.dropna().unique())
            ).fillna('')
            
            grid_df = grid_df.sort_index(level='排序辅助')
            grid_df = grid_df.reset_index(level='排序辅助', drop=True)
            
            grid_df.index.names = ['节次', '时间/姓名'] 
            grid_df.columns.name = None
            
            st.success(f"🎉 生成成功！所有周次都给您提取出来了，并且标明了班级：")
            st.dataframe(grid_df, use_container_width=True)
            
            formal_title = f"【{target_teacher}】全周期专属网格课表"
            excel_data = convert_df_to_excel_pro(grid_df, sheet_name="个人课表", title=formal_title)
            st.download_button(
                label=f"⬇️ 下载《{target_teacher}全周期网格课表》",
                data=excel_data, file_name=f"{target_teacher}_专属课表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning(f"😔 翻遍了全校数据，没有在 {st.session_state['t_start']} 到 {st.session_state['t_end']} 列找到包含【{target_teacher}】的排课信息！")

    # 模式二：【全局汇总视图】 (略，与之前保持完全一致)
    elif st.session_state['global_mode']:
        g_dates = st.session_state['g_dates']
        f_start = g_dates[0]
        f_end = g_dates[1] if len(g_dates) == 2 else g_dates[0]
        targets = st.session_state['g_targets']
        start_idx = col2num(st.session_state['g_start'])
        end_idx = col2num(st.session_state['g_end'])
        
        report_title_prefix = "全校" if st.session_state['g_scope'] == "所有班级 (全校)" else "选中班级"
        st.markdown(f"<h3 style='color:#1e3c72;'>🌐 【{report_title_prefix}】课时总汇 📅 ({f_start} 至 {f_end})</h3>", unsafe_allow_html=True)
        st.info(f"扫描 {st.session_state['g_start']} 到 {st.session_state['g_end']} 列，涉及 {len(targets)} 个班级...")
        
        all_records = []
        for s_name in targets:
            if s_name not in st.session_state['all_sheets']: continue
            s_df = st.session_state['all_sheets'][s_name]
            
            if len(s_df.columns) <= max(start_idx, end_idx): continue
            locked_cols = s_df.columns[start_idx : end_idx + 1]
            
            for col in locked_cols:
                current_date = None
                for val in s_df[col]:
                    val_str = str(val).strip()
                    m = re.search(r'(\d{4}[-/]\d{1,2}[-/]\d{1,2})', val_str)
                    if m:
                        try: current_date = pd.to_datetime(m.group(1)).date()
                        except: pass
                        continue
                    
                    if current_date and (f_start <= current_date <= f_end):
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
            
            formal_title = f"【{report_title_prefix}汇总】课时报表 ({f_start}至{f_end})"
            excel_data = convert_df_to_excel_pro(pivot_df, sheet_name="数据汇总", title=formal_title)
            st.download_button(
                label=f"⬇️ 导出《{report_title_prefix}汇报表格》为 Excel",
                data=excel_data, file_name=f"{report_title_prefix}课时报表_{f_start}至{f_end}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            with st.expander("🔍 查看抓取底层明细 (用于排错)"): st.dataframe(stat_df)
        else:
            st.warning("⚠️ 在指定的范围中，未抓取到有效课时！")
            
    # 模式三：【单班级管理视图】 (略，与之前保持完全一致)
    else:
        current = st.session_state['current_sheet']
        st.markdown(f"<h4 style='color:#1e3c72;'>👁️ 当前查看 : 【 {current} 】</h4>", unsafe_allow_html=True)
        
        df_current = st.session_state['all_sheets'][current].copy()
        display_df = df_current.astype(str).replace({' 00:00:00': ''}, regex=True).replace({'nan': '', 'None': ''})
        st.dataframe(display_df, use_container_width=True, height=350)

        st.markdown("---")
        tab1, tab2 = st.tabs(["📏 【排课表】按列字母锁定统计", "📊 【明细表】手动选列统计"])
        
        with tab1:
            col_a, col_b = st.columns(2)
            with col_a: start_choice = st.text_input("🚩 起始列 (如 Y)", value="Y")
            with col_b: end_choice = st.text_input("🏁 结束列 (如 AE)", value="AE")
                
            start_idx, end_idx = col2num(start_choice), col2num(end_choice)
            if start_idx <= end_idx and len(display_df.columns) > end_idx:
                locked_cols = display_df.columns[start_idx : end_idx + 1]
                all_dates_in_range = set()
                for col in locked_cols:
                    for val in display_df[col]:
                        m = re.search(r'(\d{4}[-/]\d{1,2}[-/]\d{1,2})', str(val).strip())
                        if m:
                            try: all_dates_in_range.add(pd.to_datetime(m.group(1)).date())
                            except: pass
                
                if all_dates_in_range:
                    min_d, max_d = min(all_dates_in_range), max(all_dates_in_range)
                    date_range = st.date_input(f"🗓️ 选择提取区间：", [min_d, max_d])
                    
                    if len(date_range) >= 1:
                        f_start = date_range[0]
                        f_end = date_range[1] if len(date_range) == 2 else date_range[0]
                        
                        if st.button("🚀 开始本班精准提取", type="primary"):
                            records = []
                            for col in locked_cols:
                                current_date = None
                                for val in display_df[col]:
                                    val_str = str(val).strip()
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
                                    data=excel_data, file_name=f"{current}_课时报表_{f_start}至{f_end}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                                with st.expander("🔍 提取明细"): st.dataframe(stat_df)
                            else:
                                st.warning("未找到可识别的课时。")
                else:
                    st.warning("⚠️ 没有在指定的列范围内扫描到包含日期的行！")
            else:
                st.warning("您填写的列超出了表格范围，或起始列大于结束列。")

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
