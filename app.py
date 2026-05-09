import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
from io import BytesIO
import base64

# 页面配置
st.set_page_config(
    page_title="Excel工单数据分析平台",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 强制设置中文字体，解决方框问题
plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# ==================== 全局状态管理 ====================
if 'df_original' not in st.session_state:
    st.session_state.df_original = None
if 'df_current' not in st.session_state:
    st.session_state.df_current = None
if 'filter_history' not in st.session_state:
    st.session_state.filter_history = []
if 'current_filter' not in st.session_state:
    st.session_state.current_filter = "无"

# ==================== 数据预处理函数 ====================
@st.cache_data
def preprocess_data(df):
    """数据预处理：区县归类、状态标记、超时判定"""
    df = df.copy()

    # 区县关键字映射（I列为空时，用B列匹配）
    district_keywords = {
        "大新": ["大新"],
        "扶绥": ["扶绥"],
        "江州": ["江州"],
        "龙州": ["龙州"],
        "宁明": ["宁明"],
        "凭祥": ["凭祥"],
        "天等": ["天等"]
    }

    # 确保列名存在（兼容大小写和常见变体）
    col_map = {}
    for col in df.columns:
        col_upper = str(col).upper().replace('列', '')
        if col_upper in ['B', 'E', 'I', 'N', 'AC', 'AE', 'AR', 'CT']:
            col_map[col_upper] = col

    # 区县归类
    def classify_district(row):
        i_col = col_map.get('I', None)
        b_col = col_map.get('B', None)

        if i_col and pd.notna(row.get(i_col)):
            val = str(row[i_col]).strip()
            if val in district_keywords:
                return val

        if b_col:
            b_val = str(row.get(b_col, ''))
            for district, keywords in district_keywords.items():
                if any(kw in b_val for kw in keywords):
                    return district
        return "未归类"

    df['区县'] = df.apply(classify_district, axis=1)

    # 回单状态（N列有数据=已回单，无数据=未回单）
    n_col = col_map.get('N', None)
    if n_col:
        df['回单状态'] = df[n_col].apply(lambda x: '已回单' if pd.notna(x) and str(x).strip() != '' else '未回单')
    else:
        df['回单状态'] = '未知'

    # 超时判定（AR列>8为超时）
    ar_col = col_map.get('AR', None)
    if ar_col:
        df['超时状态'] = pd.to_numeric(df[ar_col], errors='coerce').apply(
            lambda x: '超时' if pd.notna(x) and x > 8 else ('未超时' if pd.notna(x) else '未知')
        )
    else:
        df['超时状态'] = '未知'

    # 保存列映射供后续使用
    df.attrs['col_map'] = col_map
    return df

def calculate_work_hours(start_time, end_time):
    """计算剔除夜间的实际工作时长（20:00-7:00剔除）"""
    if pd.isna(start_time) or pd.isna(end_time):
        return np.nan

    try:
        if isinstance(start_time, str):
            start_time = pd.to_datetime(start_time)
        if isinstance(end_time, str):
            end_time = pd.to_datetime(end_time)
    except:
        return np.nan

    if end_time <= start_time:
        return 0

    total_hours = 0
    current = start_time

    while current < end_time:
        # 当天剩余时间
        day_end = current.replace(hour=23, minute=59, second=59)
        segment_end = min(end_time, day_end)

        # 计算当天有效时间（7:00-20:00）
        day_start_valid = current.replace(hour=7, minute=0, second=0)
        day_end_valid = current.replace(hour=20, minute=0, second=0)

        effective_start = max(current, day_start_valid)
        effective_end = min(segment_end, day_end_valid)

        if effective_end > effective_start:
            total_hours += (effective_end - effective_start).total_seconds() / 3600

        # 进入下一天
        current = (current + pd.Timedelta(days=1)).replace(hour=0, minute=0, second=0)

    return total_hours

# ==================== 筛选功能 ====================
def apply_filter(df, filter_name):
    """应用筛选条件"""
    if df is None:
        return None

    col_map = df.attrs.get('col_map', {})
    ae_col = col_map.get('AE')
    i_col = col_map.get('I')
    ct_col = col_map.get('CT')
    ac_col = col_map.get('AC')

    districts = ["江州", "天等", "凭祥", "宁明", "龙州", "扶绥", "大新"]

    if filter_name in districts:
        return df[df['区县'] == filter_name]

    elif filter_name == "综合":
        return df[df['区县'].isin(districts)]

    elif filter_name == "无线超时工单":
        mask = pd.Series([True] * len(df))

        if ae_col:
            mask &= df[ae_col].astype(str).str.contains('无线接入网', na=False)
        if ct_col:
            mask &= df[ct_col].astype(str).str.strip().str.lower() == '否'
        if ac_col:
            mask &= df[ac_col].isna() | (df[ac_col].astype(str).str.strip() == '')

        # 时间差筛选（简化版：如果存在时间列则计算）
        # 实际项目中请替换为真实的时间列
        return df[mask]

    return df

# ==================== 统计面板 ====================
def generate_statistics(df):
    """生成区县统计面板"""
    if df is None or df.empty:
        return pd.DataFrame()

    districts = ["大新", "扶绥", "江州", "龙州", "宁明", "凭祥", "天等"]
    stats = []

    for district in districts:
        ddf = df[df['区县'] == district]
        if len(ddf) == 0:
            continue

        total = len(ddf)
        done = len(ddf[ddf['回单状态'] == '已回单'])
        undone = len(ddf[ddf['回单状态'] == '未回单'])
        overtime = len(ddf[ddf['超时状态'] == '超时'])
        not_overtime = len(ddf[ddf['超时状态'] == '未超时'])

        # 合格率 = 未超时数 / (已回单数)  或根据你的业务逻辑调整
        qualified = not_overtime
        rate = f"{qualified/max(total,1)*100:.1f}%"

        stats.append({
            '县份': district,
            '工单总数': total,
            '已回单数': done,
            '未回单数': undone,
            '超时数': overtime,
            '未超时数': not_overtime,
            '合格率': rate
        })

    return pd.DataFrame(stats)

# ==================== 数据透视 ====================
def generate_pivot(df):
    """生成数据透视表"""
    if df is None or df.empty:
        return pd.DataFrame()

    col_map = df.attrs.get('col_map', {})
    e_col = col_map.get('E')
    ae_col = col_map.get('AE')

    index_cols = ['区县']
    if e_col:
        index_cols.append(e_col)

    try:
        pivot = pd.pivot_table(
            df,
            index=index_cols,
            values=df.columns[0],  # 任意列计数
            aggfunc='count',
            fill_value=0
        ).rename(columns={df.columns[0]: '工单数'})

        # 如果AE列存在，添加AE类型汇总
        if ae_col:
            ae_summary = df.groupby(index_cols)[ae_col].apply(
                lambda x: ', '.join(x.dropna().astype(str).unique())
            ).to_frame('AE类型')
            pivot = pivot.join(ae_summary, how='left')

        return pivot.reset_index()
    except:
        return pd.DataFrame()

# ==================== 图表生成 ====================
def create_charts(df):
    """生成统计图表"""
    if df is None or df.empty:
        return None, None

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))

    # 图表1：各县区工单数对比
    district_counts = df['区县'].value_counts()
    colors = plt.cm.Set3(np.linspace(0, 1, len(district_counts)))
    district_counts.plot(kind='bar', ax=ax1, color=colors)
    ax1.set_title('各县区工单数量分布', fontsize=14, pad=15)
    ax1.set_xlabel('县份', fontsize=12)
    ax1.set_ylabel('工单数', fontsize=12)
    ax1.tick_params(axis='x', rotation=45)

    # 图表2：回单状态饼图
    status_counts = df['回单状态'].value_counts()
    if len(status_counts) > 0:
        ax2.pie(status_counts.values, labels=status_counts.index, autopct='%1.1f%%', 
                startangle=90, colors=['#66b3ff', '#ff9999'])
        ax2.set_title('回单状态占比', fontsize=14, pad=15)

    plt.tight_layout()
    return fig

# ==================== 导出功能 ====================
def to_excel_download(df):
    """生成Excel下载链接"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='筛选结果')
    output.seek(0)
    return output

# ==================== 主界面 ====================
def main():
    st.title("📊 Excel工单数据分析平台")
    st.markdown("---")

    # ==================== 侧边栏：上传与筛选 ====================
    with st.sidebar:
        st.header("📁 数据导入")
        uploaded_file = st.file_uploader(
            "上传Excel文件", 
            type=['xlsx', 'xls'],
            help="支持 .xlsx 和 .xls 格式"
        )

        if uploaded_file is not None:
            with st.spinner('正在导入数据...'):
                try:
                    df_raw = pd.read_excel(uploaded_file)
                    df_processed = preprocess_data(df_raw)
                    st.session_state.df_original = df_processed
                    st.session_state.df_current = df_processed.copy()
                    st.session_state.filter_history = []
                    st.session_state.current_filter = "无"
                    st.success(f"✅ 导入成功！共 {len(df_processed)} 行数据")
                except Exception as e:
                    st.error(f"❌ 导入失败：{str(e)}")

        st.markdown("---")
        st.header("🔍 筛选功能")

        if st.session_state.df_original is not None:
            # 8种筛选按钮
            districts = ["江州", "天等", "凭祥", "宁明", "龙州", "扶绥", "大新"]

            st.subheader("按县份筛选")
            cols = st.columns(2)
            for idx, district in enumerate(districts):
                with cols[idx % 2]:
                    if st.button(district, key=f"btn_{district}", use_container_width=True):
                        st.session_state.filter_history.append(st.session_state.df_current.copy())
                        st.session_state.df_current = apply_filter(
                            st.session_state.df_original, district
                        )
                        st.session_state.current_filter = district
                        st.rerun()

            st.subheader("特殊筛选")
            if st.button("🏢 综合（全部县份）", use_container_width=True):
                st.session_state.filter_history.append(st.session_state.df_current.copy())
                st.session_state.df_current = apply_filter(
                    st.session_state.df_original, "综合"
                )
                st.session_state.current_filter = "综合"
                st.rerun()

            if st.button("📡 无线超时工单", use_container_width=True):
                st.session_state.filter_history.append(st.session_state.df_current.copy())
                st.session_state.df_current = apply_filter(
                    st.session_state.df_original, "无线超时工单"
                )
                st.session_state.current_filter = "无线超时工单"
                st.rerun()

            st.markdown("---")

            # 一键取消筛选
            if st.button("🔄 取消筛选（显示全部）", type="primary", use_container_width=True):
                if st.session_state.filter_history:
                    st.session_state.df_current = st.session_state.df_original.copy()
                    st.session_state.filter_history = []
                    st.session_state.current_filter = "无"
                    st.rerun()
                else:
                    st.info("当前已是全部数据")

            st.markdown("---")
            st.caption(f"当前筛选：{st.session_state.current_filter}")
            if st.session_state.df_current is not None:
                st.caption(f"显示行数：{len(st.session_state.df_current)}")

    # ==================== 主内容区 ====================
    if st.session_state.df_current is not None:
        df = st.session_state.df_current

        # 顶部统计卡片
        st.subheader("📈 核心指标")
        metric_cols = st.columns(4)
        with metric_cols[0]:
            st.metric("总工单数", len(df))
        with metric_cols[1]:
            st.metric("已回单", len(df[df['回单状态']=='已回单']))
        with metric_cols[2]:
            st.metric("未回单", len(df[df['回单状态']=='未回单']))
        with metric_cols[3]:
            st.metric("超时工单", len(df[df['超时状态']=='超时']))

        st.markdown("---")

        # 标签页：数据表格 | 统计报表 | 数据透视 | 图表
        tab1, tab2, tab3, tab4 = st.tabs(["📋 数据表格", "📊 统计报表", "🔄 数据透视", "📉 图表分析"])

        with tab1:
            st.dataframe(
                df,
                use_container_width=True,
                height=500,
                hide_index=True
            )
            # 导出按钮
            excel_data = to_excel_download(df)
            st.download_button(
                label="📥 下载当前筛选结果为Excel",
                data=excel_data,
                file_name=f"筛选结果_{st.session_state.current_filter}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with tab2:
            stats_df = generate_statistics(df)
            if not stats_df.empty:
                st.dataframe(stats_df, use_container_width=True, hide_index=True)
                # 统计表导出
                stats_excel = to_excel_download(stats_df)
                st.download_button(
                    label="📥 下载统计报表",
                    data=stats_excel,
                    file_name=f"统计报表_{st.session_state.current_filter}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("暂无统计数据")

        with tab3:
            pivot_df = generate_pivot(df)
            if not pivot_df.empty:
                st.dataframe(pivot_df, use_container_width=True, hide_index=True)
            else:
                st.info("无法生成数据透视表，请检查E列和AE列是否存在")

        with tab4:
            fig = create_charts(df)
            if fig:
                st.pyplot(fig)
            else:
                st.info("暂无图表数据")
    else:
        # 空状态
        st.info("👈 请在左侧上传Excel文件开始分析")

        st.markdown("""
        ### 使用说明
        1. **上传文件**：在左侧边栏选择你的 Excel 文件（.xlsx / .xls）
        2. **自动预处理**：系统自动完成区县归类、回单状态标记、超时判定
        3. **筛选分析**：点击县份按钮或"无线超时工单"进行筛选
        4. **一键取消**：点击"取消筛选"恢复全部数据
        5. **多维度查看**：通过上方标签页切换表格、统计、透视、图表视图
        6. **导出结果**：每个页面都支持下载当前结果为 Excel

        ### 支持的列规范
        - **B列**：工单描述（用于区县关键字匹配）
        - **E列**：网元名称
        - **I列**：县份（为空时自动用B列匹配）
        - **N列**：回单时间（有数据=已回单）
        - **AR列**：处理时长（>8为超时）
        - **AE列**：网络类型（如"无线接入网"）
        - **AC列**：辅助标记列
        - **CT列**：是否标记（如"是"/"否"）
        """)

if __name__ == "__main__":
    main()
