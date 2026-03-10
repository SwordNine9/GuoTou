import streamlit as st
import json
import os
import tempfile
import matplotlib
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml.ns import qn
from openai import OpenAI
from io import BytesIO

# 设置 matplotlib 后端，防止在无窗口环境下报错
matplotlib.use("Agg")

# ================= 1. 页面基础配置 =================
st.set_page_config(
    page_title="电力政策专刊生成助手",
    page_icon="⚡",
    layout="wide"
)

# CSS 美化按钮和背景
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { 
        width: 100%; 
        background-color: #0066cc; 
        color: white; 
        font-weight: bold;
        height: 50px;
        border-radius: 8px;
    }
    .stButton>button:hover {
        background-color: #0052a3;
    }
    </style>
    """, unsafe_allow_html=True)

# ================= 2. 核心工具函数 =================

def set_run_font(run, font_name='仿宋', size=None, is_bold=False):
    """
    统一设置中西文字体为 '仿宋'，并支持字号和加粗
    """
    # 1. 设置字体名称
    run.font.name = font_name
    
    # 2. 关键：同时设置中文(eastAsia)和西文(ascii/hAnsi)字体
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    r.rPr.rFonts.set(qn('w:ascii'), font_name)
    r.rPr.rFonts.set(qn('w:hAnsi'), font_name)
    
    # 3. 设置字号 (Pt 对象)
    if size:
        run.font.size = size
        
    # 4. 设置加粗
    if is_bold:
        run.bold = True

def generate_policy_analysis(client, summary):
    """
    调用 DeepSeek API 生成政策解析
    """
    if not client:
        return "（因未配置 API Key，系统跳过 AI 解析生成）"
    
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "你是一位电力政策分析专家。请根据提供的政策摘要，写一段简短的分析（100字左右），重点说明对发电企业的影响及应对建议。"},
                {"role": "user", "content": summary}
            ],
            stream=False
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"AI 解析生成失败: {str(e)}"

def generate_pie_chart(imp_data):
    """
    生成装机容量饼状图
    """
    categories = ["Hydro", "Thermal", "Nuclear", "Solar", "Wind"]
    keys = ["hydro_capacity", "thermal_capacity", "nuclear_capacity", "solar_capacity", "wind_capacity"]
    
    values = []
    for k in keys:
        try:
            val = float(imp_data.get(k, 0) or 0)
        except:
            val = 0
        values.append(val)
    
    if sum(values) == 0:
        return None

    # 创建临时文件保存图片
    tfile = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    
    plt.figure(figsize=(6, 6))
    plt.pie(values, labels=categories, autopct="%.1f%%", startangle=90)
    plt.title("Power Installed Capacity Structure") # 英文标题避免乱码
    plt.tight_layout()
    plt.savefig(tfile.name)
    plt.close()
    
    return tfile.name

# ================= 3. 文档处理核心逻辑 =================

def insert_content_after_keyword(doc, keyword, content_data, content_type, client):
    """
    在指定标题下方插入内容。
    顺序：标题 -> 机构 -> 时间 -> 摘要 -> 解析
    样式：仿宋 + 三号字 (16pt)
    """
    # 1. 寻找插入锚点（关键词所在段落）
    target_index = -1
    for i, paragraph in enumerate(doc.paragraphs):
        if keyword in paragraph.text:
            target_index = i
            break
    
    if target_index == -1:
        st.warning(f"⚠️ 模板中未找到章节标题：“{keyword}”，已跳过该部分。")
        return

    # 2. 确定插入位置（在标题的下一段开始插入）
    if target_index == len(doc.paragraphs) - 1:
        base_node = doc.add_paragraph("")
    else:
        base_node = doc.paragraphs[target_index + 1]

    font_size_standard = Pt(16) # 三号字

    # === 关键：正序遍历 content_data ===
    # 这样 item 1 会排在 item 2 前面
    for item in content_data:
        
        # --- 特殊处理：参考资料 ---
        if content_type == "reference":
            text = f"{item.get('source', '')}：{item.get('title', '')} ({item.get('url', '')})"
            p = base_node.insert_paragraph_before()
            run = p.add_run(text)
            set_run_font(run, font_name='仿宋', size=font_size_standard)
            continue # 参考资料不需要后续字段，直接下一条

        # --- 通用政策字段插入 (按阅读顺序) ---
        
        # 1. 政策名称 / 标题
        p = base_node.insert_paragraph_before()
        label = "政策标题" if content_type == "energy" else "政策名称"
        run = p.add_run(f"{label}：{item.get('title', '')}")
        set_run_font(run, font_name='仿宋', size=font_size_standard, is_bold=True)

        # 2. 区域 (仅能源新政)
        if content_type == "energy":
            p = base_node.insert_paragraph_before()
            run = p.add_run(f"区域：{item.get('region', '')}")
            set_run_font(run, font_name='仿宋', size=font_size_standard, is_bold=True)

        # 3. 发布机构 (仅政府文件)
        if content_type == "policy":
            p = base_node.insert_paragraph_before()
            run = p.add_run(f"发布机构：{item.get('agency', '')}")
            set_run_font(run, font_name='仿宋', size=font_size_standard)

        # 4. 发布时间
        p = base_node.insert_paragraph_before()
        date_label = "发布日期" if content_type == "energy" else "发布时间"
        run = p.add_run(f"{date_label}：{item.get('date', '')}")
        set_run_font(run, font_name='仿宋', size=font_size_standard)

        # 5. 政策摘要
        p = base_node.insert_paragraph_before()
        run = p.add_run(f"政策摘要：{item.get('summary', '')}")
        set_run_font(run, font_name='仿宋', size=font_size_standard)

        # 6. 政策解析 (生成并插入)
        analysis = generate_policy_analysis(client, item.get('summary', ''))
        p = base_node.insert_paragraph_before()
        run = p.add_run(f"政策解析：{analysis}")
        set_run_font(run, font_name='仿宋', size=font_size_standard)

        # 7. 插入空行分隔 (每条政策之间)
        base_node.insert_paragraph_before("")

def process_document(template_file, json_data, api_key):
    """
    文档处理主程序
    """
    doc = Document(template_file)
    imp_data = json_data.get("important_data_values", {})
    
    # 初始化 DeepSeek 客户端
    client = None
    if api_key:
        client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")

    # ---------------- Step 1: 替换 ** 占位符 ----------------
    fill_order = [
        imp_data.get("report_date"),
        imp_data.get("total_capacity"), imp_data.get("total_capacity_yoy"),
        imp_data.get("hydro_capacity"), imp_data.get("hydro_yoy"),
        imp_data.get("thermal_capacity"), imp_data.get("thermal_yoy"),
        imp_data.get("nuclear_capacity"), imp_data.get("nuclear_yoy"),
        imp_data.get("solar_capacity"), imp_data.get("solar_yoy"),
        imp_data.get("wind_capacity"), imp_data.get("wind_yoy"),
        imp_data.get("report_date"), 
        imp_data.get("market_trade"), imp_data.get("market_trade_yoy"),
        imp_data.get("provincial_trade"), imp_data.get("provincial_trade_yoy"),
        imp_data.get("cross_region_trade"), imp_data.get("cross_region_trade_yoy"),
        imp_data.get("green_trade"), imp_data.get("green_trade_yoy"),
        imp_data.get("report_date"), 
        imp_data.get("society_usage"), imp_data.get("society_usage_yoy"),
        imp_data.get("primary_usage"), imp_data.get("primary_usage_yoy"),
        imp_data.get("secondary_usage"), imp_data.get("secondary_usage_yoy"),
        imp_data.get("tertiary_usage"), imp_data.get("tertiary_usage_yoy")
    ]
    
    values_iter = iter(fill_order)
    
    for paragraph in doc.paragraphs:
        if "**" in paragraph.text:
            count = paragraph.text.count("**")
            for _ in range(count):
                try:
                    val = next(values_iter)
                    for run in paragraph.runs:
                        if "**" in run.text:
                            run.text = run.text.replace("**", str(val), 1)
                            # 替换后的文本也强制设为仿宋（但保持原字号）
                            set_run_font(run, font_name='仿宋')
                            break
                except StopIteration:
                    break

    # ---------------- Step 2: 插入图表 ----------------
    chart_path = generate_pie_chart(imp_data)
    if chart_path:
        for i, p in enumerate(doc.paragraphs):
            if "（一）全国电力供应数据" in p.text:
                if i + 1 < len(doc.paragraphs):
                    target_p = doc.paragraphs[i+1]
                    run = target_p.add_run()
                    run.add_break()
                    run.add_picture(chart_path, width=Inches(5.0))
                break
        os.unlink(chart_path)

    # ---------------- Step 3: 插入内容 (仿宋 + 三号) ----------------
    
    # 插入政府文件
    insert_content_after_keyword(doc, "二、政府文件", json_data.get("gov_policies", []), "policy", client)
    
    # 插入能源新政
    new_policies = json_data.get("energy_new_policies", {})
    if "电力交易新政" in new_policies:
        insert_content_after_keyword(doc, "一、电力交易新政", new_policies["电力交易新政"], "energy", client)
    if "区域电价政策" in new_policies:
        insert_content_after_keyword(doc, "二、区域电价政策", new_policies["区域电价政策"], "energy", client)
    if "重点开发政策" in new_policies:
        insert_content_after_keyword(doc, "三、重点开发政策", new_policies["重点开发政策"], "energy", client)
        
    # 插入参考资料
    insert_content_after_keyword(doc, "参考资料", json_data.get("references", []), "reference", client)

    # 导出
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# ================= 4. Streamlit 前端界面 =================

st.title("国投电力政策资讯专刊生成系统")


with st.sidebar:
    st.header("⚙️ 参数设置")
    api_key = st.text_input("DeepSeek API Key", type="password", help="输入 Key 以启用 AI 自动解析功能")
    st.info("💡 提示：请确保 Word 模板中包含 '二、政府文件' 等标准标题，否则无法定位插入点。")

# 双栏上传布局
col1, col2 = st.columns(2)
with col1:
    st.subheader("📂 1. 上传数据 (JSON)")
    uploaded_json = st.file_uploader("选择 policy_data.json", type="json")
    if uploaded_json:
        try:
            json_data = json.load(uploaded_json)
            st.success(f"✅ 数据加载成功")
        except:
            st.error("❌ JSON 格式错误")

with col2:
    st.subheader("📄 2. 上传模板 (Word)")
    uploaded_template = st.file_uploader("选择 模板.docx", type="docx")
    if uploaded_template:
        st.success("✅ 模板加载成功")

st.markdown("---")

# 生成按钮
if st.button("🚀 开始生成报告", disabled=not (uploaded_json and uploaded_template)):
    with st.spinner('正在处理：替换数据 -> 绘制图表 -> AI解析 -> 仿宋排版 ...'):
        try:
            result_doc = process_document(uploaded_template, json_data, api_key)
            
            st.balloons()
            st.success("🎉 报告生成成功！已应用 **仿宋 + 三号字 (16pt)** 格式。")
            
            st.download_button(
                label="📥 下载 Word 文档",
                data=result_doc,
                file_name="政策资讯专刊_自动生成.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"生成失败: {str(e)}")
            st.error("请检查模板标题是否被修改，或 JSON 数据结构是否完整。")