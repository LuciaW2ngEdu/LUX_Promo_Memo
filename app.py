import base64
import datetime
import io
import json
import os
import re
import streamlit as st
from pptx import Presentation
from openai import OpenAI

st.set_page_config(
    page_title="促销 Memo 工作台",
    page_icon="👓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
NL = chr(10)

# 注入极简高级定制 CSS
CUSTOM_CSS = """
<style>
    /* 隐藏默认顶部装饰条与多余菜单 */
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    [data-testid="collapsedControl"] {display: none;}

    /* 整体背景与容器间距 */
    .stApp {
        background-color: #F7F8FA;
    }
    .block-container {
        padding-top: 2.2rem !important;
        padding-bottom: 2.5rem !important;
        max-width: 1180px !important;
    }

    /* 顶部标题栏 */
    .studio-header {
        display: flex;
        align-items: baseline;
        justify-content: space-between;
        margin-bottom: 1.4rem;
        padding-bottom: 0.8rem;
        border-bottom: 1px solid #E6E8EC;
    }
    .studio-title {
        font-size: 1.65rem;
        font-weight: 700;
        color: #111827;
        letter-spacing: -0.02em;
        margin: 0;
    }
    .studio-subtitle {
        font-size: 0.88rem;
        color: #6B7280;
        margin-left: 0.75rem;
        font-weight: 400;
    }

    /* 卡片容器美化 */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #FFFFFF;
        border: 1px solid #EAECEF !important;
        border-radius: 14px !important;
        padding: 1.35rem 1.5rem !important;
        box-shadow: 0 2px 10px rgba(17, 24, 39, 0.03);
    }

    /* 卡片小标题 */
    .card-title {
        font-size: 1.02rem;
        font-weight: 600;
        color: #1F2937;
        margin-bottom: 0.9rem;
        display: flex;
        align-items: center;
        gap: 0.45rem;
    }
    .step-badge {
        background: #111827;
        color: #FFFFFF;
        font-size: 0.72rem;
        font-weight: 600;
        padding: 2px 8px;
        border-radius: 99px;
        letter-spacing: 0.03em;
    }

    /* 表单标签精简美化 */
    label[data-testid="stWidgetLabel"] p {
        font-size: 0.84rem !important;
        font-weight: 500 !important;
        color: #374151 !important;
    }

    /* 按钮圆角与质感 */
    .stButton > button, .stDownloadButton > button {
        border-radius: 10px !important;
        font-weight: 600 !important;
        padding: 0.55rem 1rem !important;
        transition: all 0.15s ease-in-out !important;
    }
    .stDownloadButton > button {
        background-color: #111827 !important;
        color: #FFFFFF !important;
        border: none !important;
        box-shadow: 0 4px 12px rgba(17, 24, 39, 0.12) !important;
    }
    .stDownloadButton > button:hover {
        background-color: #1F2937 !important;
        transform: translateY(-1px);
    }

    /* 日期联动预览条 */
    .date-pill-bar {
        background: #F3F4F6;
        border-radius: 8px;
        padding: 0.45rem 0.75rem;
        font-size: 0.78rem;
        color: #4B5563;
        margin-top: -0.2rem;
        margin-bottom: 0.85rem;
        display: flex;
        flex-wrap: wrap;
        gap: 0.8rem;
    }
    .date-pill-bar span {
        color: #111827;
        font-weight: 500;
    }
</style>
"""

BRAND_CONFIG = {
    "Oakley 欧克利": {
        "code": "OO",
        "candidate_files": [
            "templates/OO.pptx",
            "OO.pptx",
            "20260327 OO promo memo.pptx",
            "20260227 OO promo memo.pptx",
        ],
        "cover_prefix": "欧克利",
        "cover_suffix": "促销活动指引",
        "cover_date_full_year_end": False,
        "s2_title_tpl": "欧克利{festival}促销活动-概述",
        "default_discount_code": "UHC",
        "default_store_scope": "Oakely所有店铺",
        "default_Exclusion_clause": "此优惠不适用于部分产品部分限量款、指定款（ 包括 Plantaris, Plantaris SQ, Neo Jacket, Suture Jacket, Velo Kato, Retention Redefined，Aura系列）不参加。",
        "default_rule_sentence": "购买指定两副及以上产品，在吊牌价基础上享75折优惠。",
        "default_short_promo": "两副75折",
        "default_badge_promo": "两副75折",
    },
    "Sunglass Hut (SGH)": {
        "code": "SGH",
        "candidate_files": [
            "templates/SGH.pptx",
            "SGH.pptx",
            "20260327 SGH promo memo.pptx",
        ],
        "cover_prefix": "SUNGLASS HUT ",
        "cover_suffix": "促销活动指引",
        "cover_date_full_year_end": False,
        "s2_title_tpl": "SGH{festival}促销-概述",
        "default_discount_code": "UHC",
        "default_store_scope": "SGH MLC",
        "default_Exclusion_clause": "不包含26春夏新品、法拉利特别款、法拉利新品款、部分限量款及指定款。本次促销活动不参与的品牌有：Chanel、Oliver Peoples、Moncler、Prada、Prada Linea Rossa、Miu Miu、Dior、 Cartier、Lindberg、Ray-ban x A$AP rocky系列。",
        "default_rule_sentence": "购买指定两副及以上产品，在吊牌价基础上享75折优惠，包含变色片全套镜、度数太阳镜",
        "default_short_promo": "两件75折。",
        "default_badge_promo": "两件75折",
    },
    "Ray-Ban 雷朋": {
        "code": "RB",
        "candidate_files": [
            "templates/RB.pptx",
            "RB.pptx",
            "20260327 RB promo memo.pptx",
        ],
        "cover_prefix": "雷朋",
        "cover_suffix": "促销活动指引",
        "cover_date_full_year_end": False,
        "s2_title_tpl": "雷朋{festival}促销活动-概述",
        "default_discount_code": "UHC",
        "default_store_scope": "RB MLC",
        "default_Exclusion_clause": "优惠不适用于26春夏新品、Ray-ban x A$AP rocky、法拉利特别款、法拉利新品、部分限量款及指定产品；Nikon、Varilux XR镜片不参与该优惠。",
        "default_rule_sentence": "购买指定两副及以上产品，在吊牌价基础上享75折优惠。",
        "default_short_promo": "两件75折",
        "default_badge_promo": "两件75折",
    },
    "LensCrafters 亮视点": {
        "code": "LC",
        "candidate_files": [
            "templates/LC.pptx",
            "LC.pptx",
            "20260327 LC promo memo.pptx",
        ],
        "cover_prefix": "亮视点",
        "cover_suffix": "促销活动指引",
        "cover_date_full_year_end": True,
        "s2_title_tpl": "亮视点{festival}促销活动-概述",
        "default_discount_code": "UHC",
        "default_store_scope": "LC MLC",
        "default_Exclusion_clause": "镜片不参与的品牌及系列：Nikon & Varilux XR；不包含26春夏新品、法拉利特别款、法拉利新款、部分限量款及指定款。本次促销活动不参与的品牌有：Chanel、Oliver Peoples、Moncler、Prada、Prada Linea Rossa、Miu Miu、Dior、 Cartier、Lindberg、Ray-ban x A$AP rocky系列。",
        "default_rule_sentence": "购买指定两副及以上产品，在吊牌价基础上享75折优惠，包含变色片全套镜，度数太阳镜",
        "default_short_promo": "两件75折",
        "default_badge_promo": "两件75折",
    },
}

MONTH_ABBR = {
    1: "JAN", 2: "FEB", 3: "MAR", 4: "APR", 5: "MAY", 6: "JUN",
    7: "JUL", 8: "AUG", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DEC"
}


def get_backend_api_key() -> str:
    """静默从后台 Secrets 或环境变量读取密钥，不在前端暴露任何配置项"""
    try:
        key = st.secrets.get("ECNU_API_KEY", "")
        if key:
            return str(key).strip()
    except Exception:
        pass
    return os.environ.get("ECNU_API_KEY", "").strip()


def resolve_template_path(cfg: dict) -> str:
    for rel_path in cfg["candidate_files"]:
        full_p = os.path.join(BASE_DIR, rel_path)
        if os.path.exists(full_p):
            return full_p
        if os.path.exists(rel_path):
            return rel_path
    raise FileNotFoundError(f"未找到 {cfg['code']} 模板文件，请检查仓库中的模板。")


def set_paragraph_text_preserve_style(paragraph, new_text):
    if not paragraph.runs:
        paragraph.text = new_text
        return
    first_run = paragraph.runs[0]
    first_run.text = new_text
    for r in paragraph.runs[1:]:
        r.text = ""


def format_dates(start_date: datetime.date, end_date: datetime.date, full_year_end: bool = False):
    sy, sm, sd = start_date.year, start_date.month, start_date.day
    ey, em, ed = end_date.year, end_date.month, end_date.day

    dot_short = f"{sy}.{sm}.{sd} – {em}.{ed}"
    dot_cover = f"{sy}.{sm}.{sd} – {ey}.{em}.{ed}" if full_year_end else dot_short
    cn_full_dash = f"{sy}年{sm}月{sd}日 — {ey}年{em}月{ed}日"
    cn_start_only = f"{sm}月{sd}日"
    md_space_dash = f"{sm}.{sd} - {em}.{ed}"
    cn_range_to = f"{sy}年{sm}月{sd}日至{em}月{ed}日"
    cn_range_hyphen = f"{sy}年{sm}月{sd}日-{em}月{ed}日"
    md_tight_dash = f"{sm}.{sd}-{em}.{ed}"

    return {
        "dot_short": dot_short,
        "dot_cover": dot_cover,
        "cn_full_dash": cn_full_dash,
        "cn_start_only": cn_start_only,
        "md_space_dash": md_space_dash,
        "cn_range_to": cn_range_to,
        "cn_range_hyphen": cn_range_hyphen,
        "md_tight_dash": md_tight_dash,
    }


def generate_memo_pptx(
    brand_key: str,
    festival: str,
    issue_month_str: str,
    start_date: datetime.date,
    end_date: datetime.date,
    short_promo: str,
    badge_promo: str,
    discount_code: str,
    rule_sentence: str,
    exclusion_clause: str,
    a1_image_bytes: bytes = None,
) -> bytes:
    cfg = BRAND_CONFIG[brand_key]
    bcode = cfg["code"]
    tpl_path = resolve_template_path(cfg)
    prs = Presentation(tpl_path)
    d = format_dates(start_date, end_date, full_year_end=cfg["cover_date_full_year_end"])

    # SLIDE 1: 封面
    slide1 = prs.slides[0]
    for sh in slide1.shapes:
        if sh.name == "Titre 1" and sh.has_text_frame:
            p0 = sh.text_frame.paragraphs[0]
            if bcode == "SGH" and len(p0.runs) >= 4:
                p0.runs[0].text = cfg["cover_prefix"]
                p0.runs[1].text = f"{festival}{cfg['cover_suffix']}"
                p0.runs[2].text = "活动时间："
                p0.runs[3].text = f" {d['dot_cover']}"
                for r in p0.runs[4:]:
                    r.text = ""
            elif len(p0.runs) >= 3:
                p0.runs[0].text = f"{cfg['cover_prefix']}{festival}{cfg['cover_suffix']}"
                p0.runs[1].text = "活动时间："
                p0.runs[2].text = d["dot_cover"]
                for r in p0.runs[3:]:
                    r.text = ""
            else:
                set_paragraph_text_preserve_style(
                    p0, f"{cfg['cover_prefix']}{festival}{cfg['cover_suffix']}" + NL + f"活动时间：{d['dot_cover']}"
                )
        elif sh.name == "Espace réservé du texte 2" and sh.has_text_frame:
            set_paragraph_text_preserve_style(sh.text_frame.paragraphs[0], issue_month_str)

    # SLIDE 2: 促销活动-概述
    slide2 = prs.slides[1]
    for sh in slide2.shapes:
        if sh.name == "标题 2" and sh.has_text_frame:
            s2_title = cfg["s2_title_tpl"].format(festival=festival)
            set_paragraph_text_preserve_style(sh.text_frame.paragraphs[0], s2_title)

        elif sh.name == "TextBox 1" and sh.has_text_frame:
            lines = [line.strip() for line in badge_promo.splitlines() if line.strip()]
            joined_badge = (NL + NL).join(lines) if len(lines) > 1 else (lines[0] if lines else "")
            set_paragraph_text_preserve_style(sh.text_frame.paragraphs[0], joined_badge)
            for p in sh.text_frame.paragraphs[1:]:
                set_paragraph_text_preserve_style(p, "")

        elif sh.has_table:
            tbl = sh.table
            if sh.name == "Table 1":
                set_paragraph_text_preserve_style(tbl.cell(0, 1).text_frame.paragraphs[0], d["cn_full_dash"])
                space_before_qijian = "" if bcode == "OO" else " "
                main_content = f"{d['dot_short']}{space_before_qijian}期间，{short_promo}"
                set_paragraph_text_preserve_style(tbl.cell(2, 1).text_frame.paragraphs[0], main_content)
                set_paragraph_text_preserve_style(tbl.cell(4, 1).text_frame.paragraphs[0], d["cn_start_only"])
                set_paragraph_text_preserve_style(
                    tbl.cell(4, 2).text_frame.paragraphs[0],
                    f"{d['cn_start_only']}，活动开始陈列A1画面。"
                )

            elif sh.name == "Table 14":
                set_paragraph_text_preserve_style(tbl.cell(1, 1).text_frame.paragraphs[0], discount_code)
                set_paragraph_text_preserve_style(tbl.cell(1, 3).text_frame.paragraphs[0], d["md_space_dash"])
                tf = tbl.cell(1, 4).text_frame
                if bcode == "OO":
                    if len(tf.paragraphs) >= 3:
                        set_paragraph_text_preserve_style(tf.paragraphs[0], f"活动时间:{d['cn_range_hyphen']}")
                        set_paragraph_text_preserve_style(tf.paragraphs[1], rule_sentence.strip())
                        set_paragraph_text_preserve_style(tf.paragraphs[2], exclusion_clause.strip())
                elif bcode == "SGH":
                    if len(tf.paragraphs) >= 2:
                        set_paragraph_text_preserve_style(tf.paragraphs[0], f"1.   活动日期: {d['cn_range_to']}。")
                        r_clean = rule_sentence.strip().rstrip("。；;")
                        e_clean = exclusion_clause.strip()
                        set_paragraph_text_preserve_style(tf.paragraphs[1], f"2.  {r_clean}；{e_clean}")
                elif bcode == "RB":
                    if len(tf.paragraphs) >= 3:
                        set_paragraph_text_preserve_style(tf.paragraphs[0], f"活动日期: {d['cn_range_to']}。")
                        r_clean = rule_sentence.strip().rstrip("。") + "。"
                        set_paragraph_text_preserve_style(tf.paragraphs[1], r_clean)
                        set_paragraph_text_preserve_style(tf.paragraphs[2], exclusion_clause.strip())
                elif bcode == "LC":
                    if len(tf.paragraphs) >= 2:
                        set_paragraph_text_preserve_style(tf.paragraphs[0], f"活动日期: {d['cn_range_to']}。")
                        r_clean = rule_sentence.strip().rstrip("。，,；;")
                        e_clean = exclusion_clause.strip()
                        set_paragraph_text_preserve_style(tf.paragraphs[1], f"{r_clean}，{e_clean}")

    # 最后一页: 陈列指引与A1海报图
    last_slide = prs.slides[-1]
    for sh in last_slide.shapes:
        if sh.has_text_frame and "陈列时间" in sh.text_frame.text:
            set_paragraph_text_preserve_style(sh.text_frame.paragraphs[0], f"陈列时间： {d['md_tight_dash']}")

    if a1_image_bytes:
        pics = [sh for sh in last_slide.shapes if sh.shape_type == 13]
        target_slots = []
        for p_sh in pics:
            if p_sh.left < 5000000 or p_sh.left > 7000000:
                target_slots.append((p_sh.left, p_sh.top, p_sh.width, p_sh.height, p_sh))
        for left, top, width, height, old_sh in target_slots:
            old_sh._element.getparent().remove(old_sh._element)
            last_slide.shapes.add_picture(io.BytesIO(a1_image_bytes), left, top, width=width, height=height)

    out_stream = io.BytesIO()
    prs.save(out_stream)
    return out_stream.getvalue()


def smart_parse_with_ai(
    api_key: str,
    brand_key: str,
    current_state: dict,
    user_input_text: str = "",
    a1_image_bytes: bytes = None,
    a1_mime: str = "image/jpeg",
) -> dict:
    client = OpenAI(
        api_key=api_key,
        base_url="https://chat.ecnu.edu.cn/open/api/v1",
    )

    system_prompt = """你是一个专业的眼镜零售营销 Memo 数据提取与修改助手。
请根据用户提供的【A1海报图片】或【促销文案/修改要求】，提取或更新以下 JSON 字段（严格只返回合法的 JSON 对象，不要包含 markdown 代码块标记）：
{
  "festival": "节日或活动档期名称，如：清明节、女神节、五一劳动节（若未提及新节日，保留当前值）",
  "start_date": "活动开始日期，格式严格为 YYYY-MM-DD",
  "end_date": "活动结束日期，格式严格为 YYYY-MM-DD",
  "short_promo": "用于P2主要内容的简短折扣描述（将英文OFF转为中文折扣，例如：'两副75折'、'两件75折' 或 '全套镜 75 折,指定太阳镜 8 折'）",
  "badge_promo": "用于P2左下角优惠框的短语（例如：'两副75折' 或多行优惠）",
  "discount_code": "折扣代码，默认保留当前值（如 UHC），除非明确要求修改",
  "rule_sentence": "核心优惠规则句子（例如：'购买指定两副及以上产品，在吊牌价基础上享75折优惠。'）",
  "exclusion_clause": "不参与活动的商品/品牌/系列条款（除非在输入中要求增删排除系列/新品年份，否则保留当前品牌的默认排除条款文本）"
}"""

    user_text = f"""当前品牌：{brand_key}
当前内容状态：
{json.dumps(current_state, ensure_ascii=False, indent=2)}

用户输入的促销文案或修改指令：
{user_input_text if user_input_text.strip() else "（请根据上传的海报图片提取）"}

请输出更新后的完整 JSON："""

    content_list = [{"type": "text", "text": user_text}]
    if a1_image_bytes:
        b64_img = base64.b64encode(a1_image_bytes).decode("utf-8")
        content_list.append({
            "type": "image_url",
            "image_url": {"url": f"data:{a1_mime};base64,{b64_img}"}
        })

    # 默认优先使用支持多模态且响应快的 ecnu-plus，若异常自动切换 ecnu-max
    last_err = None
    for model_name in ["ecnu-plus", "ecnu-max"]:
        try:
            resp = client.chat.completions.create(
                model=model_name,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": content_list},
                ],
                temperature=0.1,
                response_format={"type": "json_object"},
            )
            raw = resp.choices[0].message.content.strip()
            raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw.strip(), flags=re.IGNORECASE)
            return json.loads(raw)
        except Exception as e:
            last_err = e
            continue
    raise RuntimeError(str(last_err))


def local_regex_parse(user_input_text: str, current_state: dict) -> dict:
    updated = dict(current_state)
    m = re.search(
        r"(202\d)[年\./-](\d{1,2})[月\./-](\d{1,2})日?\s*[至\-–—~到]\s*(?:(202\d)[年\./-])?(\d{1,2})[月\./-](\d{1,2})",
        user_input_text,
    )
    if m:
        sy, sm, sd = int(m.group(1)), int(m.group(2)), int(m.group(3))
        ey = int(m.group(4)) if m.group(4) else sy
        em, ed = int(m.group(5)), int(m.group(6))
        updated["start_date"] = f"{sy:04d}-{sm:02d}-{sd:02d}"
        updated["end_date"] = f"{ey:04d}-{em:02d}-{ed:02d}"
    return updated


def main():
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    # 极简顶部导航头
    st.markdown(
        '<div class="studio-header">'
        '<div><span class="studio-title">👓 促销 Memo 工作台</span>'
        '<span class="studio-subtitle">自动排版 · 全篇日期联动 · 完美保留品牌母版样式</span></div>'
        '</div>',
        unsafe_allow_html=True,
    )

    col_left, col_right = st.columns([1, 1.25], gap="large")

    with col_left:
        with st.container(border=True):
            st.markdown(
                '<div class="card-title"><span class="step-badge">STEP 1</span> 品牌与活动输入</div>',
                unsafe_allow_html=True,
            )

            b_col, f_col = st.columns([1.15, 1])
            with b_col:
                selected_brand = st.selectbox("选择品牌", list(BRAND_CONFIG.keys()))
                cfg = BRAND_CONFIG[selected_brand]

            # 初始化或切换品牌时的默认状态
            if "last_brand" not in st.session_state or st.session_state.last_brand != selected_brand:
                st.session_state.last_brand = selected_brand
                st.session_state.festival = st.session_state.get("festival", "清明节")
                st.session_state.start_date = st.session_state.get("start_date", datetime.date(2026, 3, 27))
                st.session_state.end_date = st.session_state.get("end_date", datetime.date(2026, 4, 6))
                st.session_state.issue_month = f"{MONTH_ABBR[datetime.date.today().month]}, {datetime.date.today().year}"
                st.session_state.short_promo = cfg["default_short_promo"]
                st.session_state.badge_promo = cfg["default_badge_promo"]
                st.session_state.discount_code = cfg["default_discount_code"]
                st.session_state.rule_sentence = cfg["default_rule_sentence"]
                st.session_state.exclusion_clause = cfg["default_Exclusion_clause"]

            with f_col:
                festival_input = st.text_input(
                    "活动档期 / 节日",
                    value=st.session_state.festival,
                    placeholder="如：清明节、女神节、五一"
                )
                st.session_state.festival = festival_input

            uploaded_a1 = st.file_uploader(
                "上传 A1 海报图（选填，将自动提取文字并替换末页陈列图）",
                type=["jpg", "jpeg", "png"],
            )

            smart_input = st.text_area(
                "粘贴活动 Wording 或直接输入修改要求",
                height=135,
                placeholder="可直接粘贴整段促销文案，或输入修改指令（例如：活动时间延后到4月10日，不参加系列加上LNY）...",
            )

            if st.button("✨ 智能识别并填入右侧", use_container_width=True):
                current_state = {
                    "festival": st.session_state.festival,
                    "start_date": st.session_state.start_date.strftime("%Y-%m-%d"),
                    "end_date": st.session_state.end_date.strftime("%Y-%m-%d"),
                    "short_promo": st.session_state.short_promo,
                    "badge_promo": st.session_state.badge_promo,
                    "discount_code": st.session_state.discount_code,
                    "rule_sentence": st.session_state.rule_sentence,
                    "exclusion_clause": st.session_state.exclusion_clause,
                }
                api_key = get_backend_api_key()
                if api_key and (smart_input.strip() or uploaded_a1):
                    with st.spinner("正在智能解析内容..."):
                        try:
                            img_bytes = uploaded_a1.getvalue() if uploaded_a1 else None
                            img_mime = uploaded_a1.type if uploaded_a1 else "image/jpeg"
                            res = smart_parse_with_ai(
                                api_key=api_key,
                                brand_key=selected_brand,
                                current_state=current_state,
                                user_input_text=smart_input,
                                a1_image_bytes=img_bytes,
                                a1_mime=img_mime,
                            )
                            st.session_state.festival = res.get("festival", st.session_state.festival)
                            if res.get("start_date"):
                                st.session_state.start_date = datetime.date.fromisoformat(res["start_date"])
                            if res.get("end_date"):
                                st.session_state.end_date = datetime.date.fromisoformat(res["end_date"])
                            st.session_state.short_promo = res.get("short_promo", st.session_state.short_promo)
                            st.session_state.badge_promo = res.get("badge_promo", st.session_state.badge_promo)
                            st.session_state.discount_code = res.get("discount_code", st.session_state.discount_code)
                            st.session_state.rule_sentence = res.get("rule_sentence", st.session_state.rule_sentence)
                            st.session_state.exclusion_clause = res.get("exclusion_clause", st.session_state.exclusion_clause)
                            st.rerun()
                        except Exception:
                            res = local_regex_parse(smart_input, current_state)
                            st.session_state.start_date = datetime.date.fromisoformat(res["start_date"])
                            st.session_state.end_date = datetime.date.fromisoformat(res["end_date"])
                            st.rerun()
                else:
                    res = local_regex_parse(smart_input, current_state)
                    st.session_state.start_date = datetime.date.fromisoformat(res["start_date"])
                    st.session_state.end_date = datetime.date.fromisoformat(res["end_date"])
                    st.rerun()

    with col_right:
        with st.container(border=True):
            st.markdown(
                '<div class="card-title"><span class="step-badge">STEP 2</span> 内容核对与导出</div>',
                unsafe_allow_html=True,
            )

            c1, c2, c3 = st.columns([1, 1, 0.85])
            with c1:
                start_d = st.date_input("开始日期", value=st.session_state.start_date)
                st.session_state.start_date = start_d
            with c2:
                end_d = st.date_input("结束日期", value=st.session_state.end_date)
                st.session_state.end_date = end_d
            with c3:
                issue_m = st.text_input("封面月份", value=st.session_state.issue_month)
                st.session_state.issue_month = issue_m

            d_preview = format_dates(
                st.session_state.start_date, st.session_state.end_date, cfg["cover_date_full_year_end"]
            )
            st.markdown(
                f'<div class="date-pill-bar">'
                f'<div>封面 <span>{d_preview["dot_cover"]}</span></div>'
                f'<div>主表 <span>{d_preview["cn_full_dash"]}</span></div>'
                f'<div>执行日 <span>{d_preview["cn_start_only"]}</span></div>'
                f'<div>陈列 <span>{d_preview["md_tight_dash"]}</span></div>'
                f'</div>',
                unsafe_allow_html=True,
            )

            c4, c5, c6 = st.columns([1.2, 1.2, 0.7])
            with c4:
                st.session_state.short_promo = st.text_input("主要优惠简述", value=st.session_state.short_promo)
            with c5:
                st.session_state.badge_promo = st.text_input("左下角优惠标签", value=st.session_state.badge_promo)
            with c6:
                st.session_state.discount_code = st.text_input("折扣代码", value=st.session_state.discount_code)

            st.session_state.rule_sentence = st.text_area(
                "优惠规则正文",
                value=st.session_state.rule_sentence,
                height=68,
            )
            st.session_state.exclusion_clause = st.text_area(
                "不参与商品及品牌说明",
                value=st.session_state.exclusion_clause,
                height=96,
            )

            try:
                a1_bytes_to_use = uploaded_a1.getvalue() if uploaded_a1 else None
                ppt_bytes = generate_memo_pptx(
                    brand_key=selected_brand,
                    festival=st.session_state.festival,
                    issue_month_str=st.session_state.issue_month,
                    start_date=st.session_state.start_date,
                    end_date=st.session_state.end_date,
                    short_promo=st.session_state.short_promo,
                    badge_promo=st.session_state.badge_promo,
                    discount_code=st.session_state.discount_code,
                    rule_sentence=st.session_state.rule_sentence,
                    exclusion_clause=st.session_state.exclusion_clause,
                    a1_image_bytes=a1_bytes_to_use,
                )

                today_str = datetime.date.today().strftime("%Y%m%d")
                out_filename = f"{today_str} {cfg['code']} {st.session_state.festival} promo memo.pptx"

                st.download_button(
                    label=f"📥 生成并下载 {cfg['code']} Memo ({out_filename})",
                    data=ppt_bytes,
                    file_name=out_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                    type="primary",
                )
            except FileNotFoundError as fnf_err:
                st.error(str(fnf_err))


if __name__ == "__main__":
    main()
