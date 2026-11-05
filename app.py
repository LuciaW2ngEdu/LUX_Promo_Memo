import base64
import datetime
import io
import json
import os
import re
import time
import streamlit as st
from pptx import Presentation
from openai import OpenAI

st.set_page_config(
    page_title="促销 Memo 生成器",
    page_icon="👓",
    layout="centered",
    initial_sidebar_state="collapsed",
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
NL = chr(10)

CUSTOM_CSS = """
<style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    [data-testid="collapsedControl"] {display: none;}

    .stApp {
        background-color: #F6F7F9;
    }
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 3rem !important;
        max-width: 780px !important;
    }

    /* 顶部标题 */
    .hero-box {
        text-align: center;
        margin-bottom: 1.5rem;
    }
    .hero-title {
        font-size: 1.65rem;
        font-weight: 700;
        color: #111827;
        margin-bottom: 0.3rem;
        letter-spacing: -0.02em;
    }
    .hero-desc {
        font-size: 0.9rem;
        color: #6B7280;
    }

    /* 白色主卡片 */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #FFFFFF;
        border: 1px solid #E5E7EB !important;
        border-radius: 16px !important;
        padding: 1.5rem 1.6rem !important;
        box-shadow: 0 4px 16px rgba(17, 24, 39, 0.04);
    }

    label[data-testid="stWidgetLabel"] p {
        font-size: 0.88rem !important;
        font-weight: 600 !important;
        color: #1F2937 !important;
    }

    .stButton > button, .stDownloadButton > button {
        border-radius: 10px !important;
        font-weight: 600 !important;
        padding: 0.65rem 1.1rem !important;
        font-size: 0.95rem !important;
    }
    .stDownloadButton > button {
        background-color: #059669 !important;
        color: #FFFFFF !important;
        border: none !important;
        box-shadow: 0 4px 12px rgba(5, 150, 105, 0.2) !important;
    }
    .stDownloadButton > button:hover {
        background-color: #047857 !important;
    }

    /* 生成结果摘要卡片 */
    .result-card {
        background: #F0FDF4;
        border: 1px solid #BBF7D0;
        border-radius: 12px;
        padding: 1rem 1.2rem;
        margin-top: 0.5rem;
        margin-bottom: 1rem;
    }
    .result-header {
        font-size: 1rem;
        font-weight: 700;
        color: #065F46;
        margin-bottom: 0.55rem;
    }
    .result-grid {
        display: grid;
        grid-template-columns: 95px 1fr;
        row-gap: 0.35rem;
        font-size: 0.85rem;
        color: #374151;
    }
    .result-label {
        color: #6B7280;
        font-weight: 500;
    }
    .result-val {
        color: #111827;
        font-weight: 500;
    }
    .a1-tag-ok {
        display: inline-block;
        background: #D1FAE5;
        color: #065F46;
        font-size: 0.78rem;
        padding: 1px 8px;
        border-radius: 99px;
        font-weight: 600;
    }
    .a1-tag-wait {
        display: inline-block;
        background: #FEF3C7;
        color: #92400E;
        font-size: 0.78rem;
        padding: 1px 8px;
        border-radius: 99px;
        font-weight: 600;
    }
</style>
"""

BRAND_CONFIG = {
    "Oakley 欧克利 (OO)": {
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
    "Ray-Ban 雷朋 (RB)": {
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
    "LensCrafters 亮视点 (LC)": {
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
    raise FileNotFoundError(f"未找到 {cfg['code']} 品牌的 PPT 母版文件，请检查仓库中的 templates 目录。")


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

    # 最后一页: 陈列指引与A1海报图（若未上传新图，保留母版默认占位图）
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
        timeout=25.0,
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

    st.markdown(
        '<div class="hero-box">'
        '<div class="hero-title">👓 促销 Memo 生成器</div>'
        '<div class="hero-desc">选择品牌并输入文案或 A1 海报（二选一即可），一键生成标准门店 PPT</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    with st.container(border=True):
        b_col, f_col = st.columns([1.15, 1])
        with b_col:
            selected_brand = st.selectbox("1. 选择品牌", list(BRAND_CONFIG.keys()))
            cfg = BRAND_CONFIG[selected_brand]

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
            st.session_state.ready = True

        with f_col:
            festival_input = st.text_input(
                "2. 活动节日名称",
                value=st.session_state.festival,
                placeholder="如：清明节、女神节、五一劳动节"
            )
            st.session_state.festival = festival_input

        smart_input = st.text_area(
            "3. 粘贴促销 Wording 文字 或 输入修改意见（与下方 A1 海报二选一即可）",
            height=115,
            placeholder="方式 A：直接粘贴整段促销 Wording 文案\n方式 B：改版时直接输入修改要求，例如「活动延后到4月10日结束，去掉LNY系列」",
        )

        uploaded_a1 = st.file_uploader(
            "4. 上传 A1 海报图片（选填：若暂未出图可不传，等供应商出图后再传图更新）",
            type=["jpg", "jpeg", "png"],
        )

        if st.button("✨ 一键生成促销 Memo", use_container_width=True, type="primary"):
            if not smart_input.strip() and not uploaded_a1:
                st.warning("💡 请至少粘贴一段促销 Wording 文字、输入修改要求，或上传一张 A1 海报图片；若只需手动微调字段，可展开下方「手动微调细节」直接下载。")
            else:
                prog = st.progress(10, text="⏳ 正在读取输入内容 (10%)...")
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
                if api_key:
                    try:
                        prog.progress(45, text="🧠 正在智能提取活动时间与折扣条款 (45%)...")
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
                        prog.progress(85, text="📝 正在同步更新全篇 PPT 日期与排版 (85%)...")
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
                        prog.progress(100, text="✅ 制作完成 (100%)！")
                        time.sleep(0.2)
                        prog.empty()
                        st.session_state.just_generated = True
                    except Exception as err:
                        prog.empty()
                        res = local_regex_parse(smart_input, current_state)
                        st.session_state.start_date = datetime.date.fromisoformat(res["start_date"])
                        st.session_state.end_date = datetime.date.fromisoformat(res["end_date"])
                        st.warning(
                            f"⚠️ 智能解析服务响应超时或不可用（原因：{err}）。已自动使用本地规则提取日期，建议点开下方「手动微调细节」核对一下优惠文字后再下载。"
                        )
                else:
                    prog.progress(70, text="📝 正在提取日期并排版 (70%)...")
                    res = local_regex_parse(smart_input, current_state)
                    st.session_state.start_date = datetime.date.fromisoformat(res["start_date"])
                    st.session_state.end_date = datetime.date.fromisoformat(res["end_date"])
                    prog.progress(100, text="✅ 制作完成 (100%)！")
                    time.sleep(0.2)
                    prog.empty()
                    st.info("ℹ️ 后台尚未配置智能解析密钥，已使用基础规则提取日期；可展开下方「手动微调细节」核对优惠文字。")

        # 标准文件命名：促销开始日期(YYYYMMDD) + 品牌简称 + promo memo.pptx
        start_str = st.session_state.start_date.strftime("%Y%m%d")
        out_filename = f"{start_str} {cfg['code']} promo memo.pptx"

        # 实时生成 PPT 文件字节流并展示下载卡片
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

            d_prev = format_dates(st.session_state.start_date, st.session_state.end_date, cfg["cover_date_full_year_end"])
            a1_status_html = (
                '<span class="a1-tag-ok">已替换最新上传的 A1 海报（共 2 处）</span>'
                if uploaded_a1
                else '<span class="a1-tag-wait">未上传新图（末页暂保留模板原图，出图后可随时重传更新）</span>'
            )

            st.markdown(
                f'<div class="result-card">'
                f'<div class="result-header">📄 已就绪：{out_filename}</div>'
                f'<div class="result-grid">'
                f'<div class="result-label">品牌与档期</div><div class="result-val">{cfg["code"]} · {st.session_state.festival}</div>'
                f'<div class="result-label">活动时间</div><div class="result-val">{d_prev["cn_full_dash"]}（已自动同步全篇 6 处日期）</div>'
                f'<div class="result-label">主要优惠</div><div class="result-val">{st.session_state.short_promo}（折扣码：{st.session_state.discount_code}）</div>'
                f'<div class="result-label">核心规则</div><div class="result-val">{st.session_state.rule_sentence}</div>'
                f'<div class="result-label">A1 海报图</div><div class="result-val">{a1_status_html}</div>'
                f'</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

            st.download_button(
                label=f"📥 立即下载 {out_filename}",
                data=ppt_bytes,
                file_name=out_filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )
        except Exception as gen_err:
            st.error(f"❌ 生成 PPT 时出现问题：{gen_err}。请检查仓库中的 templates 母版文件是否完整。")

        # 默认折叠的手动微调抽屉（非必填，消除视觉与操作负担）
        with st.expander("🔧 需要手动微调某处文字或日期？（非必填，点击展开直接改）", expanded=False):
            st.caption("修改下方任意一项后，上方的绿色预览卡片与下载文件将立即同步更新：")
            c1, c2, c3 = st.columns([1, 1, 0.85])
            with c1:
                new_sd = st.date_input("开始日期", value=st.session_state.start_date)
                if new_sd != st.session_state.start_date:
                    st.session_state.start_date = new_sd
                    st.rerun()
            with c2:
                new_ed = st.date_input("结束日期", value=st.session_state.end_date)
                if new_ed != st.session_state.end_date:
                    st.session_state.end_date = new_ed
                    st.rerun()
            with c3:
                new_im = st.text_input("封面左下角月份", value=st.session_state.issue_month)
                if new_im != st.session_state.issue_month:
                    st.session_state.issue_month = new_im
                    st.rerun()

            c4, c5, c6 = st.columns([1.2, 1.2, 0.7])
            with c4:
                new_sp = st.text_input("主要优惠简述", value=st.session_state.short_promo)
                if new_sp != st.session_state.short_promo:
                    st.session_state.short_promo = new_sp
                    st.rerun()
            with c5:
                new_bp = st.text_input("左下角优惠标签", value=st.session_state.badge_promo)
                if new_bp != st.session_state.badge_promo:
                    st.session_state.badge_promo = new_bp
                    st.rerun()
            with c6:
                new_dc = st.text_input("折扣代码", value=st.session_state.discount_code)
                if new_dc != st.session_state.discount_code:
                    st.session_state.discount_code = new_dc
                    st.rerun()

            new_rs = st.text_area("优惠规则正文", value=st.session_state.rule_sentence, height=68)
            if new_rs != st.session_state.rule_sentence:
                st.session_state.rule_sentence = new_rs
                st.rerun()

            new_ec = st.text_area("不参与商品及品牌说明", value=st.session_state.exclusion_clause, height=95)
            if new_ec != st.session_state.exclusion_clause:
                st.session_state.exclusion_clause = new_ec
                st.rerun()


if __name__ == "__main__":
    main()
