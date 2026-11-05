import datetime as dt
import os
import streamlit as st
from memo_engine import (BRAND_CONFIG, FIELDS, initial_state, read_memo, local_parse,
                         ai_parse, validate_state, generate_memo)

st.set_page_config(page_title='促销 Memo 生成器', page_icon='👓', layout='centered')

CSS='''<style>
.stApp{background:#f6f7f9}.block-container{max-width:820px;padding-top:2rem}
[data-testid="stVerticalBlockBorderWrapper"]{background:white;border-radius:14px}
h1{font-size:1.9rem!important} .stButton button,.stDownloadButton button{border-radius:9px}
</style>'''


def api_key():
    try: return str(st.secrets.get('ECNU_API_KEY', '')).strip() or os.getenv('ECNU_API_KEY','').strip()
    except Exception: return os.getenv('ECNU_API_KEY','').strip()


def reset_editor():
    for k in list(st.session_state):
        if k.startswith('edit_'): del st.session_state[k]
    st.session_state.pop('output',None)


def date_value(value):
    return dt.date.fromisoformat(value) if value else None


def main():
    st.markdown(CSS,unsafe_allow_html=True)
    st.title('👓 促销 Memo 生成器')
    st.caption('导入活动内容，核对后生成门店 PPT。')
    with st.container(border=True):
        st.subheader('1 · 导入内容')
        brand=st.selectbox('品牌',list(BRAND_CONFIG),key='brand')
        with st.expander('继续修改已有 Memo（选填）'):
            st.caption('上传后以这份 PPT 为底稿，保留其他手工调整。不上传则使用品牌标准版。')
            ppt=st.file_uploader('已有 Memo',type=['pptx'],key='source_upload')
        source=ppt.getvalue() if ppt else None
        source_id=(brand,source)
        if st.session_state.get('source_id')!=source_id:
            try:
                state=read_memo(source,brand) if source else initial_state(brand)
            except Exception as exc:
                st.error(str(exc));return
            st.session_state.source_id=source_id
            st.session_state.memo=state
            st.session_state.baseline=dict(state) if source else None
            st.session_state.review=False
            reset_editor()
        wording=st.text_area('促销文案 / 修改意见',height=120,key='wording',
            placeholder='粘贴活动 wording，或输入“门店除奥莱”“优惠不适用于……”。')
        a1=st.file_uploader('A1 海报（选填，可后补）',type=['png','jpg','jpeg'],key='a1')
        image=a1.getvalue() if a1 else None
        c1,c2=st.columns([2,1])
        with c1: recognize=st.button('识别并核对 →',type='primary',use_container_width=True)
        with c2: manual=st.button('手动填写',use_container_width=True)
        if recognize:
            if not wording.strip() and not image and source is None:
                st.warning('请粘贴文案、上传海报，或选择手动填写。')
            else:
                state=st.session_state.memo
                key=api_key()
                if key and (wording.strip() or image):
                    try:
                        with st.spinner('正在识别活动信息…'):
                            state=ai_parse(key,brand,state,wording,image,a1.type if a1 else 'image/jpeg')
                        st.session_state.parse_note='已识别，请核对档期、日期与条款；未明确的信息保留底稿内容。'
                    except Exception:
                        try: state=local_parse(wording,state)
                        except ValueError:
                            st.error('文案日期无效，请检查后重试。');return
                        st.session_state.parse_note='智能识别暂不可用，仅尝试提取明确日期和标注字段。海报内容未识别，请手动核对。'
                elif wording.strip() or image:
                    try: state=local_parse(wording,state)
                    except ValueError:
                        st.error('文案日期无效，请检查后重试。');return
                    st.session_state.parse_note='当前未配置智能识别，仅支持明确日期和“档期：…”等标注字段；海报内容请手动填写。'
                else:
                    st.session_state.parse_note='已读取现有 Memo，请确认需要修改的字段。'
                st.session_state.memo=state
                st.session_state.review=True
                st.session_state.review_input=(wording,image)
                reset_editor()
        if manual:
            st.session_state.review=True
            st.session_state.review_input=(wording,image)
            st.session_state.parse_note='请填写本次活动信息；商品排除及通用限制来自当前底稿。'
            reset_editor()
    if not st.session_state.get('review'): return
    if st.session_state.get('review_input') != (wording,image):
        st.warning('导入内容已变化，请重新识别，或点击“手动填写”继续核对。')
        return
    with st.container(border=True):
        st.subheader('2 · 核对活动信息')
        st.caption(st.session_state.get('parse_note',''))
        state=st.session_state.memo
        state['festival']=st.text_input('节日 / 档期',value=state['festival'],key='edit_festival',placeholder='如：女神节、Super Weekend（不是海报宣传语）')
        a,b=st.columns(2)
        with a:
            value=st.date_input('开始日期',value=date_value(state['start_date']),key='edit_start_date')
            state['start_date']=value.isoformat() if value else ''
        with b:
            value=st.date_input('结束日期',value=date_value(state['end_date']),key='edit_end_date')
            state['end_date']=value.isoformat() if value else ''
        state['short_promo']=st.text_input('主要优惠',value=state['short_promo'],key='edit_short_promo')
        state['store_scope']=st.text_input('参与门店',value=state['store_scope'],key='edit_store_scope')
        state['exclude_outlets']=st.checkbox('奥莱不参与',value=state['exclude_outlets'],key='edit_exclude_outlets')
        state['rule_sentence']=st.text_area('优惠规则',value=state['rule_sentence'],key='edit_rule_sentence',height=90)
        state['exclusion_clause']=st.text_area('不参与商品 / 品牌',value=state['exclusion_clause'],key='edit_exclusion_clause',height=110)
        with st.expander('更多设置'):
            state['badge_promo']=st.text_area('左侧优惠标签',value=state['badge_promo'],key='edit_badge_promo',height=80)
            a,b=st.columns(2)
            with a: state['discount_code']=st.text_input('折扣代码',value=state['discount_code'],key='edit_discount_code')
            with b: state['store_code_scope']=st.text_input('涉及店铺（内部代码）',value=state['store_code_scope'],key='edit_store_code_scope')
            state['transaction_clause']=st.text_area('交易及叠加限制',value=state['transaction_clause'],key='edit_transaction_clause')
            state['legal_clause']=st.text_area('地域渠道及咨询条款',value=state['legal_clause'],key='edit_legal_clause')
            state['issue_month']=st.text_input('封面落款月份',value=state['issue_month'],key='edit_issue_month')
        st.session_state.memo=state
        if image: st.caption('A1：生成时替换两处海报，保持原图比例。')
        elif source: st.caption('A1：保留上传 Memo 中的海报。')
        else: st.caption('A1：暂用“海报待补充”占位，可稍后上传补图。')
        # Do not serve an output made from older reviewed fields.
        if st.session_state.get('output_state') != state:
            st.session_state.pop('output',None)
        if st.button('确认并生成 PPT',type='primary',use_container_width=True):
            errors=validate_state(state)
            if errors:
                for error in errors: st.error(error)
            else:
                try:
                    with st.spinner('正在更新 Memo…'):
                        data,count=generate_memo(brand,state,source,st.session_state.baseline,image)
                    st.session_state.output=data
                    st.session_state.output_state=dict(state)
                    st.session_state.output_count=count
                except Exception as exc:
                    st.error('生成失败：'+str(exc))
        if st.session_state.get('output'):
            name=state['start_date'].replace('-','')+' '+BRAND_CONFIG[brand]['code']+' promo memo.pptx'
            st.success('Memo 已生成，下载后即可检查。')
            st.download_button('下载 PPT',st.session_state.output,file_name=name,
                mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',use_container_width=True)


if __name__=='__main__': main()
