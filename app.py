import datetime as dt
import os
import streamlit as st
from memo_engine import (BRAND_CONFIG, initial_state, read_memo, local_parse, ai_parse,
                         validate_state, generate_memo, check_source_brand, BrandMismatch, RecognitionError)

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
    with st.container(border=True):
        st.subheader('1 · 活动信息')
        brand_col,festival_col=st.columns([1,1])
        with brand_col:
            brand=st.selectbox('品牌',list(BRAND_CONFIG),key='brand')
        # Uploaded widgets update session_state before a rerun. Read the bottom
        # advanced control here so importing a Memo can populate the top fields.
        ppt=st.session_state.get('source_upload')
        source=ppt.getvalue() if ppt else None
        source_id=(brand,source)
        if st.session_state.get('source_id')!=source_id:
            try: state=read_memo(source,brand) if source else initial_state(brand)
            except Exception as exc:
                st.error(str(exc));return
            st.session_state.source_id=source_id
            st.session_state.memo=state
            st.session_state.baseline=dict(state) if source else None
            st.session_state.review=False
            st.session_state.festival_input=state['festival']
            reset_editor()
        with festival_col:
            festival=st.text_input('节日 / 档期',key='festival_input',placeholder='如：女神节、超级周末')
        st.session_state.memo['festival']=festival.strip()
        required_message=st.empty()
        method=st.radio('素材来源',['上传 A1','粘贴文案'],horizontal=True,key='input_method')
        if method=='上传 A1':
            a1=st.file_uploader('A1 海报',type=['png','jpg','jpeg'],key='a1',label_visibility='collapsed')
            with st.expander('补充文案 / 修改意见'):
                wording=st.text_area('文案 / 修改意见',height=100,key='wording',
                    placeholder='粘贴活动 wording 或修改要求。',label_visibility='collapsed')
        else:
            wording=st.text_area('文案 / 修改意见',height=100,key='wording',
                placeholder='粘贴活动 wording 或修改要求。',label_visibility='collapsed')
            with st.expander('同时上传 A1（选填）'):
                a1=st.file_uploader('A1 海报',type=['png','jpg','jpeg'],key='a1',label_visibility='collapsed')
        image=a1.getvalue() if a1 else None
        recognize=st.button('识别并核对 →',type='primary',use_container_width=True)
        with st.expander('其他方式'):
            st.file_uploader('沿用已有 Memo（选填）',type=['pptx'],key='source_upload',
                help='以此文件为底稿，保留其他手工调整。')
            manual=st.button('跳过识别，手动填写',use_container_width=True)
        if recognize or manual:
            st.session_state.pop('output',None)
            if not festival.strip():
                required_message.error('请先填写节日 / 档期。')
                st.session_state.review=False
            elif recognize and not wording.strip() and not image and source is None:
                st.warning('请上传 A1、粘贴文案，或选择手动填写。')
                st.session_state.review=False
            else:
                state=dict(st.session_state.memo)
                state.pop('_meta',None)
                try: check_source_brand(brand,wording,a1.name if a1 else '')
                except BrandMismatch as exc:
                    st.error(str(exc));st.session_state.review=False;return
                image_read=False
                warning=''
                if recognize and api_key() and (wording.strip() or image):
                    try:
                        status=st.empty()
                        try:
                            with st.spinner('正在识别素材…'):
                                state=ai_parse(api_key(),brand,state,wording,image,a1.type if a1 else 'image/jpeg',
                                               on_status=lambda message:status.info(message))
                        finally:status.empty()
                        image_read=bool(image)
                        st.session_state.parse_note='已提取素材信息。请核对日期、优惠和门店范围。'
                    except BrandMismatch as exc:
                        st.error(str(exc));st.session_state.review=False;return
                    except RecognitionError as exc:
                        warning=str(exc)+(' A1 尚未识别，当前没有自动更新海报内容。' if image else '')
                    except Exception:
                        warning='识别流程出现异常，请联系维护人员检查部署版本。'
                elif recognize and (wording.strip() or image):
                    warning='未配置智能识别。'+('A1 尚未识别，请手动核对日期、优惠和门店范围。' if image else '仅提取明确日期和标注字段。')
                else:
                    st.session_state.parse_note='请核对需要修改的字段。'
                if manual or warning or not (wording.strip() or image):
                    try: state=local_parse(wording,state)
                    except ValueError:
                        st.error('文案日期无效，请检查后重试。');return
                state['festival']=festival.strip()
                st.session_state.memo=state
                st.session_state.review=True
                st.session_state.review_input=(wording,image)
                st.session_state.parse_warning=warning
                st.session_state.image_unread=bool(image) and not image_read
                st.session_state.outlet_identified='exclude_outlets' in state.get('_meta',{}).get('recognized_fields',[])
                reset_editor()
    if not st.session_state.get('review'): return
    if st.session_state.get('review_input') != (wording,image):
        st.session_state.pop('output',None)
        st.warning('素材已变化，请重新识别；也可选择手动填写。');return
    with st.container(border=True):
        st.subheader('2 · 核对并生成')
        if st.session_state.get('parse_warning'): st.warning(st.session_state.parse_warning)
        else: st.caption(st.session_state.get('parse_note',''))
        state=dict(st.session_state.memo)
        state['festival']=festival.strip()
        a,b=st.columns(2)
        with a:
            value=st.date_input('开始日期',value=date_value(state['start_date']),key='edit_start_date')
            state['start_date']=value.isoformat() if value else ''
        with b:
            value=st.date_input('结束日期',value=date_value(state['end_date']),key='edit_end_date')
            state['end_date']=value.isoformat() if value else ''
        state['short_promo']=st.text_input('主要优惠',value=state['short_promo'],key='edit_short_promo')
        state['store_scope']=st.text_input('参与门店',value=state['store_scope'],key='edit_store_scope')
        state['exclude_outlets']=st.checkbox('奥莱不参与（识别结果，可调整）',value=state['exclude_outlets'],key='edit_exclude_outlets')
        if st.session_state.get('outlet_identified'):
            evidence=state.get('_meta',{}).get('outlet_evidence','')
            st.caption('已从素材识别奥莱范围'+('：'+evidence if evidence else '，此处仅供确认。'))
        else:
            st.caption('素材中未识别到明确奥莱规则；当前保留'+('已有 Memo' if source else '品牌标准版')+'的设置，请确认。')
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
        if image: st.caption('生成时同步更新活动字段并替换两处 A1，保持图片比例。')
        elif source: st.caption('保留上传 Memo 中的海报。')
        else: st.caption('尚无 A1，将使用“海报待补充”占位。')
        image_confirmed=True
        if st.session_state.get('image_unread'):
            image_confirmed=st.checkbox('我已手动核对这张海报的日期、优惠与门店范围',key='edit_image_confirmed')
        if st.session_state.get('output_state') != state:
            st.session_state.pop('output',None)
        if st.button('确认并生成 PPT',type='primary',use_container_width=True):
            st.session_state.pop('output',None)
            errors=validate_state(state)
            if not image_confirmed: errors.append('本次海报尚未识别，请先手动核对并确认。')
            if errors:
                for error in errors: st.error(error)
            else:
                progress=st.progress(0,text='正在准备生成…')
                try:
                    def update(value,message): progress.progress(value,text=message)
                    data,count=generate_memo(brand,state,source,st.session_state.baseline,image,progress=update)
                    st.session_state.output=data
                    st.session_state.output_state=dict(state)
                    st.session_state.output_count=count
                except Exception as exc: st.error('生成失败：'+str(exc))
                finally: progress.empty()
        if st.session_state.get('output'):
            data=st.session_state.output
            name=state['start_date'].replace('-','')+' '+BRAND_CONFIG[brand]['code']+' promo memo.pptx'
            # Clicking download must not rerun widgets and invalidate the media URL.
            st.download_button('下载 PPT · '+name,data=data,file_name=name,key='download_memo',on_click='ignore',
                mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',use_container_width=True)
            st.success(f'文件已校验并就绪（{len(data)/1024/1024:.1f} MB），请点击上方按钮下载。')


if __name__=='__main__': main()
