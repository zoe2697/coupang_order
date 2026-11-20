"""쿠팡 매입 확인 — pip install streamlit pandas openpyxl; streamlit run app_coupang_purchase.py"""
import io
from collections import defaultdict
import pandas as pd
import streamlit as st

st.set_page_config(page_title='쿠팡 매입 확인',page_icon='📦',layout='wide')
st.title('쿠팡 매입 확인')
st.caption('통합 서식 1개 업로드 · 원본 RAW 모든 행 합산 · 센터 및 SKU 매핑 검증 · 이력 저장 없음')

@st.cache_data(show_spinner=False)
def load(raw):
    xl=pd.ExcelFile(io.BytesIO(raw))
    required=['RAW','Sales Report (Coupang)','센터 참조','상품매핑관리']
    missing=set(required)-set(xl.sheet_names)
    if missing: raise ValueError('필수 시트 누락: '+', '.join(sorted(missing)))
    raw_df=pd.read_excel(xl,'RAW',dtype={'SKU번호':str,'물류센터':str})
    sales=pd.read_excel(xl,'Sales Report (Coupang)',dtype={'제품코드':str,'점포':str})
    centers=pd.read_excel(xl,'센터 참조',header=None,usecols='A:B',names=['센터코드','센터명'],dtype=str)
    products=pd.read_excel(xl,'상품매핑관리',dtype={'SKU ID':str,'ME코드':str})
    return raw_df,sales,centers,products

def clean(v):
    if pd.isna(v): return ''
    s=str(v).strip()
    return s[:-2] if s.endswith('.0') and s[:-2].isdigit() else s

def num(s): return pd.to_numeric(s,errors='coerce').fillna(0)

file=st.file_uploader('업데이트한 통합 엑셀 파일 업로드',type=['xlsx'])
if file is None:
    st.info('RAW / Sales Report (Coupang) / 센터 참조 / 상품매핑관리 시트가 있는 파일을 업로드해 주세요.')
    st.stop()
try:
    raw,sales,centers,products=load(file.getvalue())
    for sheet,df,cols in [('RAW',raw,['SKU번호','물류센터','입고/반출시각','수량','총공급가액']),('Sales Report',sales,['Date','점포','제품코드','수량','Total Amount']),('상품매핑관리',products,['SKU ID','ME코드'])]:
        absent=set(cols)-set(df.columns)
        if absent:raise ValueError(f'{sheet}: 열 누락 {sorted(absent)}')
except Exception as e:
    st.error(f'파일 구조를 확인해 주세요: {e}')
    st.stop()

# Remove empty template rows, not duplicate transactions.
raw=raw[raw['SKU번호'].notna() & raw['물류센터'].notna()].copy()
sales=sales[sales['제품코드'].notna() & sales['점포'].notna()].copy()
raw['SKU번호']=raw['SKU번호'].map(clean)
raw['센터코드']=raw['물류센터'].map(clean)
raw['입고일']=pd.to_datetime(raw['입고/반출시각'],errors='coerce').dt.normalize()
raw['입고수량']=num(raw['수량'])
raw['입고금액']=num(raw['총공급가액'])
sales['ME코드']=sales['제품코드'].map(clean)
sales['센터명']=sales['점포'].map(clean)
sales['발주수량']=num(sales['수량'])
sales['발주금액']=num(sales['Total Amount'])
sales['예정일']=pd.to_datetime(sales['Date'],errors='coerce').dt.normalize()

centers['센터코드']=centers['센터코드'].map(clean)
centers['센터명']=centers['센터명'].map(clean)
centers=centers[(centers['센터코드']!='')&(centers['센터명']!='')]
center_map=centers.groupby('센터코드')['센터명'].agg(lambda s:sorted(set(s))).to_dict()
products['SKU ID']=products['SKU ID'].map(clean)
products['ME코드']=products['ME코드'].map(clean)
products=products[products['SKU ID'].str.fullmatch(r'\d+',na=False) & products['ME코드'].str.startswith('ME',na=False)]
product_map=products.groupby('SKU ID')['ME코드'].agg(lambda s:sorted(set(s))).to_dict()
raw['센터후보']=raw['센터코드'].map(lambda x:center_map.get(x,[]))
raw['ME후보']=raw['SKU번호'].map(lambda x:product_map.get(x,[]))
raw['센터명']=raw['센터후보'].map(lambda a:a[0] if len(a)==1 else '')
raw['ME코드']=raw['ME후보'].map(lambda a:a[0] if len(a)==1 else '')
raw['매핑상태']=raw.apply(lambda r:'센터 미등록' if not r['센터후보'] else '센터 중복 매핑' if len(r['센터후보'])>1 else 'SKU 미등록' if not r['ME후보'] else 'SKU 중복 매핑' if len(r['ME후보'])>1 else '정상',axis=1)

st.sidebar.header('입고 인정 기간')
unique_dates=sorted(sales['예정일'].dropna().unique())
if not unique_dates:
    st.error('Sales Report의 Date 열에서 예정 입고일을 찾지 못했습니다.')
    st.stop()
chosen=st.sidebar.selectbox('확인할 예정 입고일',unique_dates,format_func=lambda x:pd.Timestamp(x).strftime('%Y-%m-%d'))
next_dates=[x for x in unique_dates if x>chosen]
if next_dates:
    default_end=pd.Timestamp(next_dates[0])-pd.Timedelta(days=1)
    st.sidebar.caption('다음 예정 입고일 전날까지 자동 적용')
else:
    default_end=raw['입고일'].max()
    if pd.isna(default_end) or default_end<pd.Timestamp(chosen):default_end=pd.Timestamp(chosen)
    st.sidebar.caption('다음 발주일이 파일에 없어 종료일을 직접 확인해 주세요.')
end=st.sidebar.date_input('입고 인정 종료일',value=default_end.date(),min_value=pd.Timestamp(chosen).date())
start=pd.Timestamp(chosen)
end_ts=pd.Timestamp(end)
selected_sales=sales[sales['예정일']==chosen].copy()
selected_raw=raw[raw['입고일'].between(start,end_ts,inclusive='both')].copy()
st.caption(f'발주 예정일 {start:%Y-%m-%d} | 입고 인정 기간 {start:%Y-%m-%d} ~ {end_ts:%Y-%m-%d} | RAW {len(selected_raw):,}행 (중복 제거 없음)')

# Aggregate independently. Unmapped records are excluded from matched totals, shown as exceptions.
valid=selected_raw[selected_raw['매핑상태']=='정상']
sales_group=selected_sales.groupby(['센터명','ME코드'],as_index=False).agg(발주수량=('발주수량','sum'),발주금액=('발주금액','sum'),품목명=('제품명','first'))
raw_group=valid.groupby(['센터명','ME코드'],as_index=False).agg(입고수량=('입고수량','sum'),입고금액=('입고금액','sum'),입고건수=('SKU번호','size'))
merged=sales_group.merge(raw_group,on=['센터명','ME코드'],how='outer',indicator=True)
merged['품목명']=merged['품목명'].fillna('발주목록에 없는 입고 품목')
for c in ['발주수량','발주금액','입고수량','입고금액','입고건수']:merged[c]=merged[c].fillna(0)
merged['수량차이']=merged['입고수량']-merged['발주수량']
merged['금액차이']=merged['입고금액']-merged['발주금액']
def status(r):
    if r['_merge']=='right_only':return '발주 외 입고'
    if r['입고수량']==0:return '미입고'
    if r['입고수량']<r['발주수량']:return '부분 입고'
    if r['입고수량']>r['발주수량']:return '초과 입고'
    if abs(r['금액차이'])>0.5:return '금액 확인 필요'
    return '매입 완료'
merged['매입상태']=merged.apply(status,axis=1)
# A matching ambiguity can affect a seemingly completed row: warn that results are provisional.
exceptions=selected_raw[selected_raw['매핑상태']!='정상'].copy()
cols=['센터명','ME코드','품목명','발주수량','입고수량','수량차이','발주금액','입고금액','금액차이','입고건수','매입상태']
merged=merged[cols].sort_values(['매입상태','센터명','ME코드']).reset_index(drop=True)

if not exceptions.empty:
    st.warning(f'매핑 확인이 필요한 RAW {len(exceptions):,}행이 있습니다. 해당 입고분은 매칭 집계에서 제외되어 결과가 잠정적입니다.')
raw_total=selected_raw['입고금액'].sum()
matched_total=valid['입고금액'].sum()
col=st.columns(4)
col[0].metric('발주 총금액',f"{selected_sales['발주금액'].sum():,.0f}원")
col[1].metric('실제 입고 총금액',f'{raw_total:,.0f}원')
col[2].metric('총금액 차이',f"{raw_total-selected_sales['발주금액'].sum():+,.0f}원")
col[3].metric('매입 완료 / 비교 항목',f"{(merged['매입상태']=='매입 완료').sum():,} / {len(merged):,}")
st.caption(f'매핑 완료 입고금액 {matched_total:,.0f}원 | 매핑 확인 필요 입고금액 {raw_total-matched_total:,.0f}원. 총액 일치는 개별 매입 완료를 의미하지 않습니다.')

t1,t2,t3,t4=st.tabs(['① 전체 매입 현황','② 센터별 현황','③ 품목별 상세','④ 매칭 확인 필요'])
with t1:
    st.subheader('매입 상태별 건수')
    st.dataframe(merged['매입상태'].value_counts().rename_axis('상태').reset_index(name='품목·센터 조합 수'),hide_index=True,use_container_width=True)
    st.subheader('확인 필요 항목')
    st.dataframe(merged[merged['매입상태']!='매입 완료'][cols],hide_index=True,use_container_width=True)
with t2:
    st.subheader('센터별 발주·입고 비교')
    center=merged.groupby('센터명',as_index=False).agg(발주수량=('발주수량','sum'),입고수량=('입고수량','sum'),발주금액=('발주금액','sum'),입고금액=('입고금액','sum'),확인필요=('매입상태',lambda s:(s!='매입 완료').sum()))
    center['수량차이']=center['입고수량']-center['발주수량']
    center['금액차이']=center['입고금액']-center['발주금액']
    st.dataframe(center,hide_index=True,use_container_width=True)
with t3:
    st.subheader('센터·ME코드별 매입 대조')
    choices=['전체']+sorted(merged['매입상태'].unique().tolist())
    selected=st.selectbox('매입 상태',choices)
    shown=merged if selected=='전체' else merged[merged['매입상태']==selected]
    st.dataframe(shown,hide_index=True,use_container_width=True)
with t4:
    st.subheader('센터/SKU 매핑 검증')
    st.caption('상품매핑관리의 SKU ID가 여러 ME코드에 연결된 경우 임의로 선택하지 않습니다. 센터 참조에 동일 영문코드가 서로 다른 센터명으로 등록되어도 확인 대상으로 표시합니다.')
    if exceptions.empty:st.success('선택 기간 RAW의 센터·SKU 매핑 오류가 없습니다.')
    else:
        view=exceptions[['입고/반출시각','센터코드','SKU번호','SKU명','입고수량','입고금액','매핑상태']].copy()
        view['센터 후보']=exceptions['센터후보'].map(lambda a:', '.join(a))
        view['ME 후보']=exceptions['ME후보'].map(lambda a:', '.join(a))
        st.dataframe(view,hide_index=True,use_container_width=True)
    duplicated_centers={k:v for k,v in center_map.items() if len(v)>1}
    if duplicated_centers:st.warning('센터 참조 내 중복 매핑: '+str(duplicated_centers))

buf=io.BytesIO()
with pd.ExcelWriter(buf,engine='openpyxl') as writer:
    merged.to_excel(writer,sheet_name='품목별 매입 확인',index=False)
    center.to_excel(writer,sheet_name='센터별 현황',index=False)
    if not exceptions.empty:
        exceptions[['입고/반출시각','센터코드','SKU번호','SKU명','입고수량','입고금액','매핑상태']].to_excel(writer,sheet_name='매핑 확인 필요',index=False)
st.download_button('매입 확인 결과 엑셀 다운로드',data=buf.getvalue(),file_name=f'쿠팡_매입확인_{start:%Y%m%d}.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
