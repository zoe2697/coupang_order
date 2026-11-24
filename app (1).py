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

def center_key(v):
    # RAW may use min34 / MIN34 while the reference uses another case.
    return ''.join(clean(v).split()).casefold()


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
raw['_raw_id']=range(len(raw))
raw['SKU번호']=raw['SKU번호'].map(clean)
raw['센터코드']=raw['물류센터'].map(clean)
raw['센터키']=raw['센터코드'].map(center_key)
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
centers['센터키']=centers['센터코드'].map(center_key)
centers=centers[(centers['센터코드']!='')&(centers['센터명']!='')]
center_map=centers.groupby('센터키')['센터명'].agg(lambda s:sorted(set(s))).to_dict()
products['SKU ID']=products['SKU ID'].map(clean)
products['ME코드']=products['ME코드'].map(clean)
products=products[products['SKU ID'].str.fullmatch(r'\d+',na=False) & products['ME코드'].str.startswith('ME',na=False)]
product_map=products.groupby('SKU ID')['ME코드'].agg(lambda s:sorted(set(s))).to_dict()
raw['센터후보']=raw['센터키'].map(lambda x:center_map.get(x,[]))
raw['ME후보']=raw['SKU번호'].map(lambda x:product_map.get(x,[]))
raw['센터명']=raw['센터후보'].map(lambda a:a[0] if len(a)==1 else '')
raw['ME코드']=raw['ME후보'].map(lambda a:a[0] if len(a)==1 else '')
raw['매핑상태']=raw.apply(lambda r:'센터 미등록' if not r['센터후보'] else '센터 중복 매핑' if len(r['센터후보'])>1 else 'SKU 미등록' if not r['ME후보'] else 'SKU 중복 매핑' if len(r['ME후보'])>1 else '정상',axis=1)


import re

def norm(v):
    return re.sub(r'\s+', '', clean(v)).casefold()

def dates_between(df, col, start, end):
    return df[df[col].between(pd.Timestamp(start), pd.Timestamp(end), inclusive='both')].copy()

st.sidebar.header('기간별 일괄 매입확인')
order_dates=sales['예정일'].dropna()
receipt_dates=raw['입고일'].dropna()
if order_dates.empty or receipt_dates.empty:
    st.error('입고예정일 또는 RAW 실제 입고일이 없습니다. 원본 날짜 열을 확인해 주세요.')
    st.stop()
order_start=st.sidebar.date_input('입고예정일 시작', value=order_dates.min().date(), key='os')
order_end=st.sidebar.date_input('입고예정일 종료 (포함)', value=order_dates.max().date(), key='oe')
receipt_start=order_start
receipt_end=receipt_dates.max().date()
st.sidebar.caption(f'RAW 실제 입고기간 (자동): {receipt_start} ~ {receipt_end}')
if order_start>order_end:
    st.error('입고예정일 시작일이 종료일보다 늦을 수 없습니다.')
    st.stop()
selected_sales=dates_between(sales,'예정일',order_start,order_end)
selected_raw=dates_between(raw,'입고일',receipt_start,receipt_end)
if receipt_start>receipt_end:
    st.warning('RAW 데이터의 마지막 실제 입고일이 선택한 입고예정일 시작일보다 빠릅니다. RAW 조회 결과가 비어 있습니다.')

# 센터별 입고예정일을 기준으로 실제 입고일을 가장 가까운 이전 예정일에 연결합니다.
# 선택 범위 밖의 다음 예정일도 경계로 사용하여 잘못된 배분을 방지합니다.
selected_sales['상품명키']=selected_sales.get('제품명',pd.Series('',index=selected_sales.index)).map(norm)
selected_sales=selected_sales.reset_index(drop=True)
selected_sales['연결SKU']=selected_sales['ME코드'].map(lambda me:set(products.loc[products['ME코드']==me,'SKU ID']))
parent=list(range(len(selected_sales)))
def find(i):
    while parent[i]!=i:
        parent[i]=parent[parent[i]]
        i=parent[i]
    return i
def union(i,j):
    a,b=find(i),find(j)
    if a!=b:parent[b]=a
# 같은 센터, 같은 입고예정일, 같은 품목만 하나의 발주 그룹으로 합칩니다.
for (center,planned_date),indices in selected_sales.groupby(['센터명','예정일']).groups.items():
    indices=list(indices)
    for pos,i in enumerate(indices):
        for j in indices[pos+1:]:
            a,b=selected_sales.loc[i],selected_sales.loc[j]
            if (a['상품명키'] and a['상품명키']==b['상품명키']) or (a['연결SKU'] & b['연결SKU']):
                union(i,j)
selected_sales['그룹']=[find(i) for i in range(len(selected_sales))]
items=[]
for gid,g in selected_sales.groupby('그룹',sort=False):
    items.append(dict(센터명=g['센터명'].iloc[0],ME코드=', '.join(sorted(set(g['ME코드']))),
        품목명=' / '.join(dict.fromkeys(g.get('제품명',pd.Series('',index=g.index)).map(clean))),
        입고예정일=g['예정일'].iloc[0],
        발주수량=g['발주수량'].sum(),발주금액=g['발주금액'].sum(),
        입고수량=0.,입고금액=0.,SKU집합=set().union(*g['연결SKU'].tolist()),그룹=gid))
lookup={}
for item in items:
    for sku in item['SKU집합']:
        lookup.setdefault((item['센터명'],item['입고예정일'],sku),[]).append(item)
# 센터별 전체 Sales Report 날짜를 경계로 사용합니다. (선택 범위 내 날짜만 쓰면
# 다음 발주일 이후 입고가 이전 발주에 잘못 연결될 수 있습니다.)
center_dates={c:sorted(set(g['예정일'].dropna())) for c,g in sales.groupby('센터명')}
from bisect import bisect_right
receipt_rows=[]
for _,r in selected_raw.iterrows():
    row=r.to_dict();candidates=[];target_date=pd.NaT
    if len(r['센터후보'])==1 and pd.notna(r['입고일']):
        center=r['센터후보'][0]
        dates=center_dates.get(center,[])
        idx=bisect_right(dates,r['입고일'])-1
        if idx>=0:
            target_date=dates[idx]
            candidates=lookup.get((center,target_date,r['SKU번호']),[])
            if not candidates:
                name=norm(r.get('SKU명',''))
                if name:
                    candidates=[it for it in items if it['센터명']==center and it['입고예정일']==target_date
                                and name in [norm(n) for n in it['품목명'].split(' / ')]]
    row['연결입고예정일']=target_date
    if len(candidates)==1:
        it=candidates[0]
        it['입고수량']+=r['입고수량'];it['입고금액']+=r['입고금액']
        row['연결그룹']=it['그룹'];row['매칭상태']='발주 연결'
    else:
        row['연결그룹']=None
        row['매칭상태']=('센터 매핑 확인 필요' if len(r['센터후보'])!=1 else
                       '복수 발주 후보' if len(candidates)>1 else
                       '연결 입고예정일 없음' if pd.isna(target_date) else '발주 미연결')
    receipt_rows.append(row)
receipt=pd.DataFrame(receipt_rows)
records=[]
for it in items:
    rec={k:v for k,v in it.items() if k!='SKU집합'}
    rec['수량차이']=rec['입고수량']-rec['발주수량']
    rec['금액차이']=rec['입고금액']-rec['발주금액']
    if rec['입고수량']==0:status='미입고'
    elif rec['입고수량']<rec['발주수량']:status='부분 입고'
    elif rec['입고수량']>rec['발주수량']:status='초과 입고'
    elif abs(rec['금액차이'])>0.5:status='금액 확인 필요'
    else:status='매입 완료'
    rec['매입상태']=status
    records.append(rec)
columns=['센터명','입고예정일','ME코드','품목명','발주수량','입고수량','수량차이','발주금액','입고금액','금액차이','매입상태','그룹']
merged=pd.DataFrame(records,columns=columns)
if not receipt.empty:
    status_map={r['그룹']:r['매입상태'] for r in records}
    receipt['품목매입상태']=receipt['연결그룹'].map(status_map).fillna('발주 미연결')
else:
    receipt=pd.DataFrame(columns=list(selected_raw.columns)+['연결입고예정일','연결그룹','매칭상태','품목매입상태'])

st.caption(f'입고예정일 {order_start} ~ {order_end} / RAW 실제 입고일 {receipt_start} ~ {receipt_end} (자동) · 센터별 가장 가까운 이전 입고예정일로 RAW 연결')
centers_available=sorted(set(selected_sales['센터명'].dropna()) | set(selected_raw['센터명'].dropna()))
center=st.selectbox('센터 선택',['전체 센터']+centers_available)
orders_view=merged if center=='전체 센터' else merged[merged['센터명']==center]
raw_view=receipt if center=='전체 센터' else receipt[receipt['센터명']==center]
order_sum=orders_view['발주금액'].sum()
raw_sum=raw_view['입고금액'].sum()
matched_sum=raw_view.loc[raw_view['매칭상태']=='발주 연결','입고금액'].sum()
unmatched_sum=raw_sum-matched_sum
m1,m2,m3,m4=st.columns(4)
m1.metric('기간 발주 총금액',f'{order_sum:,.0f}원')
m2.metric('기간 RAW 전체 입고금액',f'{raw_sum:,.0f}원')
m3.metric('전체 금액 차이 (RAW − 발주)',f'{raw_sum-order_sum:+,.0f}원')
m4.metric('발주 연결 입고금액',f'{matched_sum:,.0f}원')
st.caption(f'발주 미연결 RAW 금액 {unmatched_sum:,.0f}원 · 기간 차이에는 이전 발주 지연 입고, 반출, 다른 발주 건이 포함될 수 있으므로 품목별 대조가 필요합니다.')

t1,t2,t3=st.tabs(['① 기간별 매입 현황','② 센터별 발주 · RAW 상세','③ 매핑 검증'])
with t1:
    st.subheader('품목별 발주 · 입고 차이')
    statuses=sorted(orders_view['매입상태'].dropna().unique())
    status_filter=st.multiselect('매입상태 필터',statuses,default=[s for s in statuses if s!='매입 완료'])
    st.dataframe(orders_view[orders_view['매입상태'].isin(status_filter)].drop(columns=['그룹']),hide_index=True,use_container_width=True)
with t2:
    st.subheader('① Sales Report 발주목록')
    st.metric('발주 총금액',f'{order_sum:,.0f}원')
    date_source=selected_sales if center=='전체 센터' else selected_sales[selected_sales['센터명']==center]
    date_options=sorted(date_source['예정일'].dt.date.dropna().unique())
    selected_order_dates=st.multiselect('입고예정일 필터',date_options,default=date_options)
    # 발주 날짜별 원본을 별도로 표시하되 기간 전체 매입상태를 유지합니다.
    order_lines=selected_sales[selected_sales['예정일'].dt.date.isin(selected_order_dates)].copy()
    if center!='전체 센터':order_lines=order_lines[order_lines['센터명']==center]
    st.dataframe(order_lines[['예정일','센터명','ME코드','제품명','발주수량','발주금액']].rename(columns={'예정일':'입고예정일','제품명':'품목명'}),hide_index=True,use_container_width=True)
    st.caption(f'표시 발주금액 {order_lines["발주금액"].sum():,.0f}원 / 선택 센터 기간 전체 {order_sum:,.0f}원')
    st.subheader('② RAW 실제 입고 상세내역')
    st.metric('기간 RAW 전체 입고금액',f'{raw_sum:,.0f}원')
    date_options_raw=sorted(raw_view['입고일'].dt.date.dropna().unique())
    selected_raw_dates=st.multiselect('RAW 실제 입고일 필터',date_options_raw,default=date_options_raw)
    filtered=raw_view[raw_view['입고일'].dt.date.isin(selected_raw_dates)].copy()
    view_mode=st.radio('RAW 표시 범위',['특이사항만','전체 입고내역'],horizontal=True)
    st.caption(f'선택 날짜 RAW 전체 금액 {raw_view.loc[raw_view["입고일"].dt.date.isin(selected_raw_dates),"입고금액"].sum():,.0f}원')
    # 표시 목적의 일자별 합산. 판정은 연결된 입고예정일의 전체 RAW 수량으로 수행합니다.
    if not filtered.empty:
        filtered['입고일']=pd.to_datetime(filtered['입고/반출시각'],errors='coerce').dt.normalize()
        group_cols=['입고일','센터명','SKU명','ME코드','연결입고예정일','연결그룹']
        def unique_join(values):
            return ', '.join(dict.fromkeys(x for x in values.map(clean) if x))
        grouped=filtered.groupby(group_cols,dropna=False,as_index=False).agg(
            센터코드=('센터코드',unique_join),SKU번호=('SKU번호',unique_join),
            매칭상태=('매칭상태',unique_join),
            입고수량=('입고수량','sum'),입고금액=('입고금액','sum'),원본행수=('_raw_id','size'))
        status_map={r['그룹']:r['매입상태'] for r in records}
        grouped['품목매입상태']=grouped['연결그룹'].map(status_map).fillna('발주 미연결')
        grouped['입고일']=grouped['입고일'].dt.date
        grouped['연결입고예정일']=pd.to_datetime(grouped['연결입고예정일']).dt.date
        grouped=grouped[['입고일','연결입고예정일','센터명','센터코드','SKU번호','SKU명','ME코드','입고수량','입고금액','원본행수','매칭상태','품목매입상태']]
        grouped=grouped.sort_values(['입고일','센터명','SKU명','ME코드'])
    else:
        grouped=pd.DataFrame(columns=['입고일','연결입고예정일','센터명','센터코드','SKU번호','SKU명','ME코드','입고수량','입고금액','원본행수','매칭상태','품목매입상태'])
    if view_mode=='특이사항만':
        grouped=grouped[grouped['품목매입상태']!='매입 완료'].copy()
    st.caption(f'표시 RAW 금액 {grouped["입고금액"].sum():,.0f}원')
    st.dataframe(grouped,hide_index=True,use_container_width=True)
    st.caption('동일 센터·입고일·SKU명·ME코드의 RAW를 합산 표시합니다. 매입상태는 같은 센터의 가장 가까운 이전 입고예정일에 연결된 전체 입고수량·금액으로 판단합니다. RAW 표시일 필터는 매입상태 계산에 영향을 주지 않습니다.')
with t3:
    st.subheader('발주 미연결 / 센터·SKU 매핑 확인')
    issues=raw_view[raw_view['매칭상태']!='발주 연결'].copy()
    st.metric('미연결 RAW 입고금액',f'{issues["입고금액"].sum():,.0f}원')
    st.dataframe(issues[['입고/반출시각','센터코드','센터명','SKU번호','SKU명','입고수량','입고금액','매칭상태']],hide_index=True,use_container_width=True)
    st.caption('입고가 발주 조회기간보다 이전 발주에 속하는 경우도 미연결로 나타날 수 있습니다. 이를 곧바로 발주 외 입고로 확정하지 않습니다.')

buf=io.BytesIO()
with pd.ExcelWriter(buf,engine='openpyxl') as writer:
    orders_view.drop(columns=['그룹']).to_excel(writer,sheet_name='기간별 품목 대조',index=False)
    selected_sales[['예정일','센터명','ME코드','제품명','발주수량','발주금액']].to_excel(writer,sheet_name='Sales Report 발주',index=False)
    receipt[['입고/반출시각','연결입고예정일','센터명','센터코드','SKU번호','SKU명','ME코드','입고수량','입고금액','매칭상태','품목매입상태']].to_excel(writer,sheet_name='RAW 원본 대조',index=False)
    raw_view[raw_view['매칭상태']!='발주 연결'][['입고/반출시각','센터코드','SKU번호','SKU명','입고수량','입고금액','매칭상태']].to_excel(writer,sheet_name='매핑 확인 필요',index=False)
st.download_button('매입 확인 결과 엑셀 다운로드',data=buf.getvalue(),file_name=f'쿠팡_매입확인_{order_start}_{order_end}.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
