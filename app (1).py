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

st.sidebar.header('매입확인 범위')
unique_dates=sorted(pd.Timestamp(x) for x in sales['예정일'].dropna().unique())
if not unique_dates:
    st.error('Sales Report의 Date 열에서 예정 입고일을 찾지 못했습니다.')
    st.stop()
mode=st.sidebar.radio('확인 방식', ['기간별 일괄 확인', '입고예정일별 확인'])
if mode=='입고예정일별 확인':
    chosen=st.sidebar.selectbox('확인할 예정 입고일',unique_dates,format_func=lambda x:x.strftime('%Y-%m-%d'))
    target_dates=[chosen]
else:
    first=st.sidebar.date_input('발주 예정일 시작',value=unique_dates[0].date())
    last=st.sidebar.date_input('발주 예정일 종료 (포함)',value=min(pd.Timestamp.today().normalize()-pd.Timedelta(days=1),unique_dates[-1]).date())
    if first>last:
        st.error('시작일은 종료일보다 늦을 수 없습니다.')
        st.stop()
    target_dates=[d for d in unique_dates if pd.Timestamp(first)<=d<=pd.Timestamp(last)]
    if not target_dates:
        st.warning('선택 기간에 발주 예정일이 없습니다.')
        st.stop()

# Periods are determined from ALL sales dates, including dates outside the
# selected range, so the final selected order cannot consume later receipts.
periods=[]
for d in target_dates:
    next_dates=[x for x in unique_dates if x>d]
    if next_dates:
        period_end=next_dates[0]-pd.Timedelta(days=1)
    else:
        period_end=raw['입고일'].max()
        if pd.isna(period_end) or period_end<d: period_end=d
        # Do not assign receipts beyond the selected cutoff when no later
        # order date exists in the uploaded sales report.
        if mode=='기간별 일괄 확인':
            period_end=min(period_end,pd.Timestamp(last))
    periods.append((d,period_end))
if mode=='입고예정일별 확인':
    pass  # 다음 예정일 전날까지 자동 적용
st.caption(f'{mode} · 발주 예정일 {target_dates[0]:%Y-%m-%d} ~ {target_dates[-1]:%Y-%m-%d} · {len(periods)}개 입고예정일 (RAW 중복 제거 없음)')
if mode=='기간별 일괄 확인' and target_dates[-1]==unique_dates[-1]:
    st.info('마지막 발주 예정일 이후의 다음 발주일이 파일에 없으므로, 해당 건의 입고 인정 종료일은 선택한 종료일로 제한합니다. 다음 발주일을 포함한 Sales Report를 업로드하면 기간이 더 정확해집니다.')

# Match within each center by SKU-linked product identity.  Multiple ME codes
# for one SKU are not an error when the sales rows refer to the same item.
# Exact normalized sales product names also join differently coded sales rows.
import re

def normalized_name(value):
    return re.sub(r"\s+", "", clean(value)).casefold()

def reconcile_period(selected_sales,selected_raw,order_date,period_end):
    # Each center has independent item groups; never combine across centers.
    sales_rows = selected_sales.reset_index(drop=True).copy()
    sales_rows['상품명키'] = sales_rows['제품명'].map(normalized_name)
    sales_rows['연결SKU'] = sales_rows['ME코드'].map(lambda me: set(products.loc[products['ME코드']==me, 'SKU ID']))
    parent = list(range(len(sales_rows)))
    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i
    def union(i,j):
        ri,rj=find(i),find(j)
        if ri!=rj: parent[rj]=ri

    for center_name, indexes in sales_rows.groupby('센터명').groups.items():
        indexes=list(indexes)
        for n,i in enumerate(indexes):
            for j in indexes[n+1:]:
                si,sj=sales_rows.loc[i],sales_rows.loc[j]
                same_name=bool(si['상품명키']) and si['상품명키']==sj['상품명키']
                same_sku=bool(si['연결SKU'] & sj['연결SKU'])
                if same_name or same_sku:
                    union(i,j)

    sales_rows['그룹']= [find(i) for i in range(len(sales_rows))]
    items=[]
    for gid, group in sales_rows.groupby('그룹',sort=False):
        items.append({'그룹':gid,'센터명':group['센터명'].iloc[0],
            'ME코드':', '.join(sorted(set(group['ME코드']))),
            '품목명':' / '.join(dict.fromkeys(group['제품명'].map(clean))),
            '발주수량':group['발주수량'].sum(), '발주금액':group['발주금액'].sum(),
            '입고수량':0.,'입고금액':0.,'입고건수':0,
            'SKU집합':set().union(*group['연결SKU'].tolist())})

    # Raw SKU -> a unique ordered item group in the same center. All raw rows
    # count, including identical SKU/quantity/timestamp rows.
    lookup={}
    for item in items:
        for sku in item['SKU집합']:
            lookup.setdefault((item['센터명'],sku),[]).append(item)

    exceptions_rows=[]
    extra={}
    for _,r in selected_raw.iterrows():
        if len(r['센터후보'])!=1:
            exceptions_rows.append({**r.to_dict(),'매핑상태':'센터 미등록' if not r['센터후보'] else '센터 중복 매핑'})
            continue
        center_name=r['센터후보'][0]
        candidates=lookup.get((center_name,r['SKU번호']),[])
        # Some reference files have SKU/ME mapping omissions. Match exact
        # normalized product names only when they identify ONE ordered item.
        # Never use fuzzy/substring matching or force ambiguous matches.
        if not candidates:
            raw_name=normalized_name(r.get('SKU명',''))
            if raw_name:
                candidates=[item for item in items if item['센터명']==center_name and
                    raw_name in [normalized_name(n) for n in item['품목명'].split(' / ')]]

        if len(candidates)==1:
            item=candidates[0]
            item['입고수량']+=r['입고수량']
            item['입고금액']+=r['입고금액']
            item['입고건수']+=1
        elif len(candidates)>1:
            exceptions_rows.append({**r.to_dict(),'매핑상태':'동일 SKU에 여러 발주 품목 연결'})
        elif len(r['ME후보'])>1:
            # A multi-code SKU is valid only if the relevant center has a unique
            # matching order group; otherwise it remains an unmatched raw receipt.
            k=(center_name,r['SKU번호'])
            if k not in extra: extra[k]={'센터명':center_name,'ME코드':', '.join(r['ME후보']),
                 '품목명':clean(r['SKU명']),'발주수량':0,'발주금액':0,'입고수량':0.,'입고금액':0.,'입고건수':0}
            extra[k]['입고수량']+=r['입고수량'];extra[k]['입고금액']+=r['입고금액'];extra[k]['입고건수']+=1
        else:
            k=(center_name,r['SKU번호'])
            if k not in extra:extra[k]={'센터명':center_name,'ME코드':', '.join(r['ME후보']) or f"SKU {r['SKU번호']}",
                 '품목명':clean(r['SKU명']),'발주수량':0,'발주금액':0,'입고수량':0.,'입고금액':0.,'입고건수':0}
            extra[k]['입고수량']+=r['입고수량'];extra[k]['입고금액']+=r['입고금액'];extra[k]['입고건수']+=1

    records=[]
    for item in items:
        records.append({k:v for k,v in item.items() if k not in ('그룹','SKU집합')})
    for extra_item in extra.values():
        exceptions_rows.append({'입고/반출시각':pd.NaT,'센터코드':extra_item['센터명'],
            'SKU번호':extra_item['ME코드'],'SKU명':extra_item['품목명'],
            '입고수량':extra_item['입고수량'],'입고금액':extra_item['입고금액'],
            '매핑상태':'발주 품목 연결 실패', '센터후보':[extra_item['센터명']],
            'ME후보':[extra_item['ME코드']]})
    exceptions=pd.DataFrame(exceptions_rows)
    records+=list(extra.values())
    merged=pd.DataFrame(records,columns=['센터명','ME코드','품목명','발주수량','입고수량','발주금액','입고금액','입고건수'])
    merged['수량차이']=merged['입고수량']-merged['발주수량']
    merged['금액차이']=merged['입고금액']-merged['발주금액']
    def status(r):
        if r['발주수량']==0:return '발주 외 입고'
        if r['입고수량']==0:return '미입고'
        if r['입고수량']<r['발주수량']:return '부분 입고'
        if r['입고수량']>r['발주수량']:return '초과 입고'
        if abs(r['금액차이'])>0.5:return '금액 확인 필요'
        return '매입 완료'
    merged['매입상태']=merged.apply(status,axis=1)
    cols=['센터명','ME코드','품목명','발주수량','입고수량','수량차이','발주금액','입고금액','금액차이','입고건수','매입상태']
    merged=merged[cols].sort_values(['센터명','품목명']).reset_index(drop=True)

    merged.insert(0,'입고예정일',pd.Timestamp(order_date))
    if not exceptions.empty:
        exceptions.insert(0,'입고예정일',pd.Timestamp(order_date))
    return merged,exceptions

results=[]
issues=[]
raw_total=0.
order_total=0.
for order_date,period_end in periods:
    selected_sales=sales[sales['예정일']==order_date].copy()
    selected_raw=raw[raw['입고일'].between(order_date,period_end,inclusive='both')].copy()
    result,issue=reconcile_period(selected_sales,selected_raw,order_date,period_end)
    results.append(result)
    if not issue.empty: issues.append(issue)
    raw_total+=selected_raw['입고금액'].sum()
    order_total+=selected_sales['발주금액'].sum()
merged=pd.concat(results,ignore_index=True)
exceptions=pd.concat(issues,ignore_index=True) if issues else pd.DataFrame()
merged=merged.sort_values(['입고예정일','센터명','품목명']).reset_index(drop=True)

if not exceptions.empty:
    st.warning(f'센터 또는 상품 매칭 확인이 필요한 RAW {len(exceptions):,}행이 있습니다. 해당 행은 품목별 집계에서 제외되므로 결과를 확인해 주세요.')
col=st.columns(4)
col[0].metric('발주 총금액',f'{order_total:,.0f}원')
col[1].metric('실제 입고 총금액',f'{raw_total:,.0f}원')
col[2].metric('총금액 차이',f'{raw_total-order_total:+,.0f}원')
col[3].metric('매입 완료 / 비교 품목',f"{(merged['매입상태']=='매입 완료').sum():,} / {len(merged):,}")
st.caption('같은 센터에서 동일 SKU에 연결된 ME코드 또는 동일한 발주 품목명은 하나의 제품으로 비교합니다. RAW는 중복 제거 없이 전 행 합산합니다.')

t1,t2,t3=st.tabs(['① 전체 매입 현황','② 센터별 품목 상세','③ 매칭 확인 필요'])
with t1:
    st.subheader('확인 필요 품목')
    pending=merged[merged['매입상태']!='매입 완료']
    if pending.empty:st.success('품목별 매입 확인 결과 모두 완료되었습니다.')
    else:st.dataframe(pending,hide_index=True,use_container_width=True)
with t2:
    st.subheader('센터별 품목 매입 대조')
    options=['전체 센터']+sorted(merged['센터명'].dropna().unique().tolist())
    selected_center=st.selectbox('센터 선택',options)
    center_data=merged if selected_center=='전체 센터' else merged[merged['센터명']==selected_center]
    # 센터별 금액 지표는 상태 필터와 무관하게 해당 센터 전체 기준으로 표시합니다.
    a,b,c=st.columns(3)
    a.metric('선택 센터 발주금액',f"{center_data['발주금액'].sum():,.0f}원")
    b.metric('선택 센터 입고금액',f"{center_data['입고금액'].sum():,.0f}원")
    c.metric('선택 센터 금액 차이',f"{center_data['금액차이'].sum():+,.0f}원")

    date_options=sorted(center_data['입고예정일'].dt.date.unique().tolist())
    selected_dates=st.multiselect('입고예정일 필터 (복수 선택 가능)',options=date_options,default=date_options,format_func=lambda x:x.strftime('%Y-%m-%d'))
    # 엑셀 필터처럼 여러 매입상태를 선택할 수 있습니다. 기본값은 전체 상태입니다.
    status_order=['매입 완료','미입고','부분 입고','초과 입고','금액 확인 필요','발주 외 입고']
    available_statuses=[x for x in status_order if x in center_data['매입상태'].unique()]
    selected_statuses=st.multiselect(
        '매입상태 필터 (복수 선택 가능)',
        options=available_statuses,
        default=available_statuses,
        key='center_purchase_status_filter',
        placeholder='표시할 매입상태를 선택하세요',
    )
    shown=center_data[center_data['매입상태'].isin(selected_statuses) & center_data['입고예정일'].dt.date.isin(selected_dates)].copy()
    st.caption(f'표시 품목: {len(shown):,}건 / 선택 센터 전체: {len(center_data):,}건 · 금액 지표는 필터 적용 전 센터 전체 기준입니다.')
    if shown.empty:
        st.info('선택한 조건에 해당하는 품목이 없습니다.')
    else:
        st.dataframe(shown,hide_index=True,use_container_width=True)
with t3:
    st.subheader('센터/SKU 매칭 검증')
    if exceptions.empty:st.success('선택 기간의 매칭 오류가 없습니다.')
    else:
        view=exceptions[['입고예정일','입고/반출시각','센터코드','SKU번호','SKU명','입고수량','입고금액','매핑상태']].copy()
        view['센터 후보']=exceptions['센터후보'].map(lambda a:', '.join(a))
        view['ME 후보']=exceptions['ME후보'].map(lambda a:', '.join(a))
        st.dataframe(view,hide_index=True,use_container_width=True)
    st.caption('입고일이 일치해도 센터코드·SKU/ME코드 매핑이 연결되지 않으면 미입고로 표시될 수 있습니다. 이 탭에서 연결 실패 내역을 확인하세요.')

center=merged.groupby('센터명',as_index=False).agg(발주수량=('발주수량','sum'),입고수량=('입고수량','sum'),발주금액=('발주금액','sum'),입고금액=('입고금액','sum'))
center['수량차이']=center['입고수량']-center['발주수량']
center['금액차이']=center['입고금액']-center['발주금액']
buf=io.BytesIO()
with pd.ExcelWriter(buf,engine='openpyxl') as writer:
    merged.to_excel(writer,sheet_name='품목별 매입 확인',index=False)
    center.to_excel(writer,sheet_name='센터별 현황',index=False)
    if not exceptions.empty:
        exceptions[['입고예정일','입고/반출시각','센터코드','SKU번호','SKU명','입고수량','입고금액','매핑상태']].to_excel(writer,sheet_name='매칭 확인 필요',index=False)
st.download_button('매입 확인 결과 엑셀 다운로드',data=buf.getvalue(),file_name=f'쿠팡_매입확인_{target_dates[0]:%Y%m%d}_{target_dates[-1]:%Y%m%d}.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
