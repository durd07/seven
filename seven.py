# -*- coding: utf-8 -*-
import streamlit as st
import numpy as np
import pandas as pd

st.set_page_config(layout='wide')

pd.set_option("display.max_colwidth", 1000, 'display.width', 1000)

items = [
    '嗜碱性粒细胞计数(BASO#)(10^9/L)',
    '血小板平均体积(MPV)(fL)',
    '中性粒细胞计数(NEUT#)(10^9/L)',
    '中性粒细胞百分比(NEUT%)(%)',
    '血小板压积(PCT)(%)',
    '血小板分布宽度(PDW)(%)',
    '大血小板比率(P-LCR)',
    '血小板总数(PLT)(10^9/L)',
    '红细胞计数(RBC)(10^12/L)',
    '红细胞分布宽度CV(RDW-CV)(%)',
    '红细胞分布宽度-SD(RDW-SD)(fL)',
    '单核细胞百分比(MONO%)(%)',
    '单核细胞计数(MONO#)(10^9/L)',
    '平均红细胞体积(MCV)(fL)',
    '嗜碱性粒细胞百分比(BASO%)(%)',
    'C-反应蛋白(CRP)(mg/L)',
    '嗜酸性粒细胞计数(EO#)(10^9/L)',
    '嗜酸性粒细胞百分比(EO%)(%)',
    '红细胞压积(HCT)(%)',
    '血红蛋白(HGB)(g/L)',
    '淋巴细胞计数(LYMPH#)(10^9/L)',
    '淋巴细胞百分比(LYMPH%)(%)',
    '平均血红蛋白含量(MCH)(pg)',
    '平均血红蛋白浓度(MCHC)(g/L)',
    '白细胞数目(WBC)(10^9/L)'
    ]
items_ref = [x + '_ref' for x in items]
df = pd.read_excel('杜子期血常规.xlsx', engine='openpyxl')

df_new = pd.DataFrame([], index=[rv for r in zip(items, items_ref) for rv in r])

for index, row in df.iteritems():
    df_new[index] = ''
    for i, item in enumerate(row):
        if item in items:
               df_new[index][item] = row[i + 1]
               df_new[index][item + '_ref'] = row[i + 2]

st.title('杜子期血常规数据统计')
st.write(df_new)


chart_items = set()

if st.sidebar.checkbox('所有项'):
    chart_items = set(items)

for item in items:
    if st.sidebar.checkbox(item):
        chart_items.add(item)

if chart_items:
    st.line_chart(df_new.loc[chart_items, :].T)
else:
    st.line_chart(df_new.loc['血小板总数(PLT)(10^9/L)'].T)

#for index, row in df_new.iterrows():
#    if not index.endswith('ref'):
#        st.line_chart(row)

#for index, row in df_new.iteritems():
#    try:
#        row.plot(legend=True, figsize=(20, 5))
#        #df.plot_bokeh.line(x=)
#    except:
#        pass
#
#df_new.style.applymap(lambda v : 'background-color: %s' %'#FFCCFF' if v else'background-color: %s'% '#FFCCEE')
#with pd.ExcelWriter('df_style.xlsx', engine='openpyxl') as writer:
#    df_new.to_excel(writer, index=True, sheet_name='sheet')
