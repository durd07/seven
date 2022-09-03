# -*- coding: utf-8 -*-
import streamlit as st
import numpy as np
import pandas as pd
import altair as alt
from io import BytesIO

def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=True, sheet_name='杜子期血常规数据统计')
    workbook = writer.book
    worksheet = writer.sheets['杜子期血常规数据统计']
    format1 = workbook.add_format({'num_format': '0.00'})
    worksheet.set_column('A:Z', None, format1)
    writer.save()
    processed_data = output.getvalue()
    return processed_data

st.set_page_config(layout='wide')

pd.set_option("display.max_colwidth", 1000, 'display.width', 1000)

def highlight_dataframe(s):
    lst = []
    for i in range(0, len(s) - 1, 2):
        try:
            tmp = float(s[i])
            min, max = s[i+1].split('～')
            if tmp < float(min):
                lst.append('color: orange')
            elif tmp > float(max):
                lst.append('color: red')
            else:
                lst.append('')
        except Exception as e:
            lst.append('')
            #print(s[i], s[i+1], e)
        lst.append('')
    return lst

items_map = {
'白细胞计数(WBC)(10^9/L)':                  '白细胞计数(WBC)(10^9/L)',
'淋巴细胞绝对值(LYM#)(10^9/L)':             '淋巴细胞绝对值(LYM#)(10^9/L)',
'中性粒细胞绝对值(NEU#)(10^9/L)':           '中性粒细胞绝对值(NEU#)(10^9/L)',
'单核细胞绝对值(MON#)(10^9/L)':             '单核细胞绝对值(MON#)(10^9/L)',
'嗜酸性粒细胞绝对值(EOS#)(EOS#)(10^9/L)':   '嗜酸性粒细胞绝对值(EOS#)(EOS#)(10^9/L)',
'嗜碱性粒细胞绝对值(BAS#)(BAS#)(10^9/L)':   '嗜碱性粒细胞绝对值(BAS#)(BAS#)(10^9/L)',
'红细胞体积分布宽度-CV(RDW-CV)(%)':         '红细胞体积分布宽度-CV(RDW-CV)(%)',
'红细胞体积分布宽度-SD(RDW-SD)(fL)':        '红细胞体积分布宽度-SD(RDW-SD)(fL)',
'血小板体积分布宽度(PDW)(%)':               '血小板体积分布宽度(PDW)(%)',
'血小板平均体积(MPV)(fL)':                  '血小板平均体积(MPV)(fL)',
'血小板压积(PCT)(%)':                       '血小板压积(PCT)(%)',
'嗜碱性粒细胞百分比(BAS%)(BAS%)(%)':        '嗜碱性粒细胞百分比(BAS%)(BAS%)(%)',
'嗜酸性粒细胞百分比(EO%)(EOS%)(%)':         '嗜酸性粒细胞百分比(EO%)(EOS%)(%)',
'红细胞计数(RBC)(10^12/L)':                 '红细胞计数(RBC)(10^12/L)',
'血红蛋白浓度(HGB)(g/L)':                   '血红蛋白浓度(HGB)(g/L)',
'红细胞压积(HCT)(%)':                       '红细胞压积(HCT)(%)',
'平均红细胞体积(MCV)(fL)':                  '平均红细胞体积(MCV)(fL)',
'平均红细胞血红蛋白含量(MCH)(MCH)(pg)':     '平均红细胞血红蛋白含量(MCH)(MCH)(pg)',
'平均红细胞血红蛋白浓度(MCHC)(MCHC)(g/L)':  '平均红细胞血红蛋白浓度(MCHC)(MCHC)(g/L)',
'血小板计数(PLT)(10^9/L)':                  '血小板计数(PLT)(10^9/L)',
'淋巴细胞百分比(LYM%)(%)':                  '淋巴细胞百分比(LYM%)(%)',
'中性粒细胞百分比(NEU%)(%)':                '中性粒细胞百分比(NEU%)(%)',
'单核细胞百分比(MON%)(%)':                  '单核细胞百分比(MON%)(%)',
'大血小板比率(P-LC,R)':                     '大血小板比率(P-LC,R)',

'嗜碱性粒细胞计数(BASO#)(10^9/L)':          '嗜碱性粒细胞绝对值(BAS#)(BAS#)(10^9/L)',
'血小板平均体积(MPV)(fL)':                  '血小板平均体积(MPV)(fL)',
'中性粒细胞计数(NEUT#)(10^9/L)':            '中性粒细胞绝对值(NEU#)(10^9/L)',
'中性粒细胞百分比(NEUT%)(%)':               '中性粒细胞百分比(NEU%)(%)',
'血小板压积(PCT)(%)':                       '血小板压积(PCT)(%)',
'血小板分布宽度(PDW)(%)':                   '血小板体积分布宽度(PDW)(%)',
'大血小板比率(P-LCR)':                      '大血小板比率(P-LC,R)',
'血小板总数(PLT)(10^9/L)':                  '血小板计数(PLT)(10^9/L)',
'红细胞计数(RBC)(10^12/L)':                 '红细胞计数(RBC)(10^12/L)',
'红细胞分布宽度CV(RDW-CV)(%)':              '红细胞体积分布宽度-CV(RDW-CV)(%)',
'红细胞分布宽度-SD(RDW-SD)(fL)':            '红细胞体积分布宽度-SD(RDW-SD)(fL)',
'单核细胞百分比(MONO%)(%)':                 '单核细胞百分比(MON%)(%)',
'单核细胞计数(MONO#)(10^9/L)':              '单核细胞绝对值(MON#)(10^9/L)',
'平均红细胞体积(MCV)(fL)':                  '平均红细胞体积(MCV)(fL)',
'嗜碱性粒细胞百分比(BASO%)(%)':             '嗜碱性粒细胞百分比(BAS%)(BAS%)(%)',
#'C-反应蛋白(CRP)(mg/L)',
'嗜酸性粒细胞计数(EO#)(10^9/L)':            '嗜酸性粒细胞绝对值(EOS#)(EOS#)(10^9/L)',
'嗜酸性粒细胞百分比(EO%)(%)':               '嗜酸性粒细胞百分比(EO%)(EOS%)(%)',
'红细胞压积(HCT)(%)':                       '红细胞压积(HCT)(%)',
'血红蛋白(HGB)(g/L)':                       '血红蛋白浓度(HGB)(g/L)',
'淋巴细胞计数(LYMPH#)(10^9/L)':             '淋巴细胞绝对值(LYM#)(10^9/L)',
'淋巴细胞百分比(LYMPH%)(%)':                '淋巴细胞百分比(LYM%)(%)',
'平均血红蛋白含量(MCH)(pg)':                '平均红细胞血红蛋白含量(MCH)(MCH)(pg)',
'平均血红蛋白浓度(MCHC)(g/L)':              '平均红细胞血红蛋白浓度(MCHC)(MCHC)(g/L)',
'白细胞数目(WBC)(10^9/L)':                  '白细胞计数(WBC)(10^9/L)'
}

all_draw_items = [
        '血小板计数(PLT)(10^9/L)',
        '血红蛋白浓度(HGB)(g/L)',
        '白细胞计数(WBC)(10^9/L)',
        '红细胞计数(RBC)(10^12/L)',
        '中性粒细胞绝对值(NEU#)(10^9/L)',
        '单核细胞绝对值(MON#)(10^9/L)',
        '嗜酸性粒细胞绝对值(EOS#)(EOS#)(10^9/L)',
        '嗜碱性粒细胞绝对值(BAS#)(BAS#)(10^9/L)',

        '淋巴细胞绝对值(LYM#)(10^9/L)',
        '嗜碱性粒细胞百分比(BAS%)(BAS%)(%)',
        '嗜酸性粒细胞百分比(EO%)(EOS%)(%)',

        '红细胞体积分布宽度-CV(RDW-CV)(%)',
        '红细胞体积分布宽度-SD(RDW-SD)(fL)',

        '血小板体积分布宽度(PDW)(%)',
        '血小板平均体积(MPV)(fL)',
        '血小板压积(PCT)(%)',
        '大血小板比率(P-LC,R)',

        '红细胞压积(HCT)(%)',
        '平均红细胞体积(MCV)(fL)',
        '平均红细胞血红蛋白含量(MCH)(MCH)(pg)',
        '平均红细胞血红蛋白浓度(MCHC)(MCHC)(g/L)',
        '淋巴细胞百分比(LYM%)(%)',
        '中性粒细胞百分比(NEU%)(%)',
        '单核细胞百分比(MON%)(%)',
        ]

items = set(items_map.values())
items_ref = [x + '_参考范围' for x in items]
df = pd.read_excel('杜子期血常规.xlsx', engine='openpyxl')

df_new = pd.DataFrame([], index=[rv for r in zip(items, items_ref) for rv in r])

for index, row in df.iteritems():
    df_new[index] = ''
    for i, item in enumerate(row):
        if item in items_map:
            try:
                df_new[index][items_map[item]] = float(row[i + 1])
            except:
                df_new[index][items_map[item]] = np.nan
            df_new[index][items_map[item] + '_参考范围'] = row[i + 2]

df_new.columns = np.array([x.date() for x in df_new.columns])

st.title('杜子期血常规数据统计')
df_new_str = df_new.astype(str)
st.write(df_new_str.style.apply(highlight_dataframe, axis=0))
st.download_button("Export to Excel", data=to_excel(df_new), file_name='杜子期血常规数据统计.xlsx')

chart_items = list()

#other = st.sidebar.expander('其他选项')
#if other.checkbox('显示原始数据'):
#    st.write(df)

#st.sidebar.write('')
st.sidebar.write('请选择画图项')
if st.sidebar.checkbox('所有项'):
    chart_items = all_draw_items

for item in all_draw_items:
    if st.sidebar.checkbox(item):
        chart_items.append(item)

if chart_items:
    df = df_new.loc[chart_items, :].T
    #df.index = df.index.to_numpy(dtype='datetime64')
    #st.line_chart(df)
else:
    df = df_new.loc[['血小板计数(PLT)(10^9/L)'], :].T
    #st.line_chart(df_new.loc['血小板计数(PLT)(10^9/L)'].T)

st.write("### 图表展示")

df['date'] = df.index

options = np.array(df['date']).tolist()

(start_time, end_time) = st.select_slider("请选择时间序列长度：",
     #min_value = datetime(2013, 10, 1,),
     #max_value = datetime(2018, 10, 31,),
     options = options,
     value= (options[0],options[-1],),
 )

df = df[(df['date']>=start_time) & (df['date']<=end_time)]

st.write("#### 单表展示")
for column in df.columns:
    if column == 'date':
        continue

    st.write(column)
    st.vega_lite_chart(data=df, spec={
        'layer': [
            {
                'mark': {
                    'type': 'line',
                    'point': {"filled": False, "fill": "white"},
                    'tooltip': True
                }
            },
            {
                'mark': {
                    'type': 'text',
                    'align': 'center',
                    'baseline': 'line-bottom',
                    'dx': 3,
                    'size': 14
                },
                'encoding': {
                    'text': {'field': column, 'type': 'quantitative'}
                }
            }
        ],
        'encoding': {
            'x': {
                "type": "temporal",
                #'timeUnit': 'date',
                'field': 'date',
                },
            'y': {
                "type": "quantitative",
                'field': column,
                #'aggregate': 'mean'
                },
            #'color': {'field': 'field', 'type': 'nominal'},
            },
        }, use_container_width=True)

st.write("#### 合并展示")
df = df.drop('date', axis=1)
ndf = df.melt(var_name='field', value_name='data')
xx = pd.concat([df.index.to_series()] * int((ndf.shape[0] / len(df.index))))
ndf['date'] = xx.values

st.vega_lite_chart(data=ndf, spec={
    'layer': [
        {
            'mark': {
                'type': 'line',
                'point': {"filled": False, "fill": "white"},
                'tooltip': True
            }
        },
        {
            'mark': {
                'type': 'text',
                'align': 'center',
                'baseline': 'line-bottom',
                'dx': 3,
                'size': 14
            },
            'encoding': {
                'text': {'field': 'data', 'type': 'quantitative'}
            }
        }
    ],
    'encoding': {
        'x': {
            "type": "temporal",
            #'timeUnit': 'date',
            'field': 'date',
            },
        'y': {
            "type": "quantitative",
            #'field': '血小板计数(PLT)(10^9/L)'
            'field': 'data',
            #'aggregate': 'mean'
            },
        'color': {'field': 'field', 'type': 'nominal'},
        },
    }, use_container_width=True)

st.write('相关系数矩阵')
df = df_new.filter(regex='^((?!_参考范围$).)*$', axis=0).astype(float)
st.write(df.T.corr())


cor_data = df.T.corr().stack().reset_index().rename(columns={0: 'correlation', 'level_0': 'variable', 'level_1': 'variable2'})
cor_data['correlation_label'] = cor_data['correlation'].map('{:.2f}'.format)

base = alt.Chart(cor_data).encode(
    x='variable2:O',
    y='variable:O'
)

# Text layer with correlation labels
# Colors are for easier readability
text = base.mark_text().encode(
    text='correlation_label',
    color=alt.condition(
        alt.datum.correlation > 0.5,
        alt.value('white'),
        alt.value('black')
    )
)

# The correlation heatmap itself
cor_plot = base.mark_rect().encode(
    color='correlation:Q'
)

st.altair_chart(cor_plot + text, use_container_width=True)
