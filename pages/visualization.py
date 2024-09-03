import pandas as pd 
import streamlit as st 
import os
import numpy as np
import random
import time
from pathlib import Path
from streamlit_gsheets import GSheetsConnection
import plotly.express as px
import plotly.graph_objects as go
conn = st.connection('gsheets', type=GSheetsConnection)
st.success('**DEVELOPED BY Dr. LUMINSA DESIRE**')

df = conn.read(worksheet='MPIGI', usecols=list(range(25)), ttl=5)
#df = df.dropna(how='all')
#st.write(df.columns)

weeeks = df['WEEK'].unique()
fac = df['FACILITY'].unique()
dfy = []
for every in fac:
    dff = df[df['FACILITY']== every]
    dff = dff.drop_duplicates(subset=['FACILITY'], keep = 'last')
    dfy.append(dff)
dff = pd.concat(dfy)

dfs=[]   
for each in weeeks:
    dfa = df[df['WEEK']==each]
    dfa = dfa.drop_duplicates(subset=['FACILITY'], keep = 'last')
    dfs.append(dfa)
df = pd.concat(dfs)
df['WEEK'] = df['WEEK'].astype(int)
df['POTENTIAL'] = df['POTENTIAL'].astype(int)
#st.write(df['TXML'])
df['TXML'] = df['TXML'].astype(int)
df['TO'] = df['TO'].astype(int)
df['TI'] = df['TI'].astype(int)
df['Q3 CURR'] = df['Q3 CURR'].astype(int)
df['Q4 CURR'] = df['Q4 CURR'].astype(int)
df['EXPECTED'] = df['EXPECTED'].astype(int)
df['NO VL'] = df['NO VL'].astype(int)
df['HAS VL'] = df['HAS VL'].astype(int)
#file = r"C:\Users\Desire Lumisa\Downloads\TXML (5).xlsx"
#df = pd.read_excel(file)
st.sidebar.subheader('Filter from here ')
week = st.sidebar.multiselect('Pick a week', df['WEEK'].unique())
file2 = r'ALL.xlsx'
dfx = pd.read_excel(file2)
#create for the state
if not week:
    df2 = df.copy()
else:
    df2 = df[df['WEEK'].isin(week)]

#create for district
district = st.sidebar.multiselect('Choose a district', df2['DISTRICT'].unique())
if not district:
    df3 = df2.copy()
else:
    df3 = df2[df2['DISTRICT'].isin(district)]
 
#for facility
facility = st.sidebar.multiselect('Choose a facility', df3['FACILITY'].unique())

#Filter Week, District, Facility

if not week and not district and not facility:
    filtered_df = df
    filtered_dff = dff
elif not district and not facility:
    filtered_df = df[df['WEEK'].isin(week)]
    filtered_dff = dff
elif not week and not facility:
    filtered_df = df[df['DISTRICT'].isin(district)]
    filtered_dff = dff[df['DISTRICT'].isin(district)]
elif district and facility:
    filtered_df = df3[df['DISTRICT'].isin(district)& df3['FACILITY'].isin(facility)]
    filtered_dff = dff[dff['DISTRICT'].isin(district)& dff['FACILITY'].isin(facility)]
elif week and facility:
    filtered_df = df3[df['WEEK'].isin(week)& df3['FACILITY'].isin(facility)]
    filtered_dff = dff[dff['FACILITY'].isin(facility)]
elif week and district:
    filtered_df = df3[df['WEEK'].isin(week)& df3['DISTRICT'].isin(district)]
    filtered_dff = dff[dff['DISTRICT'].isin(district)]
elif facility:
    filtered_df = df3[df3['FACILITY'].isin(facility)]
    filtered_dff = dff[dff['FACILITY'].isin(facility)]
else:
    filtered_df = df3[df3['WEEK'].isin(week) & df3['DISTRICT'].isin(district)&df3['FACILITY'].isin(facility)]
    filtered_dff = dff[dff['DISTRICT'].isin(district)& dff['FACILITY'].isin(facility)]
#################################################################################################
st.divider()
current_time = time.localtime()
k = time.strftime("%V", current_time)
k = int(k)
k = k+13
dfa = dfx[['DISTRICT', 'FACILITY']].copy()
dfb = df[df['WEEK'] == k].copy()
dfb = dfb[['DISTRICT' , 'FACILITY']]
merged = dfa.merge(dfb, on=['DISTRICT', 'FACILITY'], how='left', indicator=True)
none = merged[merged['_merge'] == 'left_only'].drop(columns=['_merge'])
none = none.reset_index()
none = none.drop(columns='index')
all = none.shape[0]
buk = none[none['DISTRICT']=='BUTAMBALA'].copy()
semb = none[none['DISTRICT']=='GOMBA'].copy()
dist = none[none['DISTRICT']=='MPIGI'].copy()
kal = none[none['DISTRICT']=='KALUNGU'].copy()
city = none[none['DISTRICT']=='MKA CITY'].copy()

bu = buk.shape[0]
se = semb.shape[0]
di = dist.shape[0]

 
if bu ==0:
   b = 'all facilities have reported'
else:
   b = f'{bu} have not reported'
    
if se ==0:
   s = 'all facilities reported'
else:
   s = f'{se} have not reported'
    
if di ==0:
   d = 'all facilities reported'
else:
   d = f'{di} have not reported'
    

if all ==0:
    st.write('** ALL FACILITIES IN THE CLUSTER REPORTED**')
else:
    st.markdown(f'**{all} FACILITIES HAVE NOT REPORTED IN THIS WEEK**')
    st.markdown(f'**BUTAMBALA {b}, GOMBA {s}, MPIGI {k}**')
    with st.expander('Click to see them'):
        st.dataframe(none)
   
st.divider()
#############################################################################################
#filtered_df = filtered_df[filtered_df['WEEK']==k].copy()
pot = filtered_dff['POTENTIAL'].sum()
Q3 = filtered_dff['Q3 CURR'].sum()
ti = filtered_dff['TI'].sum()
new = filtered_dff['TX NEW'].sum()
uk = filtered_dff['UNKNOWN GAIN'].sum()
los = filtered_dff['TXML'].sum()
to  = filtered_dff['TO'].sum()
dd = filtered_dff['DEAD'].sum()
Q4 = filtered_dff['Q4 CURR'].sum()

labels = ["Q3 Curr", "TI", "TX NEW", 'Unkown',"Potential", "TXML","DEAD", "TO", "Q4 Curr"]
values = [Q3, ti, new, uk, pot, -los,-dd, -to, Q4]
measure = ["absolute", "relative", "relative","relative", "total", "relative", "relative", "relative","total"]
# Create the waterfall chart
fig = go.Figure(go.Waterfall(
    name="Waterfall",
    orientation="v",
    measure=measure,
    x=labels,
    textposition="outside",
    text=[f"{v}" for v in values],
    y=values
))

# Add titles and labels and adjust layout properties
fig.update_layout(
    title="Waterfall Analysis",
    xaxis_title="Categories",
    yaxis_title="Values",
    showlegend=True,
    height=425,  # Adjust height to ensure the chart fits well
    margin=dict(l=20, r=20, t=60, b=20),  # Adjust margins to prevent clipping
    yaxis=dict(automargin=True)
)
# Show the plot
#fig.show()
#st.title("Waterfall Chart in Streamlit")
st.plotly_chart(fig)
st.divider()
##########################################################################
#######ONE YEAR COHORT
#filtered_df = filtered_df[filtered_df['WEEK']==k].copy()
original = filtered_dff['ORIGINAL COHORT'].sum()
newti = filtered_dff['ONE YEAR TI'].sum()
newlydx = original - newti

newlos = filtered_dff['ONE YEAR LOST'].sum()
newto  = filtered_dff['ONE YEAR TO'].sum()
newdd = filtered_dff['ONE YEAR DEAD'].sum()
active = filtered_dff['ONE YEAR ACTIVE'].sum()

labels = ["NEWLY DX", "TIs", "ORIGINAL", 'LTFU',"TOs","DEAD", "ACTIVE"]
values = [newlydx, newti, original, -newlos, -newto,-newdd, active]
measure = ["absolute", "relative", "total", "relative", "relative", "relative","total"]
# Create the waterfall chart
figy = go.Figure(go.Waterfall(
    name="Waterfall",
    orientation="v",
    measure=measure,
    x=labels,
    textposition="outside",
    text=[f"{v}" for v in values],
    y=values
))

# Add titles and labels and adjust layout properties
figy.update_layout(
    title="ONE YEAR COHORT ANALYSIS",
    xaxis_title="Categories",
    yaxis_title="Values",
    showlegend=True,
    height=425,  # Adjust height to ensure the chart fits well
    margin=dict(l=20, r=20, t=60, b=20),  # Adjust margins to prevent clipping
    yaxis=dict(automargin=True)
)
# Show the plot
#fig.show()
#st.title("Waterfall Chart in Streamlit")
st.plotly_chart(figy)

########################################################################################
#ONE YEAR PIE CHART
col1, col2 = st.columns(2)
#dfv = filtered_dff.copy()
#dfv = dfv[['FACILITY','ORIGINAL COHORT', 'ONE YEAR TO', 'ONE YEAR LOST', 'ONE YEAR TO', 'ONE YEAR DEAD', 'ONE YEAR ACTIVE']].copy()
#dfv = dfv.rename(columns ={'ORIGINAL COHORT': 'ORIGINAL', 'ONE YEAR TO': 'TO', 'ONE YEAR LOST': 'LTFU', 'ONE YEAR TO': 'TO', 'ONE YEAR DEAD': 'DEAD', 'ONE YEAR ACTIVE':'ACTIVE'})

pied = filtered_dff.copy()#[filtered_df['WEEK']==k]
pied['LOST NEW'] = pied['ORIGINAL COHORT']- pied['ONE YEAR ACTIVE'] 
pied = pied[['LOST NEW', 'ONE YEAR ACTIVE']]
melted = pied.melt(var_name='Category', value_name='values')
fig = px.pie(melted, values= 'values', title='ONE YEAR RETENTION RATE', names='Category', hole=0.5,color='Category',  
             color_discrete_map={'LOST NEW': 'red', 'ONE YEAR ACTIVE': 'blue'} )
    #fig.update_traces(text = 'RETENTION', text_position='Outside')
with col1:
    if pied.shape[0]==0:
        pass
    else:
        st.plotly_chart(fig, use_container_width=True)
with col2:
    st.write('')
    st.write('')
    st.write('')
    st.write('')
    st.write('')
    st.write('')
    st.write('')
    st.write('')
    st.write('')
    st.write('')
    st.write('')
    st.write('')
    dfv = filtered_dff.copy()
    dfv = dfv[dfv['WEEK']>46].copy()
    dfv = dfv[['FACILITY','ORIGINAL COHORT', 'ONE YEAR TO', 'ONE YEAR LOST', 'ONE YEAR DEAD', 'ONE YEAR ACTIVE']].copy()
    dfv = dfv.rename(columns ={'ORIGINAL COHORT': 'ORGNAL', 'ONE YEAR LOST': 'LTFU', 'ONE YEAR TO': 'TO', 'ONE YEAR DEAD': 'DEAD', 'ONE YEAR ACTIVE':'ACTIVE'})
    dfv[['ORGNAL', 'LTFU', 'TO', 'DEAD', 'ACTIVE']]= dfv[['ORGNAL', 'LTFU', 'TO', 'DEAD', 'ACTIVE']].astype(int)
    dfv.reset_index(drop=True, inplace=True)
    dfv.set_index('FACILITY', inplace= True)
    with st.expander('CLICK HERE TO VIEW ONE YEAR COHORT DATA'):
          st.table(dfv)
        

########################################################################################

#LINE GRAPHS
st.divider()
st.subheader('TXML PERFORMANCE')
grouped = filtered_df.groupby('WEEK').sum(numeric_only=True).reset_index()

melted = grouped.melt(id_vars=['WEEK'], value_vars=['Q3 CURR', 'Q4 CURR', 'POTENTIAL'],
                            var_name='OUTCOME', value_name='Total')

melted2 = grouped.melt(id_vars=['WEEK'], value_vars=['TXML', 'TO'],
                            var_name='OUTCOME', value_name='Total')
melted['WEEK'] = melted['WEEK'].astype(str)
melted2['WEEK'] = melted2['WEEK'].astype(str)
fig2 = px.line(melted, x='WEEK', y='Total', color='OUTCOME', markers=True,
              title='Trends in TXML and TO', labels={'WEEK':'WEEK', 'Total': 'No. of clients', 'OUTCOME': 'Outcomes'})

fig3 = px.line(melted2, x='WEEK', y='Total', color='OUTCOME', markers=True, color_discrete_sequence=['blue','red'],
              title='Trends in TXML and TO', labels={'WEEK':'WEEK', 'Total': 'No. of clients', 'OUTCOME': 'Outcomes'})

fig2.update_layout(
    width=800,  # Set the width of the plot
    height=400,  # Set the height of the plot
    xaxis=dict(showline=True, linewidth=1, linecolor='black'),  # Show x-axis line
    yaxis=dict(showline=True, linewidth=1, linecolor='black')   # Show y-axis line
)
fig2.update_xaxes(type='category')
fig3.update_layout(
    width=800,  # Set the width of the plot
    height = 400,  # Set the height of the plot
    xaxis=dict(showline=True, linewidth=1, linecolor='black'),  # Show x-axis line
    yaxis=dict(showline=True, linewidth=1, linecolor='black')   # Show y-axis line
)
fig3.update_xaxes(type='category')
colx,coly = st.columns([2,1])
with colx:
    st.plotly_chart(fig2, use_container_width= True)

with coly:
    st.plotly_chart(fig3, use_container_width= True)
    #st.plotly_chart(fig3, use_container_width=True)
############################################################################################
group = grouped[grouped['WEEK']>46]
melted = group.melt(id_vars=['WEEK'], value_vars=['Q3 CURR', 'POTENTIAL', 'Q4 CURR'],
                            var_name='OUTCOME', value_name='Total')
fig5 = px.bar(
    melted,
    x='WEEK',
    y='Total',
    color='OUTCOME',
    title='Trends in TXML and TO',
    labels={'WEEK': 'WEEK', 'Total': 'No. of clients', 'OUTCOME': 'Outcomes'},
    barmode='group'  # Group bars by 'OUTCOME'
)

# Update the layout of the plot
fig5.update_layout(
    width=800,  # Set the width of the plot
    height=400,  # Set the height of the plot
    xaxis=dict(
        showline=True,  # Show x-axis line
        linewidth=1,    # Width of the x-axis line
        linecolor='black'  # Color of the x-axis line
    ),
    yaxis=dict(
        showline=True,  # Show y-axis line
        linewidth=1,    # Width of the y-axis line
        linecolor='black'  # Color of the y-axis line
    )
)

# To display the figure (assuming you are in a Jupyter notebook or a compatible environment)
st.plotly_chart(fig5, use_container_width= True)
#############################################################################################
# #HIGHEST TXML 
st.divider()
current_time = time.localtime()
k = time.strftime("%V", current_time)
k = int(k) +13
m = k-1
highest = filtered_dff[filtered_dff['TXML']>50]

highest = highest.sort_values(by=['TXML'])#, ascending=False)
highesta = highest.shape[0]
# highestb = highest[highest['WEEK']==m]

coly, colu = st.columns(2)
if highesta ==0:
    st.write('**FACILITY SELECTED IS NOT AMONGST THOSE WITH HIGH TXML**')
    pass
else:
    coly, colu = st.columns(2)
    with coly:
        figa = px.bar(
        highest,
        x='TXML',
        y='FACILITY',
        orientation='h',
        title='Facilities with highest TXML',
        labels={'TXML': 'TXML Value', 'FACILITY': 'Facility'}
         )
        st.plotly_chart(figa, use_container_width=True)
    with colu:
        st.markdown('##')
        with st.expander('FACILITIES WITH HIGHEST TXML'):
            highest = highest[['FACILITY', 'Q3 CURR' ,'Q4 CURR', 'TXML']].copy()
            highest['Q3 CURR'] = highest['Q3 CURR'].astype(int)
            highest['Q4 CURR'] = highest['Q4 CURR'].astype(int)
            highest['TXML'] = highest['TXML'].astype(int)
            highesty = highest.set_index('FACILITY', inplace = True)
            st.table(highest)
# if highestb.shape[0]==0:
#     with colu:
#          st.markdown('##')
#          st.markdown('##')
#          st.write("Selected facility or facilities do not have high TXML or didn't report last week")
# else:
#     figk = px.bar(
#     highestb,
#     x='TXML',
#     y='FACILITY',
#     orientation='h',
#     title='Facilities with highest TXML LAST WEEK',
#     labels={'TXML': 'TXML Value', 'FACILITY': 'Facility'}
#      )
#     with colu:
#         st.plotly_chart(figk, use_container_width=True)
#         # highest = highest[['FACILITY', 'TXML']]
#         # highest. set_index('FACILITY', inplace= True)
#         # highest['TXML'] = highest['TXML'].astype(int)

#############################################################################################
st.divider()
col1, col2,col3 = st.columns(3)
col2.write('**VL SECTION**')
meltvl = grouped.melt(id_vars='WEEK', value_vars=['EXPECTED', 'HAS VL'], var_name='PERFORMANCE', value_name='VL COVERAGE')

fig4 = px.line(meltvl, x='WEEK', y='VL COVERAGE', color='PERFORMANCE', markers=True, color_discrete_sequence=['blue','red'],
              title='Weekly VL Trend', labels={'WEEK':'WEEK', 'VL COVERAGE': 'No. BLED', 'PERFORMANCE': 'performance'})

fig4.update_layout(
    width=800,  # Set the width of the plot
    height=400,  # Set the height of the plot
    xaxis=dict(showline=True, linewidth=1, linecolor='black'),  # Show x-axis line
    yaxis=dict(showline=True, linewidth=1, linecolor='black')   # Show y-axis line
)

col1, col2 = st.columns(2)
with col1:
     st.plotly_chart(fig4, use_container_width=True)
     poorvl = filtered_dff[filtered_dff['VL COV (%)']<95]
     poorvl = poorvl[poorvl['WEEK']>46]
     poorvl = poorvl[poorvl['VL COV (%)']>0]
     poorvl= poorvl.sort_values(by = ['VL COV (%)'])
     poorvl = poorvl[['FACILITY','VL COV (%)']]
     poorvl.set_index('FACILITY', inplace=True)
     with st.expander('FACILITIES WITH POOR VL COV'):
         st.dataframe(poorvl)

pied = filtered_dff#[filtered_df['WEEK']==k]    
pied = pied[['HAS VL', 'NO VL']]
melted = pied.melt(var_name='Category', value_name='values')
fig = px.pie(melted, values= 'values', title='LASTEST VL COVERAGE', names='Category', hole=0.5)
    #fig.update_traces(text = 'VL COVERAGE', text_position='Outside')
if pied.shape[0]==0:
    with col2:
        st.markdown('##')
        st.markdown('##')
        st.write("Selected facility or facilities didn't report this week, so nothing to show")
else:
    with col2:
        st.plotly_chart(fig, use_container_width=True)

st.divider()
achieved = filtered_dff[filtered_dff['BALANCE'].isin(['EXCEEDED', 'EVEN'])].reset_index().copy()
achieved = achieved.drop_duplicates(subset =['FACILITY'], keep= 'last')
achieved['VL COV (%)'] = achieved['VL COV (%)'].astype(int)
achieved['TXML'] = achieved['TXML'].astype(int)
#achieved['BALANCE'] = achieved['BALANCE'].astype(int)
achieved['Q3 CURR'] = achieved['Q3 CURR'].astype(int)
achieved['Q4 CURR'] = achieved['Q4 CURR'].astype(int)
num = achieved['FACILITY'].nunique()
achieved = achieved[['DISTRICT','FACILITY', 'Q3 CURR', 'Q4 CURR', 'BALANCE','TXML', 'VL COV (%)']].copy() 
st.write('FACILITIES THAT HAVE ACHIEVED')
st.markdown(f"<h6>{num} facilities have achieved so far</h6>", unsafe_allow_html=True)
st.table(achieved)

st.divider()
notachieved = filtered_dff[~filtered_dff['BALANCE'].isin(['EXCEEDED', 'EVEN'])].reset_index().copy()
notachieved = notachieved.drop_duplicates(subset =['FACILITY'], keep= 'last')
notachieved = notachieved[notachieved['WEEK']>46].copy()
notachieved['BALANCE'] = notachieved['BALANCE'].astype(int)
notachieved['Q3 CURR'] = notachieved['Q3 CURR'].astype(int)
notachieved['Q4 CURR'] = notachieved['Q4 CURR'].astype(int)
notachieved['TXML'] = notachieved['TXML'].astype(int)
notachieved['VL COV (%)'] = notachieved['VL COV (%)'].astype(int)
numb = notachieved['FACILITY'].nunique()
notachieved = notachieved[['DISTRICT','FACILITY', 'Q3 CURR', 'Q4 CURR', 'BALANCE','TXML', 'VL COV (%)']].copy() 
st.write('FACILITIES THAT HAVE NOT ACHIEVED')
st.markdown(f"<h6>{numb} facilities have not achieved so far</h6>", unsafe_allow_html=True)
with st.expander( 'CLICK HERE TO VIEW THEM'):
    st.table(notachieved)
st.divider()
st.subheader('ALL DATA SET')
all = filtered_df[filtered_df['WEEK']>46]
st.write(all)
