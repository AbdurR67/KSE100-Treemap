import pandas as pd
import matplotlib.pyplot as plt
import os
from datetime import datetime
import plotly.express as px
username = os.getlogin()

sectors_df = pd.read_excel(fr'C:\Users\{username}\Dropbox\AR\Sectors_KSE100.xlsx') #excel file location

url = 'https://dps.psx.com.pk/indices/KSE100'
data = pd.read_html(url)[0]
data.columns = ['SYMBOL', 'NAME', 'LDCP', 'CURRENT', 'CHANGE', 'CHANGE (%)', 'IDX WTG (%)', 'IDX POINT', 'VOLUME', 'FREEFLOAT (M)', 'MARKET CAP (M)']
data['SYMBOL'] = data['SYMBOL'].str.replace(r'(XD|XB|XR)', '', regex=True)
data['Change (%)'] = data['CHANGE (%)'].str.rstrip('%').astype(float)
data['CHANGE (%)2'] = data['CHANGE (%)'] #Just to display with % sign

data = data[data['VOLUME'] != 0] #in case want to make box size based on volume

sectors_df['Symbol'] = sectors_df['Symbol'].str.split(', ')
sectors_df = sectors_df.explode('Symbol')
merged_data = pd.merge(data, sectors_df, left_on='SYMBOL', right_on='Symbol', how='left')


fig = px.treemap(merged_data, path=['Sector', 'SYMBOL'], values='MARKET CAP (M)',
                 color='Change (%)', hover_data=['NAME', 'MARKET CAP (M)', 'CHANGE (%)2'],
                 color_continuous_scale=['red', '#414554', 'green'], color_continuous_midpoint=0,
                 range_color=[-3, 3])


fig.update_traces(marker_line_color='#242c34', marker_line_width=1.5)
fig.update_coloraxes(showscale=False)

# To show Values also
fig.update_traces(texttemplate='%{label}<br>%{customdata[2]}')
fig.update_traces(textposition='middle center', selector=dict(type='treemap'))

fig.update_traces(root_color='#242c34', selector=dict(type='treemap'))

fig.update_layout(
    title='KSE100 - Powered by mettisglobal.news',
    title_font_size=12,
    title_x=0.5,
    title_y=0.987,
    title_font_color='white',
    margin=dict(t=0, l=0, r=0, b=0),
    font_size=14,
    plot_bgcolor='#242c34',
    paper_bgcolor='#242c34'
)

#fig.update_coloraxes(colorbar_orientation='h', colorbar_len=0.5, colorbar_thickness=20)

fig.write_image('KSE100_Map.png', width=960, height=640, scale=3)
#fig.show()  #can enable this to view interatctive chart on browser
