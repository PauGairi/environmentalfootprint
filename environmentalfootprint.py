#Upload files:
from google.colab import files

uploaded = files.upload()

#print the file uploaded general info:
for fn in uploaded.keys():
  print('User uploaded file "{name}" with length {length} bytes'.format(
      name=fn, length=len(uploaded[fn])))

#Analize the metadata info if the file is readable:
#TODO: Add a verification flow to block non Word documents.
#specific to extracting information from word documents
import os
import zipfile
#other tools useful in extracting the information from our document
import re
#to pretty print our xml:
import xml.dom.minidom
from xml.etree import ElementTree

for fn in uploaded.keys():
  document = zipfile.ZipFile(fn)

  uglyXmlapp = xml.dom.minidom.parseString(document.read('docProps/app.xml')).toprettyxml(indent='  ')
  uglyXmlcore = xml.dom.minidom.parseString(document.read('docProps/core.xml')).toprettyxml(indent='  ')

  text_re = re.compile('>\n\s+([^<>\s].*?)\n\s+</', re.DOTALL)
  prettyXmlapp = text_re.sub('>\g<1></', uglyXmlapp)
  prettyXmlcore = text_re.sub('>\g<1></', uglyXmlcore)

  #print(prettyXmlapp)
  #print(prettyXmlcore)

  xmlexample = xml.dom.minidom.parseString(document.read('docProps/app.xml'))

  pagesTemp = xmlexample.getElementsByTagName("Pages")
  pages = int(pagesTemp[0].firstChild.nodeValue)
  print('Nº pagines: ',pagesTemp[0].firstChild.nodeValue)

  totalTimeTemp = xmlexample.getElementsByTagName("TotalTime")
  totalTime = int(totalTimeTemp[0].firstChild.nodeValue)
  print('Temps total edició (minuts): ',totalTimeTemp[0].firstChild.nodeValue)

  wordsTemp = xmlexample.getElementsByTagName("Words")
  words = int(wordsTemp[0].firstChild.nodeValue)
  print('Nº paraules: ',wordsTemp[0].firstChild.nodeValue)

  #Calculations for Co2 footprint:
#From PC Electricity with TotalTime variable:
PCPower = 0.250 #in kW
CO2pkW = 0.25 #in kgCO2/kwh
CO2Elec = PCPower * CO2pkW * totalTime/60 #in kgCO2
print(CO2Elec)
#From paper usage from printing with nº pages variable:
CO2pPaperW = 1.84 #kgCO2/kg
CO2pPaperR = 0.61 #kgCO2/kg

PaperWeight = 80*pages*0.06237/1000 #kg

CO2PaperSingleW = PaperWeight*CO2pPaperW #Kg CO2
CO2PaperSingleR = PaperWeight*CO2pPaperR #Kg CO2
CO2PaperDoubleW = CO2PaperSingleW/2 #Kg CO2
CO2PaperDoubleR = CO2PaperSingleR/2 #Kg CO2

print(CO2PaperSingleW)
print(CO2PaperSingleR)
print(CO2PaperDoubleW)
print(CO2PaperDoubleR)

import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.io as pio
import plotly.express as px

#Scalate the number of KM:
CarCO2Cost = 120.4 #acording to: https://www.eea.europa.eu/data-and-maps/indicators/average-co2-emissions-from-motor-vehicles/assessment-1
CO2Total= (CO2Elec + CO2PaperSingleW)*1000
kmRunnedRaw = CO2Total / CarCO2Cost
kmRunned = int(kmRunnedRaw) + 1
print(kmRunned)

if kmRunned <= 1:
  kmRunnedMax = 1
elif kmRunned <= 10:
  kmRunnedMax = 10
elif kmRunned <= 100:
  kmRunnedMax = 100
else:
  kmRunnedMax = 1000

#build graph data:

#labels1 = ['kmem', 'kmtotal']
#values1 = [CO2Total, (kmRunnedMax*CarCO2Cost - CO2Total)]

#labels2 = ['CO2 Edition', 'CO2 compare']
#values2 = [CO2Elec, 1-CO2Elec]

#colors = ['#2D6A4F', '#B7E4C7']

labels1 = ['noth']
values1 = [100]

colors = ['#95D5B2', '#B7E4C7']

fig = go.Figure(data=[go.Pie(labels=labels1, values=values1, hole=0, marker=dict(colors=colors)

)])




fig.update_layout(
    # Add annotations in the center of the donut pies.
    annotations=[dict(text= str("<b>{:.1f}<b>".format(kmRunnedRaw)), x=0.5, y=0.55, font_size=100, showarrow=False, font_color='#2D6A4F'),
                 dict(text='km driven', x=0.5, y=0.35, font_size=40, showarrow=False, font_color='#2D6A4F')
                 ],
)

fig.update_layout({
    'plot_bgcolor': 'rgba(0,0,0,0)',
    'paper_bgcolor': 'rgba(0,0,0,0)'

})

fig.update_layout(showlegend=False)

fig.update_traces(textinfo='none',)

fig.update_layout(
    width=600,
    height=600)

#fig.show()
fig.write_image("car.png")

labels2 = ['noth']
values2 = [100]

colors = ['#2D6A4F', '#B7E4C7']

fig = go.Figure(data=[go.Pie(labels=labels2, values=values2, hole=.0, marker=dict(colors=colors)

)])


fig.update_layout(
    # Add annotations in the center of the donut pies.
    annotations=[dict(text= str("<b>{:.1f}</b>".format(CO2Total)), x=0.5, y=0.55, font_size=100, showarrow=False, font_color='#D8F3DC'),
                 dict(text="{}\u2082".format('g CO'), x=0.5, y=0.35, font_size=40, showarrow=False, font_color='#D8F3DC'),
                 ],
)

fig.update_layout({
    'plot_bgcolor': 'rgba(0,0,0,0)',
    'paper_bgcolor': 'rgba(0,0,0,0)'

})

fig.update_layout(showlegend=False)

fig.update_traces(textinfo='none',)

fig.update_layout(
    width=600,
    height=600)

#fig.show()
fig.write_image("co2.png")

colors = ['#f94144','#f9844a','#f9c74f','#90be6d','#43aa8b']

savings = [ 100-((CO2PaperDoubleW + CO2Elec)*100 / (CO2PaperSingleW + CO2Elec)),  100-((CO2PaperSingleR + CO2Elec)*100  / (CO2PaperSingleW + CO2Elec)),  100-((CO2PaperDoubleR + CO2Elec)*100  / (CO2PaperSingleW + CO2Elec)),  100-(CO2Elec*100 / (CO2PaperSingleW + CO2Elec))]

fig = go.Figure(data=[go.Bar(
    x=['Normal<br>Paper<br>2 Sides', 'Recycled<br>Paper<br>1 Side',
       'Recycled<br>Paper<br>2 Sides', 'Only<br>Digital'],
    y=savings,
    marker_color=colors,# marker color can be a single color value or an iterable
    text=["<b>" + str(round(item, 1)) + " %</b>" for item in savings]
)])
fig.update_layout({
    'plot_bgcolor': 'rgba(0,0,0,0)',
    'paper_bgcolor': 'rgba(0,0,0,0)',
})


fig.update_layout(
    yaxis=dict(
        title='% footprint reduction',
        titlefont_size=22,
        tickfont_size=22,
        titlefont_color='#081C15',
        tickfont_color='#081C15'
    ),

    xaxis=dict(
        tickfont_size=22,
        tickfont_color='#081C15',
        tickangle = 0,

    ),

    barmode='stack',
    bargap=0.0, # gap between bars of adjacent location coordinates.
    bargroupgap=0.05 # gap between bars of the same location coordinate.
)


fig.update_traces(textfont_size=45, textangle=90, textposition="inside", cliponaxis=False)

fig.update_layout(showlegend=False)

fig.update_layout(
    width=600,
    height=600)

#fig.show()
fig.write_image("reduction.png")

from PIL import Image
#uploaded = files.upload()

background = Image.open('Environment.png')
item1 = Image.open('car.png').resize((400,400))
item2 = Image.open('co2.png').resize((400,400))
item3 = Image.open('reduction.png').resize((400,400))

background.paste(item1, (280,270), item1)
background.paste(item2, (-30,270), item2)
background.paste(item3, (640,270), item3)

from PIL import ImageFont
from PIL import ImageDraw

font = ImageFont.truetype("Roboto-MediumItalic.ttf", size=20)

# draw.text((x, y),"Sample Text",(r,g,b))
output=ImageDraw.Draw(background)
output.text((150, 710),"This document was produced in " + str(totalTime) +
            " minutes which contains " + str(pages) +
            " pages and " + str(words)  +  " words",(45,106,79),font=font)

background.save("footprint.png")
from IPython.display import Image
display(Image('footprint.png'))
files.download("footprint.png") 
