import os
from flask import Flask, flash, request, redirect, render_template,  Response
from werkzeug.utils import secure_filename
from io import BytesIO

from zipfile import ZipFile
#other tools useful in extracting the information from our document
import re
#to pretty print our xml:
import xml.dom.minidom
from xml.etree import ElementTree
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw

app=Flask(__name__)

app.secret_key = "secret key"
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

path = os.getcwd()
# file Upload
UPLOAD_FOLDER = os.path.join(path, 'uploads')

if not os.path.isdir(UPLOAD_FOLDER):
    os.mkdir(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


ALLOWED_EXTENSIONS = set(['docx'])
imageIO = BytesIO()
background = Image.open(path + '/img/Environment.png')

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def draw_rotated_text(image, angle, xy, text, fill, *args, **kwargs):
    """ Draw text at an angle into an image, takes the same arguments
        as Image.text() except for:

    :param image: Image to write text into
    :param angle: Angle to write text at
    """
    # get the size of our image
    width, height = image.size
    max_dim = max(width, height)

    # build a transparency mask large enough to hold the text
    mask_size = (max_dim * 2, max_dim * 2)
    mask = Image.new('L', mask_size, 0)

    # add text to mask
    draw = ImageDraw.Draw(mask)
    draw.text((max_dim, max_dim), text, 255, *args, **kwargs)

    if angle % 90 == 0:
        # rotate by multiple of 90 deg is easier
        rotated_mask = mask.rotate(angle)
    else:
        # rotate an an enlarged mask to minimize jaggies
        bigger_mask = mask.resize((max_dim*8, max_dim*8),
                                  resample=Image.BICUBIC)
        rotated_mask = bigger_mask.rotate(angle).resize(
            mask_size, resample=Image.LANCZOS)

    # crop the mask to match image
    mask_xy = (max_dim - xy[0], max_dim - xy[1])
    b_box = mask_xy + (mask_xy[0] + width, mask_xy[1] + height)
    mask = rotated_mask.crop(b_box)

    # paste the appropriate color, with the text transparency mask
    color_image = Image.new('RGBA', image.size, fill)
    image.paste(color_image, mask)

def generateFoorptint(file):
    #Analize the metadata info if the file is readable:
    #TODO: Add a verification flow to block non Word documents.
    #specific to extracting information from word documents
    #filePathDoc = os.path.join('uploads', filename)
    #document = ZipFile(BytesIO(filename).encode())
    #file_like_object = BytesIO(file)
    document = ZipFile(file)

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
    #print('Nº pagines: ',pagesTemp[0].firstChild.nodeValue)

    totalTimeTemp = xmlexample.getElementsByTagName("TotalTime")
    totalTime = int(totalTimeTemp[0].firstChild.nodeValue)
    #print('Temps total edició (minuts): ',totalTimeTemp[0].firstChild.nodeValue)

    wordsTemp = xmlexample.getElementsByTagName("Words")
    words = int(wordsTemp[0].firstChild.nodeValue)
    #print('Nº paraules: ',wordsTemp[0].firstChild.nodeValue)

      #Calculations for Co2 footprint:
    #From PC Electricity with TotalTime variable:
    PCPower = 0.250 #in kW
    CO2pkW = 0.25 #in kgCO2/kwh
    CO2Elec = PCPower * CO2pkW * totalTime/60 #in kgCO2
    #print(CO2Elec)
    #From paper usage from printing with nº pages variable:
    CO2pPaperW = 1.84 #kgCO2/kg
    CO2pPaperR = 0.61 #kgCO2/kg

    PaperWeight = 80*pages*0.06237/1000 #kg

    CO2PaperSingleW = PaperWeight*CO2pPaperW #Kg CO2
    CO2PaperSingleR = PaperWeight*CO2pPaperR #Kg CO2
    CO2PaperDoubleW = CO2PaperSingleW/2 #Kg CO2
    CO2PaperDoubleR = CO2PaperSingleR/2 #Kg CO2

    #print(CO2PaperSingleW)
    #print(CO2PaperSingleR)
    #print(CO2PaperDoubleW)
    #print(CO2PaperDoubleR)

    #Scalate the number of KM:
    CarCO2Cost = 120.4 #acording to: https://www.eea.europa.eu/data-and-maps/indicators/average-co2-emissions-from-motor-vehicles/assessment-1
    CO2Total= (CO2Elec + CO2PaperSingleW)*1000
    kmRunnedRaw = CO2Total / CarCO2Cost
    kmRunned = int(kmRunnedRaw) + 1
    #print(kmRunned)

    if kmRunned <= 1:
      kmRunnedMax = 1
    elif kmRunned <= 10:
      kmRunnedMax = 10
    elif kmRunned <= 100:
      kmRunnedMax = 100
    else:
      kmRunnedMax = 1000

    #kmRunnedRaw

    #CO2Total

    #savings = [ 100-((CO2PaperDoubleW + CO2Elec)*100 / (CO2PaperSingleW + CO2Elec)),  100-((CO2PaperSingleR + CO2Elec)*100  / (CO2PaperSingleW + CO2Elec)),  100-((CO2PaperDoubleR + CO2Elec)*100  / (CO2PaperSingleW + CO2Elec)),  100-(CO2Elec*100 / (CO2PaperSingleW + CO2Elec))]


    #uploaded = files.upload()

    background = Image.open(path + '/img/Environment.png')
    item1 = Image.open(path + '/img/car.png').resize((400,400))
    item2 = Image.open(path + '/img/co2.png').resize((400,400))
    item3 = Image.open(path + '/img/reduction.png').resize((400,400))

    background.paste(item1, (280,270), item1)
    background.paste(item2, (-30,270), item2)
    background.paste(item3, (640,270), item3)



    font = ImageFont.truetype("Roboto-MediumItalic.ttf", size=20)
    fonttitle = ImageFont.truetype("Roboto-Bold.ttf", size=70)
    fontgraph = ImageFont.truetype("Roboto-Bold.ttf", size=35)

    # draw.text((x, y),"Sample Text",(r,g,b))
    output=ImageDraw.Draw(background)
    output.text((150, 710),"This document was produced in " + str(totalTime) +
                " minutes which contains " + str(pages) +
                " pages and " + str(words)  +  " words",(45,106,79),font=font)
    text = str("{:.1f}".format(kmRunnedRaw))
    w, h = output.textsize(text, font=fonttitle)
    output.text(((480-w/2),(460-h/2)),text,(45,106,79),font=fonttitle)
    text = str("{:.1f}".format(CO2Total))
    w, h = output.textsize(text, font=fonttitle)
    output.text(((170-w/2),(460-h/2)),text,(216,243,220),font=fonttitle)


    max_dim = max(200, 200)
    # build a transparency mask large enough to hold the text
    mask_size = (max_dim * 2, max_dim * 2)
    mask = Image.new('L', mask_size, 0)

    offset = 1
    shift = 3

    text = str(round(100-((CO2PaperDoubleW + CO2Elec)*100 / (CO2PaperSingleW + CO2Elec)), 1)) + "%"
    w, h = output.textsize(text, font=fonttitle)
    draw_rotated_text(background, 90, (705-offset+shift,590), text, (245, 245, 245), font=fontgraph)

    text = str(round(100-((CO2PaperSingleR + CO2Elec)*100  / (CO2PaperSingleW + CO2Elec)), 1)) + "%"
    w, h = output.textsize(text, font=fonttitle)
    draw_rotated_text(background, 90, (780-offset*2+shift,590), text, (66, 66, 66), font=fontgraph)

    text = str(round(100-((CO2PaperDoubleR + CO2Elec)*100  / (CO2PaperSingleW + CO2Elec)), 1)) + "%"
    w, h = output.textsize(text, font=fonttitle)
    draw_rotated_text(background, 90, (855-offset*3+shift,590), text, (66, 66, 66), font=fontgraph)

    text = str(round(100-(CO2Elec*100 / (CO2PaperSingleW + CO2Elec)), 1)) + "%"
    w, h = output.textsize(text, font=fonttitle)
    draw_rotated_text(background, 90, (930-offset*4+shift,590), text, (66, 66, 66), font=fontgraph)


    background.save(imageIO, 'PNG')
    imageIO.seek(0)

@app.route('/')
def upload_form():
    return render_template('upload.html')


@app.route('/', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No file selected for uploading')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_like_object = file.stream._file
            #file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            generateFoorptint(file_like_object)
            flash('File successfully uploaded')
            return Response(imageIO , mimetype="image/png",
                                headers={"Content-Disposition":
                                "attachment;filename=footprint.png"}

            )
            #return redirect('/')
        else:
            flash('Try again: Allowed file types are only docx')
            return redirect(request.url)

if __name__ == "__main__":
    app.run(host = '127.0.0.1',port = 5000, debug = False)
