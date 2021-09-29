#Imports Required Packages from PIL - To manipulate image
from PIL import Image, ImageDraw, ImageFont

#We use pandas for our excel sheet
import pandas as pd

#Import the excel file that contains all the details
data = pd.read_excel("ParticipantList.xlsx")

#Import 'Name' List from the imported .xlsx file
name_list = data['Name'].to_list()

#Determining the length of the list
max_no = len(name_list)

#The Loops for creating certificates in bulk
# for i in name_list:
for idx, i in enumerate(name_list):

    im = Image.open("Certificate.jpg")
    d = ImageDraw.Draw(im)
    text_color = (0, 0, 0)
    bounding_box = [2113, 2721, 7577, 2697]
    x1, y1, x2, y2 = bounding_box  # For easy reading
    font = ImageFont.truetype("Montserrat-SemiBold.ttf", 350, encoding="unic")
    w, h = d.textsize(i.title(), font=font)
    x = (x2 - x1 - w) / 2 + x1
    y = (y2 - y1 - h) / 2 + y1

    # Write the text to the image, where (x,y) is the top left corner of the text
    d.text((x, y), i.title(), align='center', font=font, fill=text_color)
    im.save("Certificate_CyVIT2021_"+i+".pdf")
    print("(%d/%d) Certificate Created for:  %s" % (idx+1, max_no, i.title()))

print("****************************\nCode by Priyanka Giri.\nGitHub: @itspgiri\nThis code was initially made for bulk making of 1500 certificate for CyVIT 2021\n****************************")