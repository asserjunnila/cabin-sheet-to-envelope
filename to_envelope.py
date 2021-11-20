
#!/usr/bin/env python
# -*- coding: utf-8 -*-
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape, A5
from reportlab.lib.units import inch, cm
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Table
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import PageBreak
from reportlab import *
import openpyxl
import os
import sys
import random
import json


def parse_cabins(file_name):
  print(file_name)
  wb = openpyxl.load_workbook(file_name)
  #sheet = wb.get_sheet_by_name('hytit')
  sheet = wb.active
  cabins_arr = []
  cabin_arr = []
  current = sheet['B3'].value
  #A=hyttiluokka
  #B=hyttinumero
  #C=sukunimi
  #D=etunimi
  #G=DIN1
  #H=DIN2
  #I=BRE
  #J=LUN
  for row in range(3, sheet.max_row + 1):
    if sheet['A'+str(row)].value is None and sheet['B'+str(row)].value is None and sheet['C'+str(row)].value is None:
      continue
    if sheet['B' + str(row)].value == current or sheet['B' + str(row)].value is None:
      cabin_arr.append([sheet['A' + str(row)].value, sheet['C' + str(row)].value, sheet['D' + str(row)].value,
      sheet['G' + str(row)].value, sheet['H' + str(row)].value, sheet['I' + str(row)].value,
      sheet['J' + str(row)].value])
    else:
      cabins_arr.append(cabin_arr)
      current = sheet['B' + str(row)].value
      cabin_arr = [[sheet['A' + str(row)].value, sheet['C' + str(row)].value, sheet['D' + str(row)].value,
      sheet['G' + str(row)].value, sheet['H' + str(row)].value, sheet['I' + str(row)].value,
      sheet['J' + str(row)].value]]
      if row == sheet.max_row:
        cabin_arr.append([sheet['A' + str(row)].value, sheet['C' + str(row)].value, sheet['D' + str(row)].value,
        sheet['G' + str(row)].value, sheet['H' + str(row)].value, sheet['I' + str(row)].value, sheet['J' + str(row)].value])
        cabins_arr.append(cabin_arr)
  return cabins_arr


def create_table(cabin_arr, styleSheet):
  H0 = Paragraph('''<b>Hyttiluokka</b>''', styleSheet["BodyText"])
  H1 = Paragraph('''<b>Sukunimi</b>''', styleSheet["BodyText"])
  H2 = Paragraph('''<b>Etunimi</b>''', styleSheet["BodyText"])
  H3 = Paragraph('''<para align=center><b>I 1</b></para>''', styleSheet["BodyText"])
  H4 = Paragraph('''<para align=center><b>I 2</b></para>''', styleSheet["BodyText"])
  H5 = Paragraph('''<para align=center><b>A</b></para>''', styleSheet["BodyText"])
  H6 = Paragraph('''<para align=center><b>L</b></para>''', styleSheet["BodyText"])

  # TODO: fonts, alignment etc.
  data = []
  header = [H0, H1, H2, H3, H4, H5, H6]

  data.append(header)
  for passenger in cabin_arr:
    data.append(passenger)

  t = Table(data, style=[
    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
    ('ALIGN', (3, 0), (-1, -1), 'CENTER')
  ])

  t._argW[0] = 3.5 * cm
  for i in range(1, 7):
    if i < 3:
      t._argW[i] = 5 * cm
    else:
      t._argW[i] = 1 * cm
  return t

def createPDF(CABINS_XLSX, OUTPUT_FILE):

  # try:  
  #   with open('jokes.json') as json_file:
  #     jokes = json.load(json_file)
  # except:
  #   print("jokes.json not found and script aborted")
  #   quit()

  jokesStyle = ParagraphStyle('yourtitle',
                          fontName="Helvetica-Oblique",
                          fontSize=15,
                          parent=getSampleStyleSheet()['Heading2'],
                          alignment=1,
                          spaceAfter=14,
                          textColor="red")                            #backColor = "rgb(239, 20, 36)",
  # Header image
  # Picture name given as argument
  try:
    picture_file_name = sys.argv[1]
  except:
    print("pls give picture file name as an argument and try again")
    quit()

  I = Image(picture_file_name)
  I.drawHeight = 6.5 * cm
  I.drawWidth = 17.5 * cm

  cabins = parse_cabins(CABINS_XLSX)

  #doc = SimpleDocTemplate(OUTPUT_FILE, pagesize=landscape(A5), topMargin=1, bottomMargin=0)
  doc = SimpleDocTemplate(OUTPUT_FILE, pagesize=landscape(A5), topMargin=1, bottomMargin=0)
  # container for the 'Flowable' objects
  elements = []

  styleSheet = getSampleStyleSheet()
  # Add all elements to doc

  try:  
    with open('notes.json') as json_file:
      notes = json.load(json_file)
  except:
    print("notes.json not found and script aborted")
    quit()

  for cabin in cabins:
    elements.append(I)
    elements.append(create_table(cabin, styleSheet))
    # elements.append(Paragraph("Rannekkeiden kontrolliosat toimivat Avajaisshow'ssa arpalippuina! Pääpalkintona "
    #                         "arvonnassa on hulppea Suite-hytti! Katso ohjelmasta lisätiedot!", styleSheet[
    #                             "Heading4"]))
    elements.append(Paragraph(notes['heading'], styleSheet[
                                "Heading4"]))
    # elements.append(
    #     Paragraph("Laivayhtiö veloittaa asiakkailta alkaen 100€ hyteistä, jotka on sotkettu tai joissa on tupakoitu.",
    #             styleSheet["BodyText"]))
    elements.append(
        Paragraph(notes['first'],
                styleSheet["BodyText"]))

    # elements.append(Paragraph("Omien ja Tax Freesta ostettujen alkoholijuomien käyttö laivalla ehdottomasti kielletty.",
    #                         styleSheet["BodyText"]))
    elements.append(Paragraph(notes['second'],
                            styleSheet["BodyText"]))

    # elements.append(
    #     Paragraph("Älä laita hyttikorttiasi puhelimen suojakoteloon, sillä kotelon magneetti vaurioittaa korttia!",
    #             styleSheet["BodyText"]))
    elements.append(
        Paragraph(notes['third'],
                styleSheet["BodyText"]))

    elements.append(Paragraph(" ",styleSheet["BodyText"]))
    elements.append(Paragraph(" ",styleSheet["BodyText"]))


    elements.append(Paragraph('"'+notes['quote'][random.randrange(0,len(notes['quote']))]+'"', jokesStyle))

    elements.append(PageBreak())


  # write the document to disk
  doc.build(elements)

def main():
  dirpath = os.getcwd()
  files = []
  #which files?

  # r=root, d=directories, f = files
  for r, d, f in os.walk(dirpath):
    for file in f:
      if file.endswith('.xlsx') and '$' not in file and "lyyti" not in file.lower() and "muutokset" not in file.lower():
        files.append(os.path.join(r, file))
  
  for xlsx in files:
    path = os.getcwd()
    folder_name = "envelope_print"
    #print("/" + folder_name + "/envelope_print_" + domain + "_" + xlsx.split('/')[-1].split('.')[0]+".pdf")
    try:
      os.mkdir(folder_name)
    except:
      pass
      print("envelope_print folder already exists so let's use that one")

    createPDF(xlsx, path + "/" + folder_name + "/" + "envelope_print_" + xlsx.split('/')[-1].split('.')[0]+".pdf")

main()