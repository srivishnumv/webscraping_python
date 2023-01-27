from fpdf import FPDF
from bs4 import BeautifulSoup
import requests,openpyxl
import classpag as s

import PyPDF2
from win32com import client

excel = openpyxl.Workbook()
print(excel.sheetnames)
sheet = excel.active
sheet.title = 'Top Rated Movies'
print(excel.sheetnames)
sheet.append(['Movie Rank','Movie Name', 'Year of Release', 'IMDB Rating'])
s.IMDB(pdf_name='IMDB')

def write_to_file_and_pdf(text):
    with open("IMDB.pdf", "a") as f:
        f.write(text)
    pdf.cell(200, 10, txt=text, ln=1, align='C')

pdf = FPDF()
pdf.add_page()
pdf.set_font("Arial", size = 15)

try:
     source = requests.get('https://www.imdb.com/chart/top/')
     source.raise_for_status()
except Exception as e:
   print(e)
soup=BeautifulSoup(source.text,'html.parser')
     #print(soup.encode("utf-8"))
     
movies=soup.find('tbody', class_="lister-list").find_all('tr')
     # print(len(movies))


for movie in movies:
        
   name = movie.find('td', class_="titleColumn").a.text
        
   rank = movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0]

   year = movie.find('td', class_="titleColumn").span.text.strip('()')

   rating = movie.find('td', class_="ratingColumn imdbRating").strong.text

   write_to_file_and_pdf(name)
   print(rank,name,year,rating) 

pdf.output("IMDB.pdf")

pdf_file = "IMDB.pdf"
watermark = "wmark1.pdf"
merged_file = "merged.pdf"

input_file = open(pdf_file,'rb')
input_pdf = PyPDF2.PdfReader(input_file)

watermark_file = open(watermark,'rb')
watermark_pdf = PyPDF2.PdfReader(watermark_file)

pdf_page = input_pdf.pages[0]

watermark_page = watermark_pdf.pages[0]

pdf_page.merge_page(watermark_page)

output = PyPDF2.PdfWriter()

output.add_page(pdf_page)

merged_file = open(merged_file,'wb')
output.write(merged_file)

merged_file.close()
watermark_file.close()
input_file.close()




        
   

