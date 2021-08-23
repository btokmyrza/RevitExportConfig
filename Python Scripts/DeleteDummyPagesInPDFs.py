import time
start_time = time.time()

import os
import sys

import natsort
from PyPDF2 import PdfFileWriter, PdfFileReader

path = sys.argv[1]

if "I:" in path:
	path = path.replace("I:", r"\\Iaspserv\iasp")
elif "B:" in path:
    path = path.replace("B:", r"\\Iaspserv\big")
elif "O:" in path:
    path = path.replace("O:", r"\\Zhasulan-ov\овик")

if "1_Design" in path:
    path = path.split("1_Design")[0] + "4_Publication"

exclude_folder_names = ("BI STANDART", "BIM коллизии", "Titul-Stamp", "SMETA", "GOSEXPERTIZA",
                        "OPZ", "POS", "OVOS", "DWG", "INcoming", "OUTgoing", "2_NA VIDACHU",
                        "0_URNA", "Urna", "#", "Revit_temp")

PDFs = []

print("\n************************************************************")
print("Waiting for pdfFactory Pro to finish printing...")
# Wait 2 minutes for pdfFactory Pro to finish printing
time.sleep(7)

print("Scanning for PDF files with the name DELETE...")

for dirpath, dirnames, filenames in os.walk(path, followlinks=True):
    for name in filenames:
        if name.endswith((".pdf")):
            if "DELETE" in name:
                full_path = dirpath +"\\"+ name
                PDFs.append(full_path)

if not PDFs:
	print("No DELETE PDFs found.")
else:
	print("Removing Dummy sheets from PDFs...")

for pdf_file in PDFs:
	infile = PdfFileReader(pdf_file, 'rb')
	output = PdfFileWriter()

	savePath = pdf_file[::-1].split('\\', 1)[-1][::-1]
	saveName = pdf_file[::-1].split('\\', 1)[0][::-1].split("_DELETE")[0]+".pdf"
	export_path = savePath +"\\"+saveName

	print(" "+saveName)

	pages_info = infile.getOutlines()

	# page numbering starts from 0
	pages_to_delete = []
	pages_to_add = []
	sheet_nums = []

	page_titles_delete = {}
	page_titles_add = {}

	counter = 0

	for page in pages_info:
		if "Dummy" in page['/Title']:
			page_titles_delete[page['/Title']] = counter
			sheet_number = page['/Title'].split("Sheet")[1].replace(" ", "")[1:].split("-")[0]
			page_bookmark = page['/Title']
			page_bookmark = page_bookmark.replace("_detached.rvt", "").replace("Sheet", "Лист")
			pages_to_delete.append([sheet_number, counter, page_bookmark])
		else:
			page_titles_add[page['/Title']] = counter
			sheet_number = page['/Title'].split("Sheet")[1].replace(" ", "")[1:].split("-")[0]
			page_bookmark = page['/Title']
			page_bookmark = page_bookmark.replace("_detached.rvt", "").replace("Sheet", "Лист")
			pages_to_add.append([sheet_number, counter, page_bookmark])

		counter += 1

	pages_to_add = natsort.natsorted(pages_to_add, key=lambda x: (x[0],x[1]))

	# Add pages
	for i in pages_to_add:
		p = infile.getPage(i[1])
		output.addPage(p)

	# Add bookmarks
	page_counter = 0
	for i in pages_to_add:
		output.addBookmark(i[2], page_counter)
		page_counter = page_counter + 1


	with open(export_path, 'wb') as f:
		output.write(f)

	# Delete the PDF with Dummy pages
	delete_file = savePath +"\\"+pdf_file[::-1].split('\\', 1)[0][::-1]
	os.remove(delete_file)


print("Done!")

print("The program took %s seconds to run" % (time.time() - start_time))

time.sleep(7)
