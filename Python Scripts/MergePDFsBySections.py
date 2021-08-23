import os, sys
import time
from itertools import groupby
import shutil

from PyPDF2 import PdfFileMerger
import natsort


path = sys.argv[1]

if "I:" in path:
	path = path.replace("I:", r"\\Iaspserv\iasp")
elif "B:" in path:
	path = path.replace("B:", r"\\Iaspserv\big")
elif "O:" in path:
	path = path.replace("O:", r"\\Zhasulan-ov\овик")

if "1_Design" in path:
	path = path.replace("1_Design", "4_Publication")
elif "1_PROJECT" and "RVT" in path:
	path = path.split("RVT")[0] + "PDF"

# Remove Dummmy sheets
print("Removing pdf sheets conataining the word <Dummy>")
for dirpath, dirnames, filenames in os.walk(path):
	for name in filenames:
		file_path = dirpath +"\\"+name
		if "Dummy" in name:
			os.remove(file_path)

print("Dummy sheets Removed!")

pdf_files = []
rvt_name = ""

print("Renaming the pdf sheets...")
for dirpath, dirnames, filenames in os.walk(path):
	for name in filenames:
		subgroup = " ()"

		if name.endswith((".pdf")) and "_detached" in name:
			new_filename = name
			new_filename = new_filename.replace("Sheet_ ","").replace(".rvt","")
			rvt_name = new_filename.split(" - ")[0]

			if dirpath.split(rvt_name.replace("_detached",""))[1]:
				subgroup = dirpath.split(rvt_name.replace("_detached",""))[1].replace("\\"," (")+")"

			new_filename = new_filename.replace("_detached",subgroup)
			os.rename(dirpath +"\\"+ name, dirpath +"\\"+ new_filename)


for dirpath, dirnames, filenames in os.walk(path):
	for name in filenames:
		if name.endswith((".pdf")) and "(" in name and ")" in name and " - " in name:
			subgroup = name[name.find("(")+1 : name.find(")")]
			file = name.split(" - ")
			file.append(subgroup)
			file.insert(0,dirpath+"\\")
			pdf_files.append(file)

pdf_files_grouped = [list(g) for k, g in groupby(pdf_files, lambda x: x[1])]

print("Processing files ("+str(len(pdf_files_grouped))+"):")
for pdf_group in pdf_files_grouped:
	sorted_pdf_group = natsort.natsorted(pdf_group, key=lambda x: x[2])

	merger = PdfFileMerger()
	pdfs_to_merge = []

	export_path = ""
	export_name = ""

	for pdf in sorted_pdf_group:
		full_path = " - ".join(pdf[:-1])
		full_path = full_path.replace(" - ","",1)
		export_name = full_path.rsplit("\\",1)[1].split(" - ")[0] + ".pdf"
		export_path = full_path.split("PDF\\")
		export_path = export_path[0] +"PDF\\"+ export_path[1].split("\\")[0]
		pdfs_to_merge.append(full_path)

	print("  "+export_name)

	page_counter = 0
	for pdf in pdfs_to_merge:
		merger.append(pdf)
		bookmark = pdf.rsplit("\\",1)[1].replace(".pdf","")
		merger.addBookmark(bookmark, page_counter)
		page_counter = page_counter + 1
	merger.write(export_path+"\\"+export_name)
	merger.close()

'''
# Remove Folders with pdf sheets
print("Removing the unneeded folders with pdf sheets...")
for dirpath, dirnames, filenames in os.walk(path):
	for folder in dirnames:
		if folder=="PDF" and "#" not in dirpath:
			for time_folder in os.listdir(dirpath+"\\"+folder):
				export_path = dirpath+"\\"+folder+"\\"+time_folder
				remove_folders = [name for name in os.listdir(export_path) if os.path.isdir(os.path.join(export_path, name))]
				for folder in remove_folders:
					remove_folder_path = export_path +"\\"+ folder
					shutil.rmtree(remove_folder_path)
print("Done!")
'''
