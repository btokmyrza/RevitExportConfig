# -*- coding: UTF-8 -*-
import os
import sys

path = sys.argv[1]

if "I:" in path:
	path = path.replace("I:", r"\\Iaspserv\iasp")
elif "B:" in path:
    path = path.replace("B:", r"\\Iaspserv\big")
elif "O:" in path:
    path = path.replace("O:", r"\\Zhasulan-ov\овик")

revit_files = []

exclude_folder_names = ("BI STANDART", "BIM коллизии", "Titul-Stamp", "SMETA", "GOSEXPERTIZA",
                        "OPZ", "POS", "OVOS", "DWG", "INcoming", "OUTgoing",
                        "0_URNA", "Urna", "#", "Revit_temp")

for dirpath, dirnames, filenames in os.walk(path, followlinks=True):
    for name in filenames:
        if name.endswith((".rvt")):
                if any([x in dirpath for x in exclude_folder_names]) == False:
                    full_path = dirpath +"\\"+ name
                    revit_files.append(full_path)

revit_files_part1 = []
revit_files_part2 = []
revit_files_part3 = []

for i in range(len(revit_files)):
	if i>=0 and i<int(len(revit_files)/3):
		revit_files_part1.append(revit_files[i])
	elif i>=int(len(revit_files)/3) and i<int(2*len(revit_files)/3):
		revit_files_part2.append(revit_files[i])
	elif i>=int(2*len(revit_files)/3) and i<int(len(revit_files)):
		revit_files_part3.append(revit_files[i])

output_path1 = r"\\iasp-dc2\Documents\05_Soft\Revit Batch Processor\RevitExportConfig\temp\revit_file_list_part_1.txt"
output_path2 = r"\\iasp-dc2\Documents\05_Soft\Revit Batch Processor\RevitExportConfig\temp\revit_file_list_part_2.txt"
output_path3 = r"\\iasp-dc2\Documents\05_Soft\Revit Batch Processor\RevitExportConfig\temp\revit_file_list_part_3.txt"

with open(output_path1, 'w', encoding='utf-8') as f:
    for item in revit_files_part1:
        f.write(item +"\n")

with open(output_path2, 'w', encoding='utf-8') as f:
    for item in revit_files_part2:
        f.write(item +"\n")

with open(output_path3, 'w', encoding='utf-8') as f:
    for item in revit_files_part3:
        f.write(item +"\n")
