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

sections = ("AR", "EM", "GP", "KG", "KM", "OV", "PB", "SH", "SPS", "SS", "VK")

output_path = r"\\iasp-dc2\Documents\05_Soft\Revit Batch Processor\RevitExportConfig\temp\revit_file_list_"

for section in sections:
	for file in revit_files:
		if section in file:
			with open(output_path+section+".txt", 'a', encoding='utf-8') as f:
				f.write(file +"\n")
