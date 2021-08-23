# -*- coding: UTF-8 -*-
# This section is common to all Python task scripts.
import clr
import System
import os
from datetime import datetime, timedelta

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import *
from Autodesk.Revit.ApplicationServices import *
from Autodesk.Revit.Attributes import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *

from System.Collections.Generic import List
from System.IO import *

import revit_script_util
from revit_script_util import Output

sessionId = revit_script_util.GetSessionId()
uiapp = revit_script_util.GetUIApplication()

doc = revit_script_util.GetScriptDocument()
revitFilePath = revit_script_util.GetRevitFilePath()

# The code above is boilerplate, everything below is all yours.
# You can use almost any part of the Revit API here!
def ExportDWFwithCustomOptions(document, full_export_path):
    options = DWFExportOptions()

    options.CropBoxVisible = False
    options.ExportingAreas = True
    options.ExportObjectData = True
    options.ExportTexture = True
    options.ImageFormat = DWFImageFormat.Lossless
    options.ImageQuality = DWFImageQuality.High
    options.MergedViews = True
    options.PaperFormat = ExportPaperFormat.Default
    options.PortraitLayout = False
    options.StopOnError  = False

    views = ViewSet()
    # find the sheets
    coll = FilteredElementCollector(doc)
    allSheets = coll.OfClass(ViewSheet).ToElements()

    for sheet in allSheets:
        if sheet.CanBePrinted == True:
            views.Insert(sheet)
        else:
            Output("  This view, <"+ sheet.SheetNumber+"-"+sheet.Name +"> is not printable!")
            continue

    saveName = doc.Title.split("_detached")[0]
    saveName = saveName.split("_отсоединено")[0]
    saveName = saveName.split(".rvt")[0]

    document.Export(full_export_path, saveName, views, options)

def getSaveTime(current_sessionId):
    saveTime = current_sessionId
    saveTime = saveTime.replace("<", "")
    saveTime = saveTime.replace(">", "")
    saveTime = saveTime.replace("T", ".")
    saveTime = saveTime.replace(":", ".")
    saveTime = saveTime.split(".")
    saveTime = saveTime[:-2]

    year  = int(saveTime[0].split("-")[0])
    month = int(saveTime[0].split("-")[1])
    day   = int(saveTime[0].split("-")[2])
    hour  = int(saveTime[1])
    mins  = int(saveTime[2])
    current_time = datetime(year, month, day, hour, mins, 0)
    six_hours_from_now = current_time + timedelta(hours=6)

    saveTime = format(six_hours_from_now, '%Y-%m-%d_%H.%M')

    return saveTime

# Determine the filesystem for saving the exported files
def determineFileSystem(path, file_type_folder):
    working_directory = path
    rvt_folder_name = ""
    # Determine which file system we are working with:
    if "1_Design" in path:
        working_directory = working_directory.split(r"\1_Design")[0]
        rvt_folder_name = path.split(r"\1_Design")[1]
        rvt_folder_name = rvt_folder_name[1:]
        rvt_folder_name = rvt_folder_name.partition("\\")[0]
    elif "1_V RABOTE" in path:
        working_directory = working_directory.split(r"\RVT")[0]
    elif "2_NA VIDACHU" in path:
        working_directory = working_directory.split(r"\RVT")[0]
    else:
        working_directory = working_directory.rsplit('\\', 1)[0]

    # Create/Check the appropriate folders for determined file system:
    if not rvt_folder_name:
        pre_export_path = working_directory +"\\"+ file_type_folder
        if not os.path.exists(pre_export_path):
            os.makedirs(pre_export_path)

        full_export_path = pre_export_path +"\\"+ getSaveTime(sessionId)
        if not os.path.exists(full_export_path):
            os.makedirs(full_export_path)
    else:
        pre_export_path = working_directory + r"\4_Publication" +"\\"+ rvt_folder_name
        if not os.path.exists(pre_export_path):
            os.makedirs(pre_export_path)

        export_path = working_directory + r"\4_Publication" +"\\"+ rvt_folder_name +"\\"+ file_type_folder
        if not os.path.exists(export_path):
            os.makedirs(export_path)

        full_export_path = export_path +"\\"+ getSaveTime(sessionId)
        if not os.path.exists(full_export_path):
            os.makedirs(full_export_path)

    return (working_directory, rvt_folder_name, full_export_path)

file_type = "DWF"
working_directory = determineFileSystem(revitFilePath, file_type)[0]
rvt_folder_name = determineFileSystem(revitFilePath, file_type)[1]
the_export_path = determineFileSystem(revitFilePath, file_type)[2]

Output(" Session ID : " + sessionId + "\n")
Output(" The working directory is: " + working_directory + "\n")
Output(" The revit folder name: " + rvt_folder_name + "\n")
#########################################################################################

Output(" Export path: " + the_export_path)

with Transaction(doc,"DWF Export") as tr:
    tr.Start()
    ExportDWFwithCustomOptions(doc, the_export_path)
    tr.RollBack()
