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
def ExportDWGwithCustomOptions(document, view, viewName, output_path):
    options = DWGExportOptions()

    options.Colors = ExportColorMode.IndexColors
    options.ExportOfSolids = SolidGeometry.Polymesh
    options.FileVersion = ACADVersion.R2013
    options.HatchPatternsFileName = r"C:\Program Files\Autodesk\Revit 2019\ACADInterop\acdbiso.pat"
    options.HideScopeBox = True
    options.HideUnreferenceViewTags = True
    options.HideReferencePlane = True
    options.LayerMapping = "AIA"
    options.LineScaling = LineScaling.PaperSpace
    options.LinetypesFileName = r"C:\Program Files\Autodesk\Revit 2019\ACADInterop\acdbiso.lin"
    options.PropOverrides = PropOverrideMode.ByEntity
    options.TextTreatment = TextTreatment.Exact
    options.MergedViews = True

    # Export the active view
    views = []
    views.append(view.Id)

    views_collection = List[ElementId](views)

    document.Export(output_path, viewName, views_collection, options)


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
def determineFileSystem(path, file_type_folder, file_name_folder):
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

        export_path = pre_export_path +"\\"+ getSaveTime(sessionId)
        if not os.path.exists(export_path):
            os.makedirs(export_path)

        full_export_path = export_path +"\\"+ file_name_folder
        if not os.path.exists(full_export_path):
            os.makedirs(full_export_path)
    else:
        pre_export_path = working_directory + r"\4_Publication" +"\\"+ rvt_folder_name
        if not os.path.exists(pre_export_path):
            os.makedirs(pre_export_path)

        export_path = working_directory + r"\4_Publication" +"\\"+ rvt_folder_name +"\\"+ file_type_folder
        if not os.path.exists(export_path):
            os.makedirs(export_path)

        _export_path = export_path +"\\"+ getSaveTime(sessionId)
        if not os.path.exists(_export_path):
            os.makedirs(_export_path)

        full_export_path = _export_path +"\\"+ file_name_folder
        if not os.path.exists(full_export_path):
            os.makedirs(full_export_path)

    return (working_directory, rvt_folder_name, full_export_path)

file_type = "DWG"

file_name_detached = Path.GetFileNameWithoutExtension(doc.PathName)
file_name = file_name_detached.split("_detached")[0]
file_name = file_name.split("_отсоединено")[0]

working_directory = determineFileSystem(revitFilePath, file_type, file_name)[0]
rvt_folder_name = determineFileSystem(revitFilePath, file_type, file_name)[1]
the_export_path = determineFileSystem(revitFilePath, file_type, file_name)[2]

Output("Session ID : " + sessionId + "\n")
Output("The working directory is: " + working_directory + "\n")
Output("The revit folder name: " + rvt_folder_name + "\n")
Output("Export path: " + the_export_path)
#########################################################################################

# Turn off the Linked Files
linked_files_collector = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()

for el in linked_files_collector:
    linkDoc = el.GetLinkDocument()
    type = doc.GetElement(el.GetTypeId())
    type.Unload(None)
#########################################################################################

# find the sheets
coll = FilteredElementCollector(doc)
allSheets = coll.OfClass(ViewSheet).ToElements()

group_names = ("BI_Раздел проекта", "BI_Раздел_проекта", "BI Раздел проекта", "BI_раздел_проекта", "Раздел проекта")

Output("\nSheets in this document (Original Name vs Save Name): ")
for sheet in allSheets:

    # Check if the sheet is printable:
    if sheet.CanBePrinted == True:
        pass
    else:
        Output("This view, <"+ sheet.SheetNumber+"-"+sheet.Name +"> is not printable!")
        continue
    # BI Group naming convention?
    #saveName = file_name[::-1].split('_', 1)[-1][::-1] + "_DR_"+sheet.SheetNumber
    #saveName = file_name +"_DR_"+ sheet.SheetNumber

    # Default naming convention
    saveName = sheet.SheetNumber +"-"+ sheet.Name
    saveName = saveName.split('+')[0]

    if "/" in saveName:
        saveName = saveName.replace("/", ".")

    if "\\" in saveName:
        saveName = saveName.replace("\\", "%")

    if "*" in saveName:
        saveName = saveName.replace("*", "^")

    if ":" in saveName:
        saveName = saveName.replace(":", ";")

    if "?" in saveName:
        saveName = saveName.replace("?", "!")

    if "<" in saveName:
        saveName = saveName.replace("<", "(")

    if ">" in saveName:
        saveName = saveName.replace(">", ")")

    if "\"" in saveName:
        saveName = saveName.replace("\"", "\'")

    if "|" in saveName:
        saveName = saveName.replace("|", "&")


    Output("  "+sheet.SheetNumber+"-"+sheet.Name)

    subgroup = ""

    for param in sheet.GetOrderedParameters():
        if any([x in param.Definition.Name for x in group_names]) == True:
            if param.HasValue:
                subgroup = param.AsString()
                Output("\t-->"+subgroup)
                break

    final_export_path = the_export_path + "\\" + subgroup
    if not os.path.exists(final_export_path):
        os.makedirs(final_export_path)


    saveName_final = saveName.split("_detached")[0]
    saveName_final = saveName_final.split("_отсоединено")[0]

    ExportDWGwithCustomOptions(doc, sheet, saveName_final, final_export_path)
