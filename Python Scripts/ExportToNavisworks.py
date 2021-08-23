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

from System.IO import *
from System.Collections.Generic import List

import revit_script_util
from revit_script_util import Output

sessionId = revit_script_util.GetSessionId()
uiapp = revit_script_util.GetUIApplication()

doc = revit_script_util.GetScriptDocument()
revitFilePath = revit_script_util.GetRevitFilePath()

# The code above is boilerplate, everything below is all yours.
# You can use almost any part of the Revit API here!

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

        full_export_path = pre_export_path
        if not os.path.exists(full_export_path):
            os.makedirs(full_export_path)
    else:
        pre_export_path = working_directory + r"\4_Publication" +"\\"+ rvt_folder_name
        if not os.path.exists(pre_export_path):
            os.makedirs(pre_export_path)

        export_path = working_directory + r"\4_Publication" +"\\"+ rvt_folder_name +"\\"+ file_type_folder
        if not os.path.exists(export_path):
            os.makedirs(export_path)

        full_export_path = export_path
        if not os.path.exists(full_export_path):
            os.makedirs(full_export_path)

    return (working_directory, rvt_folder_name, full_export_path)

file_type = "NWC"
working_directory = determineFileSystem(revitFilePath, file_type)[0]
rvt_folder_name = determineFileSystem(revitFilePath, file_type)[1]
the_export_path = determineFileSystem(revitFilePath, file_type)[2]

Output("\nSession ID : " + sessionId)
Output("The working directory is: " + working_directory)
Output("Revit folder name: " + rvt_folder_name +"\n")
#########################################################################################
# Define Navisworks Export Options
opt = NavisworksExportOptions()
opt.ExportLinks = False
opt.ExportUrls = False
opt.FindMissingMaterials = False
opt.ExportRoomAsAttribute = False
opt.ExportRoomGeometry = False
opt.DivideFileIntoLevels = True
opt.Coordinates = NavisworksCoordinates.Shared
opt.ExportScope = NavisworksExportScope.Model

# Get all 3D Views in the document
coll = FilteredElementCollector(doc)
allViews = coll.OfClass(View3D).ToElements()
all3DViews= []
v3d = None

for view in allViews:
    if view.IsTemplate == False:
        all3DViews.append(view)


Output("The List of 3D Views in this file: ")
for view3D in all3DViews:
    Output("   " + view3D.Name)


Output("\nExporting the whole project...\n")
#########################################################################################
# *Turn off the Linked Files*
linked_files_collector = FilteredElementCollector(doc).OfClass(RevitLinkType).ToElements()

#Output("Linked files in this document: ")
for el in linked_files_collector:
    #linkDoc = el.GetLinkDocument()
    fileRef = el.GetExternalFileReference()

    #Linked Document name
    #doc_name = el.Name
    #type.Unload(None)

    #TO GET STATUS
    status =fileRef.GetLinkedFileStatus().ToString();

    #TO GET REFERENCETYPE
    rTYPE = el.get_Parameter(BuiltInParameter.RVT_LINK_REFERENCE_TYPE).AsValueString()

    #Output(" "+doc_name+" ("+status+") "+"["+rTYPE+"]")
    if status == "Loaded" and rTYPE == "Overlay":
        el.Unload(None)
    elif status == "Loaded" and rTYPE == "Attachment":
        opt.ExportLinks = True
#########################################################################################
# *Remove "IMPORTED" CAD files*
imported_dwg_collector = FilteredElementCollector(doc).OfClass(ImportInstance)
DWGsToDelete = []
for dwg in imported_dwg_collector:
    DWGsToDelete.append(dwg.Id)

DWGsToDeleteIList = List[ElementId](DWGsToDelete)

with Transaction(doc,"Delete Imported DWGs") as tr:
    tr.Start()
    doc.Delete(DWGsToDeleteIList)
    tr.Commit()
#########################################################################################
# Do the exporting
saveName = doc.Title.split("_detached")[0]
saveName = saveName.split("_отсоединено")[0]
saveName = saveName.split(".rvt")[0]

Output("Export path: " + the_export_path +"\\"+saveName)

doc.Export(the_export_path, saveName+".nwc", opt)
doc.Close(0)
