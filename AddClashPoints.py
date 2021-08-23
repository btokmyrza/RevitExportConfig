# -*- coding: UTF-8 -*-
# This section is common to all Python task scripts.
import clr
import System
import os
import sys

'''
To install 3rd party modules:
- Uninstall IronPython 2.7.3 that comes with Revit be default.
- Download and Install IronPython 2.7.9 (Google it!)
- Add IronPython paths to system variables (variable name = Path):
    C:\Program Files\IronPython 2.7\
    C:\Program Files\IronPython 2.7\Scripts\
- To do the installation of the needed module open cmd.exe as admin:
    ipy -X:Frames -m ensurepip
    ipy -X:Frames -m pip install Package
'''

sys.path.append(r"C:\Program Files\IronPython 2.7\Lib")
sys.path.append(r"C:\Program Files\IronPython 2.7\Lib\site-packages")
import natsort

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import *
from Autodesk.Revit.ApplicationServices import *
from Autodesk.Revit.Attributes import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.DB.Structure import *

from System.Collections.Generic import List
from System.IO import *

import revit_script_util
from revit_script_util import Output

sessionId = revit_script_util.GetSessionId()
uiapp = revit_script_util.GetUIApplication()

doc = revit_script_util.GetScriptDocument()
revitFilePath = revit_script_util.GetRevitFilePath()

path_to_family = r"\\iasp-dc2\Documents\02_Библиотека\BIM\ClashDetective\BIM-Конфликт.rfa"

# The code above is boilerplate, everything below is all yours.
# You can use almost any part of the Revit API here!
#########################################################################################
def addFamily(X_Position, Y_Position, Z_Position, family_to_add, elementid, level):
    insert_position = XYZ(float(X_Position), float(Y_Position), float(Z_Position))
    element_to_add_to = doc.GetElement(	ElementId(elementid) )
    instance = None

    with Transaction(doc,"Add Clash Point Family") as tr:
        tr.Start()

        if not family_to_add.IsActive:
            family_to_add.Activate()
            doc.Regenerate()

        instance = doc.Create.NewFamilyInstance(insert_position, family_to_add, element_to_add_to, level, StructuralType.NonStructural)

        tr.Commit()

    return instance

def getFamilySymbol(family_symbol_name):
    family_coll = FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements()
    family_symbol = None

    for fsymbol in family_coll:
        if fsymbol.Family.Name == family_symbol_name: # "BIM-Конфликт"
            family_symbol = fsymbol

    return family_symbol

def getLevel(level_name):
    lvl_coll = FilteredElementCollector(doc).OfClass(Level).ToElements()
    level = None

    for lvl in lvl_coll:
        if lvl.Name == level_name:
            level = lvl

    return level

def getElementCoordinates(elementid):
    p = XYZ.Zero
    element = doc.GetElement(ElementId(elementid))
    loc = element.Location

    p = loc.Curve.GetEndPoint(0);
    locationStart = str(p.X) +" "+ str(p.Y) +" "+ str(p.Z)

    p = loc.Curve.GetEndPoint(1);
    locationEnd = str(p.X) +" "+ str(p.Y) +" "+ str(p.Z)

    return (locationStart, locationEnd)

# TOKYO_H_S7_VK.rvt
# Load the family into the file
with Transaction(doc, "Load the Family") as tr:
    tr.Start()
    doc.LoadFamily(path_to_family)
    tr.Commit()

fsymbol_to_add = getFamilySymbol("BIM-Конфликт")
elementid = 639939
level = getLevel("01 Этаж")

# Add the family into the file at specified coordinates
added_family_instance = addFamily("0.0", "0.0", "0.0", fsymbol_to_add, elementid, level)

save_path = r"\\Iaspserv\iasp\2018_TSE_TOKYO_H\01_BIM\1_Design\VK\TOKYO_H_S7_VK_test.rvt"

saOptions = SaveAsOptions();
wsOptions = WorksharingSaveAsOptions()
wsOptions.SaveAsCentral = True
saOptions.SetWorksharingOptions(wsOptions)

doc.SaveAs(save_path, saOptions)
