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
schedule_collector = FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()
documentSettings = doc.Settings
groups = documentSettings.Categories

cell_text = ""
old_GOST = "ГОСТ Р 52544-2006"
new_GOST = "ГОСТ 34028-2016"

'''
for s in schedule_collector:
    if "Арматура ВРС" in s.Name:
        Output("Schedule name: "+s.Name)
        table = s.GetTableData()
        section = table.GetSectionData(SectionType.Body)
        for row in range(section.NumberOfRows):
            for column in range(section.NumberOfColumns):
                cell_text = section.GetCellText(row, column)
                if old_GOST in cell_text:
                    Output("The text: "+cell_text)
'''

formulas = []

Output("Schedule fields: ")
for s in schedule_collector:
    if "2_Арматура Ростверк Рм-1.1" in s.Name:
        sdef = s.Definition
        for i in range(sdef.GetFieldCount()):
            field = sdef.GetField(i)
            field_name = field.GetName()
            if field.IsCalculatedField and field_name == "*Обозначение":
                
        break
