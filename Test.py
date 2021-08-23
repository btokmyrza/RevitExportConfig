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
def create3dView(doc, view_name):
    vft = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
    vft_3D = []

    for v in vft:
        if v.ViewFamily == ViewFamily.ThreeDimensional:
            vft_3D.append(v)

    v3d_created = View3D.CreateIsometric(doc, vft_3D[0].Id)
    v3d_created.Name = view_name

    return v3d_created


def AddScheduleFilter(scheduleObj, schedule_field, filter_value):
    filterByParameter = BuiltInParameter.SHEET_NAME
    scheduleFilterType = ScheduleFilterType.NotContains

    with Transaction(doc,"Add Filter") as tr:
        tr.Start()

        definition = scheduleObj.Definition
        scheduleField = None

        ids = definition.GetFieldOrder()

        for id in ids:
            param = definition.GetField(id)
            if param.GetName() == schedule_field:
                scheduleField = param.FieldId

        filter = ScheduleFilter(scheduleField, scheduleFilterType, filter_value)
        definition.AddFilter(filter)

        tr.Commit()

# Add a Schedule filter for the Dummy sheets so that they don't show up in the table
schedule_collector = FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()

schedule_name = "Ведомость рабочих чертежей основного комплекта"
schedule_field = "Sheet Name"
filter_value = "Dummy"

Output("Schedules: ")
for s in schedule_collector:
    if s.Name == schedule_name:
        Output("  "+s.Name)
        AddScheduleFilter(s, schedule_field, filter_value)
        Output("Filter Added+")
