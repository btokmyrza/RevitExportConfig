# -*- coding: UTF-8 -*-
# This section is common to all Python task scripts.
import clr
import System
import os
import getpass
import sys
import time
import math
from datetime import datetime, timedelta
from itertools import groupby

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

from System.Collections.Generic import List
from System.IO import *
from Microsoft.Win32 import *

import revit_script_util
from revit_script_util import Output

sessionId = revit_script_util.GetSessionId()
uiapp = revit_script_util.GetUIApplication()

doc = revit_script_util.GetScriptDocument()
revitFilePath = revit_script_util.GetRevitFilePath()

reg_path_PrinterDriverData = r"Software\FinePrint Software\pdfFactory6\FinePrinters\pdfFactory Pro\PrinterDriverData"
reg_path_pdfFactoryPro = r"Software\FinePrint Software\pdfFactory6\FinePrinters\pdfFactory Pro"
printer_name = "pdfFactory Pro"
#######################################################################################################
# The code above is boilerplate, everything below is all yours.
# You can use almost any part of the Revit API here!
def ExportPDF(document, printer, view_set, swidth, sheight, stype, setName):
    printManager = document.PrintManager
    printParameters = printManager.PrintSetup.CurrentPrintSetting.PrintParameters
    printManager.PrintSetup.CurrentPrintSetting = printManager.PrintSetup.InSession
    printSetup = printManager.PrintSetup

    printManager.SelectNewPrintDriver(printer)

    printParameters.ZoomType = ZoomType.Zoom
    printParameters.Zoom = 100
    printParameters.PaperPlacement = PaperPlacementType.Margins
    printParameters.MarginType = MarginType.PrinterLimit
    printParameters.ColorDepth = ColorDepthType.Color
    printParameters.RasterQuality =RasterQualityType.High

    printParameters.HiddenLineViews = HiddenLineViewsType.VectorProcessing
    printParameters.ViewLinksinBlue = False
    printParameters.HideReforWorkPlanes = True
    printParameters.HideUnreferencedViewTags = False
    printParameters.HideCropBoundaries = True
    printParameters.HideScopeBoxes = True
    printParameters.ReplaceHalftoneWithThinLines = False
    printParameters.MaskCoincidentLines = False

    paperSizeArray = printManager.PaperSizes
    current_paper_type = None

    for paper_size in paperSizeArray:
        if paper_size.Name.ToString() == stype:
            printParameters.PaperSize = paper_size
            current_paper_type = stype
            break


    if current_paper_type is None:
        Output("\nThis type of sheet, "+stype+", is not defined in the printer settings\n")


    if int(swidth) >= int(sheight):
        printParameters.PageOrientation = PageOrientationType.Landscape
    else:
        printParameters.PageOrientation = PageOrientationType.Portrait


    printManager.PrintRange = PrintRange.Select
    viewSheetSetting = printManager.ViewSheetSetting
    viewSheetSetting.CurrentViewSheetSet.Views = view_set

    with Transaction(doc,"Create ViewSet & Save Print Settings") as tr:
        tr.Start()

        viewSheetSetting.SaveAs("Set_"+setName)
        printSetup.SaveAs("ExportToPDF_"+setName)

        printManager.Apply()
        printManager.SubmitPrint()

        viewSheetSetting.Delete()

        tr.Commit()


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


def CreateDummySheet(snumber):
    #create a filter to get all the title block type
    FEC = FilteredElementCollector(doc)
    FEC.OfCategory(BuiltInCategory.OST_TitleBlocks)
    FEC.WhereElementIsElementType()

    #get elementid of first title block type
    titleblockid = FEC.FirstElementId()

    #Create sheet
    SHEET = None

    with Transaction(doc,"Create a dummy ViewSheet") as tr:
        tr.Start()
        SHEET = ViewSheet.Create(doc, titleblockid)
        SHEET.Name = "Dummy"
        SHEET.SheetNumber = "0000000"+str(snumber)
        tr.Commit()

    return SHEET

# Adds information about the page orientation to the sheet info list
def AddPageOrientationInfo(sheet_info_set):
    if int(sheet_info_set[2]) > int(sheet_info_set[3]):
        #Landscape
        sheet_info_set.append("1")
        return 1
    else:
        #Portrait
        sheet_info_set.append("0")
        return 0

def ReadRegValue(reg_key_name, reg_path):
    try:
        root_key = Registry.CurrentUser.OpenSubKey(reg_path, True)
        Pathname= root_key.GetValue(reg_key_name).ToString()
        root_key.Close()

        if (""==Pathname):
            raise WindowsError

        return Pathname
    except WindowsError:
        return [""]


def SetRegValue(reg_key_name, new_value, reg_path):
    try:
        Registrykey = Registry.CurrentUser.OpenSubKey(reg_path, True)

        value_type = None


        if new_value.isdigit() == False:
            value_type = RegistryValueKind.String
        elif new_value.isdigit() == True:
            value_type = RegistryValueKind.DWord
            new_value = int(new_value)

        Registrykey.SetValue(reg_key_name, new_value, value_type)
        Registrykey.Close()
        return True
    except WindowsError:
        return False


def AddCustomPaper(paper, width, height, id_num):
    custom_papers_path = r"Software\FinePrint Software\pdfFactory6\CustomPapers"
    paper_path = custom_papers_path +"\\"+paper

    custom_papers_reg = Registry.CurrentUser.OpenSubKey(custom_papers_path)
    if custom_papers_reg is not None:
        pass
    else:
        Registry.CurrentUser.CreateSubKey(custom_papers_path)

    paper_reg = Registry.CurrentUser.OpenSubKey(paper_path)
    if paper_reg is not None:
        pass
    else:
        Registry.CurrentUser.CreateSubKey(paper_path)
        paper_reg = Registry.CurrentUser.OpenSubKey(paper_path, True)
        paper_reg.SetValue("ID", id_num+211, RegistryValueKind.DWord)
        paper_reg.SetValue("Units", 2, RegistryValueKind.DWord)

    paper_reg = Registry.CurrentUser.OpenSubKey(paper_path, True)
    paper_reg.SetValue("Width", int(width)*1000, RegistryValueKind.DWord)
    paper_reg.SetValue("Height", int(height)*1000, RegistryValueKind.DWord)
    paper_reg.Close()


def convert_size(size_bytes):
    if size_bytes == 0:
        return "0B"

    size_name = ("B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
    i = int(math.floor(math.log(size_bytes, 1024)))
    p = math.pow(1024, i)
    s = round(size_bytes / p, 2)

    return (s, size_name[i])


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

        full_export_path = export_path #+"\\"+ file_name_folder
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

        full_export_path = _export_path #+"\\"+ file_name_folder
        if not os.path.exists(full_export_path):
            os.makedirs(full_export_path)

    return (working_directory, rvt_folder_name, full_export_path)

file_type = "PDF"

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
###########################################################################################

# find the sheets
family_instance_collector = FilteredElementCollector(doc)
family_instance_collector.OfCategory(BuiltInCategory.OST_TitleBlocks)
family_instance_collector.OfClass(FamilyInstance)

sheet_info_list = []

Output("\nInfo about the sheets: ")
for e in family_instance_collector:
    p = e.get_Parameter(BuiltInParameter.SHEET_NUMBER)
    sheet_number = p.AsString()

    p = e.get_Parameter(BuiltInParameter.SHEET_WIDTH)
    width = p.AsDouble()
    swidth = p.AsValueString()
    if "," in swidth:
        swidth = swidth.replace(",", ".")
        swidth = str(int(math.ceil(float(swidth))))
    else:
        swidth = str(int(math.ceil(float(swidth))))


    p = e.get_Parameter(BuiltInParameter.SHEET_HEIGHT)
    height = p.AsDouble()
    sheight = p.AsValueString()
    if "," in sheight:
        sheight = sheight.replace(",", ".")
        sheight = str(int(math.ceil(float(sheight))))
    else:
        sheight = str(int(math.ceil(float(sheight))))


    typeId = e.GetTypeId()
    type = doc.GetElement(typeId)
    sheet_type = Element.Name.GetValue(type)

    if sheet_type=="М_марка имя листа" or sheight=="0" or swidth=="0":
        continue

    if sheet_type=="А2К+50(по высоте)+100(по ширине)":
        sheet_type = "А2К_plus_50h_100w"

    if sheet_type=="А2К+100(по высоте)+100(по ширине)":
        sheet_type = "А2К_plus_100h_100w"

    sheet_info = []
    sheet_info.append(sheet_number)
    sheet_info.append(sheet_type)
    sheet_info.append(swidth)
    sheet_info.append(sheight)

    sheet_info_list.append(sheet_info)

    Output("   Sheet number: "+sheet_number+", Sizes: "+swidth+" x "+sheight+", type: "+sheet_type)
Output("*********************************************************************")

sheet_info_list = natsort.natsorted(sheet_info_list)
sheet_info_list = [list(g) for k, g in groupby(sheet_info_list, lambda x: x[1])]

# Insert Dummy Sheets info sets between the real sheet sets:
counter = 0
#length = len(sheet_info_list)*2 if len(sheet_info_list)%2 == 0 else len(sheet_info_list)+2
length = len(sheet_info_list)*2

for i in range(0, length, 2):
	sheet_number = "0000000"+str(int(i/2))
	sheet_type = sheet_info_list[int(i/2)+counter][0][1] + " (Dummy)"
	sheet_width = sheet_info_list[int(i/2)+counter][0][2]
	sheet_height =sheet_info_list[int(i/2)+counter][0][3]
	sheet_orientation = "1" if int(sheet_width) >= int(sheet_height) else "0"
	dummy_set = [[sheet_number, sheet_type, sheet_width, sheet_height, sheet_orientation]]
	sheet_info_list.insert(i, dummy_set)
	counter += 1
#######################################################################################################

# Delete all print settings in the document
document_print_settings = doc.GetPrintSettingIds()

if document_print_settings is not None:
    with Transaction(doc,"Delete previous Print Settings") as tr:
        tr.Start()
        doc.Delete(document_print_settings)
        tr.Commit()
#######################################################################################################
# Add a Schedule filter for the Dummy sheets so that they don't show up in the table
schedule_collector = FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()

possible_schedule_names = ( "Ведомость чертежей основного комплекта", "Ведомость рабочих чертежей основного комплекта", "Ведомость чертежей",
                            "Ведомость рабочих чертежей основного комплекта \"ЭЛ\"", "Ведомость основного комплекта чертежей",
                            "0Д_Ведомость чертежей", "01_АР_Ведомость чертежей", "Список листов", "ВЕДОМОСТЬ РАБОЧИХ ЧЕРТЕЖЕЙ ОСНОВНОГО КОМПЛЕКТА",
                            "ПД_ОВ_ВЕДОМОСТЬ РАБОЧИХ ЧЕРТЕЖЕЙ ОСНОВНОГО КОМПЛЕКТА" )

schedule_field = "Sheet Name"
filter_value = "Dummy"

Output("Schedules: ")
for s in schedule_collector:
    if any([x in s.Name for x in possible_schedule_names]) == True:
        Output("  "+s.Name)
        AddScheduleFilter(s, schedule_field, filter_value)
        Output("Filter for <Dummy> sheets Added")
Output("*********************************************************************")
#######################################################################################################
# Prepare to print all sheets in the document together
view_sheet_collector = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()

allSheets = []

for sheet in view_sheet_collector:
    allSheets.append(sheet)

allSheets = natsort.natsorted(allSheets, key=lambda vs: vs.SheetNumber)
# Add Dummies between the sheets:
length = len(allSheets)*2

for i in range(0, length, 2):
    allSheets.insert( i, CreateDummySheet(int(i/2)) )


# This means that sheets are to be combined into one PDF
SetRegValue("CollectJobs", "300", reg_path_pdfFactoryPro)

# Set the output filename to the name of the current rvt document
SetRegValue("OutputFilePerm", the_export_path+"\\"+file_name+"_DELETE"+".pdf", reg_path_pdfFactoryPro)

counter = 0
Output("Sheet sets to be printed: ")
for vset in sheet_info_list:
    vwSet = ViewSet()
    AddPageOrientationInfo(vset[0])
    # Check if the paper type and size exist in the printer's registry, and overwrite if necessary:
    '''
    if any(c.islower() for c in vset[0][1].split("Dummy")[0]):
        paper_type_str = vset[0][1]
        paper_type_str = paper_type_str.split("Dummy")
        paper_type_str[0] = paper_type_str[0].upper()
        paper_type_str = "".join(paper_type_str)

        vset[0][1] = paper_type_str
    '''
    if vset[0][1] == "А3а":
        vset[0][1] = "А3А"
    elif vset[0][1] == "А3а (Dummy)":
        vset[0][1] = "А3А (Dummy)"

    AddCustomPaper(vset[0][1], vset[0][2], vset[0][3], counter)

    # Modify Windows Registry Values to change the paper size:
    SetRegValue("PaperName", vset[0][1], reg_path_PrinterDriverData)
    SetRegValue("PaperWidth", vset[0][2], reg_path_PrinterDriverData)
    SetRegValue("PaperHeight", vset[0][3], reg_path_PrinterDriverData)
    SetRegValue("Orientation", vset[0][4], reg_path_PrinterDriverData)
    SetRegValue("MarginBottom", "1", reg_path_PrinterDriverData)
    SetRegValue("MarginTop", "1", reg_path_PrinterDriverData)
    SetRegValue("MarginLeft", "1", reg_path_PrinterDriverData)
    SetRegValue("MarginRight", "1", reg_path_PrinterDriverData)

    Output(" Set [" +vset[0][1]+ "]: ")
    for vsheet in vset:
        Output("\t"+vsheet[0])
        for vs in allSheets:
            if vs.SheetNumber == vsheet[0] and vs.CanBePrinted == True:
                vwSet.Insert(vs)
    ExportPDF(doc, printer_name, vwSet, vset[0][2], vset[0][3], vset[0][1], str(counter))
    counter += 1

    Output("-----------------------------------------------------------")

# Set CollectJobs to -1 so that the printer will combine all PDFs into one
reg = Registry.CurrentUser.OpenSubKey(reg_path_pdfFactoryPro, True)
reg.SetValue("CollectJobs", -1, RegistryValueKind.DWord)
reg.Close()

# This print job will be ignored and printed PDF sheets will be combined into One
vwSet_dummy0 = ViewSet()
vwSet_dummy0.Insert(allSheets[0])
vset_dummy0 = sheet_info_list[0]
ExportPDF(doc, printer_name, vwSet_dummy0, vset_dummy0[0][2], vset_dummy0[0][3], vset_dummy0[0][1], str(counter))

# Set the output filename to whatever (it is not supposed to be printed anyway)
SetRegValue("OutputFilePerm", the_export_path+"\\"+"Dummy"+".pdf", reg_path_pdfFactoryPro)

Output("*********************************************************************")
######################################################################################################################
# Reset the paper size to default (Dummy):
AddCustomPaper("Dummy", "3000", "3000", 19)
SetRegValue("PaperName", "Dummy", reg_path_PrinterDriverData)
SetRegValue("PaperWidth", "3000", reg_path_PrinterDriverData)
SetRegValue("PaperHeight", "3000", reg_path_PrinterDriverData)
SetRegValue("Orientation", "1", reg_path_PrinterDriverData)
######################################################################################################################
# Wait the calculated amount of time
multiplier = 2
if "2019_TSE_CAPITAL_PARK_3E" in working_directory and rvt_folder_name == "KG":
    if file_name == "TS_CP_3E_KG_P1_PRM" or file_name == "TS_CP_3E_KG_P1":
        multiplier = 30

wait_time = int(convert_size(os.path.getsize(revitFilePath))[0]*multiplier)

Output("\nWaiting " +str(wait_time)+ " seconds so that " +printer_name+ " would finish its printjob...")
time.sleep(wait_time)
Output(" Now the dummy sheets need to be removed!")
