from Outputmanager import Outputmanager
from Printermanager import Printermanager

###############################################################################
#                                     PARAMETER                               #
###############################################################################
LOAD_PRINTER_LIST_SUPPORT = True
LOAD_PRINTER_LIST_INSPECT = True
################################
LOAD_PRINTER_LIST_DEP_AAU = True
LOAD_PRINTER_LIST_DEP_ADM = True
LOAD_PRINTER_LIST_DEP_DIS = True
LOAD_PRINTER_LIST_DEP_FIN = False
LOAD_PRINTER_LIST_DEP_TEC = True
LOAD_PRINTER_LIST_DEP_ZUL = True

################################
PRINT_TO_CONSOLE = False

GENERATE_OUTPUT_EXCELFILE_ROBOT_CARI_AND_PRINTERSERVER_ALL_PRINTER = True
GENERATE_OUTPUT_EXCELFILE_ROBOT_PRINTERSERVER_ALL_PRINTER = True
GENERATE_OUTPUT_EXCELFILE_ROBOT_CARI_ONLY_IF_CARI_RELEVANT_PRINTER = True
GENERATE_OUTPUT_EXCELFILE_ROBOT_PRINTERSERVER_ONLY_IF_CARI_RELEVANT_PRINTER = True
GENERATE_OUTPUT_EXCELFILE_SERDAR = True

SET_TWOSIDED_FROM_NONE_TO_TRUE = True
SET_INSPECT_FROM_NONE_TO_FALSE = True
SET_ACTIVE_FROM_NONE_TO_1 = True

PRINTER_LIST_SUPPORT = "input/all_printers_list_of_support/Druckerliste_Support_20062023.xlsx"
PRINTER_LIST_INSPECT = "input/printers_inspect/Druckerliste_Inspect_20062023.xlsx"
PRINTER_LIST_AAU = "input/printers_by_department/AAU.xlsx"
PRINTER_LIST_ADM = "input/printers_by_department/ADM.xlsx"
PRINTER_LIST_DIS = "input/printers_by_department/DIS.xlsx"
PRINTER_LIST_FIN = "input/printers_by_department/FIN.xlsx"
PRINTER_LIST_TEC = "input/printers_by_department/TEC.xlsx"
PRINTER_LIST_ZUL = "input/printers_by_department/ZUL.xlsx"

###############################################################################
#                                     ROBOT-OPTIONS                           #
###############################################################################
# 1     Roboter macht nichts
# 2     Nur in CARi erfassen
# 3     Nur auf Druckerserver erfassen
# 4     In CARi und Druckerserver erfassen

###############################################################################
#                                     INPUT LIST OF SUPPORT                   #
###############################################################################
#initialize printermanager without any printer
printermanager = Printermanager()

if LOAD_PRINTER_LIST_SUPPORT:
    #load the printerlist of support into printermanger.printers
    printermanager.load_printerlist_of_support(PRINTER_LIST_SUPPORT)

if LOAD_PRINTER_LIST_INSPECT:
    printermanager.load_printerlist_inspect(PRINTER_LIST_INSPECT)

if LOAD_PRINTER_LIST_DEP_AAU:
    #load the printerlist of department AAU
    printermanager.load_printerlist_of_department(PRINTER_LIST_AAU, "aau")

if LOAD_PRINTER_LIST_DEP_ADM:
    #load the printerlist of department ADM
    printermanager.load_printerlist_of_department(PRINTER_LIST_ADM, "adm")

if LOAD_PRINTER_LIST_DEP_DIS:
    #load the printerlist of department DIS
    printermanager.load_printerlist_of_department(PRINTER_LIST_DIS, "dis")

if LOAD_PRINTER_LIST_DEP_FIN:
    #load the printerlist of department FIN
    printermanager.load_printerlist_of_department(PRINTER_LIST_FIN, "fin")

if LOAD_PRINTER_LIST_DEP_TEC:
    #load the printerlist of department TEC
    printermanager.load_printerlist_of_department(PRINTER_LIST_TEC, "tec")

if LOAD_PRINTER_LIST_DEP_ZUL:
    #load the printerlist of department ZUL
    printermanager.load_printerlist_of_department(PRINTER_LIST_ZUL, "zul")
print("---------------------------------------------------------------------")

###############################################################################
#                                     PREPARE DATA                            #
###############################################################################

if SET_TWOSIDED_FROM_NONE_TO_TRUE:
    #set the attribute twosided for all printers from None to True. If it has another value then None nothing is changed
    for printer in printermanager.printers:
        printer.set_twosided_from_None_to_True()

if SET_INSPECT_FROM_NONE_TO_FALSE:
    #set the attribute inspect for all printers from None to False. If it has another value then None nothing is changed
    for printer in printermanager.printers:
        printer.set_inspect_from_None_to_False()

if SET_ACTIVE_FROM_NONE_TO_1:
    #set the attribute active for all printers from None to 1. If it has another value then None nothing is changed
    for printer in printermanager.printers:
        printer.set_active_from_None_to_1()

# add wcps to the matching papersources of the printerlist in printermanager
printermanager.add_wcps_to_matching_papersource_of_printers()

###############################################################################
#                                     PLAYGROUND                              #
###############################################################################



###############################################################################
#                                     PRINT TO CONSOLE                        #
###############################################################################

if PRINT_TO_CONSOLE:
    for printer in printermanager.printers:
        print(printer)

###############################################################################
#                                     OUTPUT                                  #
###############################################################################

# WARNING, every change to printermanager.printers after instantiating outputmanager will not be transfered to output
# The reason for that is, that outputmanager makes an independent copy of the list printermanager.printers after being instantiated.

outputmanager = Outputmanager(printermanager)
list_with_header_names = ["Standort", "Arbeitsplatz (Büro)", "Printserver Link", "Druckername", "Schacht Name", "Druckername Printerserver", "IP Drucker (nur zur Information für Vergleich QIP)", "Portname", "Formular CARI / Prüfbahn / Parkplatz", "Format des Formulars", "2-sided", "Inspect", "Zuständige Fachabteilung", "Aktiv", "Drucker Modell", "Drucker Treiber", "Bemerkung"]

###############################################################
path_with_filename = "output/robot_cari_printerserver_all_printers"
title_of_worksheet = "kectest"

if GENERATE_OUTPUT_EXCELFILE_ROBOT_CARI_AND_PRINTERSERVER_ALL_PRINTER:
    printerlist = outputmanager.return_deep_copy_of_printermanger_printers()
    outputmanager.create_output_excel_list_for_robot(path_with_filename=path_with_filename, title_of_worksheet=title_of_worksheet, list_with_header_names=list_with_header_names, printer_list=printerlist)
###############################################################

path_with_filename = "output/robot_cari_only_cari_relevant"
title_of_worksheet = "kectest"

if GENERATE_OUTPUT_EXCELFILE_ROBOT_CARI_ONLY_IF_CARI_RELEVANT_PRINTER:
    printerlist = outputmanager.return_deep_copy_of_printermanger_printers()
    printerlist = outputmanager.delete_printer_without_wcps(printerlist)
    outputmanager.create_output_excel_list_for_robot(path_with_filename=path_with_filename, title_of_worksheet=title_of_worksheet, list_with_header_names=list_with_header_names, printer_list=printerlist)
###############################################################

path_with_filename = "output/robot_printerserver_all_printers"
title_of_worksheet = "kectest"

if GENERATE_OUTPUT_EXCELFILE_ROBOT_PRINTERSERVER_ALL_PRINTER:
    printerlist = outputmanager.return_deep_copy_of_printermanger_printers()
    printerlist = outputmanager.delete_all_wcps(printerlist)
    printerlist = outputmanager.delete_papersource_with_inspect(printerlist)
    outputmanager.create_output_excel_list_for_robot(path_with_filename=path_with_filename, title_of_worksheet=title_of_worksheet, list_with_header_names=list_with_header_names, printer_list=printerlist)
###############################################################

path_with_filename = "output/robot_printerserver_only_cari_relevant"
title_of_worksheet = "kectest"
if GENERATE_OUTPUT_EXCELFILE_ROBOT_PRINTERSERVER_ONLY_IF_CARI_RELEVANT_PRINTER:
    printerlist = outputmanager.return_deep_copy_of_printermanger_printers()
    printerlist = outputmanager.delete_printer_without_wcps(printerlist)
    printerlist = outputmanager.delete_all_wcps(printerlist)
    printerlist = outputmanager.delete_papersource_with_inspect(printerlist)
    outputmanager.create_output_excel_list_for_robot(path_with_filename=path_with_filename, title_of_worksheet=title_of_worksheet, list_with_header_names=list_with_header_names, printer_list=printerlist)
###############################################################

path_with_filename = "output/serdar"
title_of_worksheet = "kectest"
list_with_header_names = ["printername", "paperslots", "users"]

if GENERATE_OUTPUT_EXCELFILE_SERDAR:
    printerlist = outputmanager.return_deep_copy_of_printermanger_printers()
    list_of_dicts = []
    for printer in printerlist:
        list_of_dicts.append(printer.get_users_paperslots_workspaces(printermanager))
    outputmanager.create_output_excel_list_for_serdar(path_with_filename=path_with_filename, title_of_worksheet=title_of_worksheet, list_with_header_names=list_with_header_names, printer_list=list_of_dicts)

###############################################################################
#                                     OUTPUT-DEBUG                            #
###############################################################################