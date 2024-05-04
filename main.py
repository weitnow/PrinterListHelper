import os.path

from Outputmanager import Outputmanager
from Printermanager import Printermanager
from copy import deepcopy

###############################################################################
#                                     PARAMETER                               #
###############################################################################
LOAD_PRINTER_LIST_SUPPORT = True
LOAD_PRINTER_LIST_INSPECT = True

LOAD_PRINTER_LIST_DEP_AAU = True
LOAD_PRINTER_LIST_DEP_ADM = True
LOAD_PRINTER_LIST_DEP_DIS = True
LOAD_PRINTER_LIST_DEP_DIV = True
LOAD_PRINTER_LIST_DEP_FIN = True
LOAD_PRINTER_LIST_DEP_TEC = True
LOAD_PRINTER_LIST_DEP_ZUL = True
LOAD_PRINTER_LIST_DEP_SCH = True

LOAD_USER_TO_WINDOWSPRINTER_FROM_LIST_DEPS = True
LOAD_PC_TO_DEFAULT_WINDOWSPRINTER_FROM_LIST_DEPS = True

ADD_WINDOWSUSER_FROM_WCPS_TO_PRINTERMANAGER_PRINTERS = True

GENERATE_STATISTICS = True

################################
SAVE_PICKLE = False
SAVE_PICKLE_NAME = "09082023"
LOAD_PICKLE = False
LOAD_PICKLE_NAME = "09082023"
################################

# Komplette Druckerliste, Alle Drucker, Erfassung in CARi und auf Druckerserver
GENERATE_OUTPUT_EXCELFILE_ROBOT_CARI_AND_PRINTERSERVER_ALL_PRINTER = True
# -> robot_cari_printerserver_all_printers

# Alle Drucker, Erfassung auf Druckerserver
GENERATE_OUTPUT_EXCELFILE_ROBOT_PRINTERSERVER_ALL_PRINTER = True
# -> robot_printerserver_all_printers

# Nur CARi-Drucker, Erfassung in CARi
GENERATE_OUTPUT_EXCELFILE_ROBOT_CARI_ONLY_IF_CARI_RELEVANT_PRINTER = True
# -> robot_cari_only_cari_relevant

# Nur CARi-Drucker, Erfassung auf Druckerserver
GENERATE_OUTPUT_EXCELFILE_ROBOT_PRINTERSERVER_ONLY_IF_CARI_RELEVANT_PRINTER = True
# -> robot_printerserver_only_cari_relevant

GENERATE_OUTPUT_EXCELFILE_SERDAR = True
GENERATE_OUTPUT_EXCELFILE_GILLES = True

SET_TWOSIDED_FROM_NONE_TO_TRUE = True
SET_INSPECT_FROM_NONE_TO_FALSE = True
SET_ACTIVE_FROM_NONE_TO_1 = True

PRINTER_LIST_SUPPORT = "input/all_printers_list_of_support/Druckerliste_Support.xlsx"
PRINTER_LIST_INSPECT = "input/printers_inspect/Druckerliste_Inspect_20062023.xlsx"
PRINTER_LIST_AAU = "input/printers_by_department/AAU.xlsx"
PRINTER_LIST_ADM = "input/printers_by_department/ADM.xlsx"
PRINTER_LIST_DIS = "input/printers_by_department/DIS.xlsx"
PRINTER_LIST_DIV = "input/printers_by_department/DIV.xlsx"
PRINTER_LIST_FIN = "input/printers_by_department/FIN.xlsx"
PRINTER_LIST_TEC = "input/printers_by_department/TEC.xlsx"
PRINTER_LIST_ZUL = "input/printers_by_department/ZUL.xlsx"
PRINTER_LIST_SCH = "input/printers_by_department/SCH.xlsx"

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
    printermanager.load_printerlist_of_department(PRINTER_LIST_AAU, "aau", LOAD_USER_TO_WINDOWSPRINTER_FROM_LIST_DEPS, LOAD_PC_TO_DEFAULT_WINDOWSPRINTER_FROM_LIST_DEPS)

if LOAD_PRINTER_LIST_DEP_ADM:
    #load the printerlist of department ADM
    printermanager.load_printerlist_of_department(PRINTER_LIST_ADM, "adm", LOAD_USER_TO_WINDOWSPRINTER_FROM_LIST_DEPS, LOAD_PC_TO_DEFAULT_WINDOWSPRINTER_FROM_LIST_DEPS)

if LOAD_PRINTER_LIST_DEP_DIS:
    #load the printerlist of department DIS
    printermanager.load_printerlist_of_department(PRINTER_LIST_DIS, "dis", LOAD_USER_TO_WINDOWSPRINTER_FROM_LIST_DEPS, LOAD_PC_TO_DEFAULT_WINDOWSPRINTER_FROM_LIST_DEPS)

if LOAD_PRINTER_LIST_DEP_DIV:
    #load the printerlist of department DIVers
    printermanager.load_printerlist_of_department(PRINTER_LIST_DIV, "div", LOAD_USER_TO_WINDOWSPRINTER_FROM_LIST_DEPS, LOAD_PC_TO_DEFAULT_WINDOWSPRINTER_FROM_LIST_DEPS)

if LOAD_PRINTER_LIST_DEP_FIN:
    #load the printerlist of department FIN
    printermanager.load_printerlist_of_department(PRINTER_LIST_FIN, "fin", LOAD_USER_TO_WINDOWSPRINTER_FROM_LIST_DEPS, LOAD_PC_TO_DEFAULT_WINDOWSPRINTER_FROM_LIST_DEPS)

if LOAD_PRINTER_LIST_DEP_TEC:
    #load the printerlist of department TEC
    printermanager.load_printerlist_of_department(PRINTER_LIST_TEC, "tec", LOAD_USER_TO_WINDOWSPRINTER_FROM_LIST_DEPS, LOAD_PC_TO_DEFAULT_WINDOWSPRINTER_FROM_LIST_DEPS)

if LOAD_PRINTER_LIST_DEP_ZUL:
    #load the printerlist of department ZUL
    printermanager.load_printerlist_of_department(PRINTER_LIST_ZUL, "zul", LOAD_USER_TO_WINDOWSPRINTER_FROM_LIST_DEPS, LOAD_PC_TO_DEFAULT_WINDOWSPRINTER_FROM_LIST_DEPS)

if LOAD_PRINTER_LIST_DEP_SCH:
    #load the printerlist of department ZUL
    printermanager.load_printerlist_of_department(PRINTER_LIST_SCH, "sch", LOAD_USER_TO_WINDOWSPRINTER_FROM_LIST_DEPS, LOAD_PC_TO_DEFAULT_WINDOWSPRINTER_FROM_LIST_DEPS)

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

# add cari-windows-users  to the each printer. if for example printer pstva1111-s1 has b126kec and pstva1111-s2 has b126swm then both users are added to pstva1111. the same is true
# for computer !!!! WARNING, this method can only be used AFTER add_wcps_to_matching_papersource_of_printers() has been executed because it takes the data from there!
if ADD_WINDOWSUSER_FROM_WCPS_TO_PRINTERMANAGER_PRINTERS:
    printermanager.add_cari_users_to_printer()

###############################################################################
#                                     OUTPUT                                  #
###############################################################################
outputmanager = Outputmanager(printermanager)

### PICKLE ###
if SAVE_PICKLE:
    outputmanager.pickle_save(printermanager.printers, "printermanager-printers", SAVE_PICKLE_NAME)

if LOAD_PICKLE:
    loaded_printers_from_picklefile = outputmanager.pickle_load("printermanager-printers", LOAD_PICKLE_NAME)
    printermanager.printers_from_loaded_pickle = deepcopy(loaded_printers_from_picklefile)

    #compare the current printerlist from printermanger.printers with the loaded printermanager.printers from pickle
    #if the printer is identical with the printer in the other list (true, [printer]) is returned. if it is not identical then (false, [printer1, printer2, printerdelta])is returned
    loaded_printers_from_picklefile = printermanager.compare_two_printerlists_and_return_difference(loaded_printers_from_picklefile, printermanager.printers)
    for item in loaded_printers_from_picklefile:
        if item[0] == False:
            #if printer is not identical we get (False, [printer1, printer2, printerdelta])
            printermanager.printers_which_have_changed_compared_with_loaded_pickle.append(item[1][1])
            printermanager.printers_which_have_changed_compared_with_loaded_pickle_only_changed_attributes.append(item[1][2])

list_with_header_names = ["Standort", "Arbeitsplatz (Büro)", "Printserver Link", "Druckername", "Schacht Name", "Druckername Printerserver", "IP Drucker (nur zur Information für Vergleich QIP)", "Portname", "Formular CARI / Prüfbahn / Parkplatz", "Format des Formulars", "2-sided", "Inspect", "Zuständige Fachabteilung", "Aktiv", "Drucker Modell", "Drucker Treiber", "Bemerkung"]

###############################################################
path_with_filename = "output/robot_cari_printerserver_all_printers"
title_of_worksheet = "CARi Druckerwarteschlangen"

if GENERATE_OUTPUT_EXCELFILE_ROBOT_CARI_AND_PRINTERSERVER_ALL_PRINTER:
    printerlist = outputmanager.return_deep_copy_of_printermanger_printers()
    outputmanager.create_output_excel_list_for_robot(path_with_filename=path_with_filename, title_of_worksheet=title_of_worksheet, list_with_header_names=list_with_header_names, printer_list=printerlist)
###############################################################

path_with_filename = "output/robot_cari_only_cari_relevant"
title_of_worksheet = "CARi Druckerwarteschlangen"

if GENERATE_OUTPUT_EXCELFILE_ROBOT_CARI_ONLY_IF_CARI_RELEVANT_PRINTER:
    printerlist = outputmanager.return_deep_copy_of_printermanger_printers()
    printerlist = outputmanager.delete_printer_without_wcps(printerlist)
    outputmanager.create_output_excel_list_for_robot(path_with_filename=path_with_filename, title_of_worksheet=title_of_worksheet, list_with_header_names=list_with_header_names, printer_list=printerlist)
###############################################################

path_with_filename = "output/robot_printerserver_all_printers"
title_of_worksheet = "CARi Druckerwarteschlangen"

if GENERATE_OUTPUT_EXCELFILE_ROBOT_PRINTERSERVER_ALL_PRINTER:
    printerlist = outputmanager.return_deep_copy_of_printermanger_printers()
    printerlist = outputmanager.delete_all_wcps(printerlist)
    printerlist = outputmanager.delete_papersource_with_inspect(printerlist)
    outputmanager.create_output_excel_list_for_robot(path_with_filename=path_with_filename, title_of_worksheet=title_of_worksheet, list_with_header_names=list_with_header_names, printer_list=printerlist)
###############################################################

path_with_filename = "output/robot_printerserver_only_cari_relevant"
title_of_worksheet = "CARi Druckerwarteschlangen"
if GENERATE_OUTPUT_EXCELFILE_ROBOT_PRINTERSERVER_ONLY_IF_CARI_RELEVANT_PRINTER:
    printerlist = outputmanager.return_deep_copy_of_printermanger_printers()
    printerlist = outputmanager.delete_printer_without_wcps(printerlist)
    printerlist = outputmanager.delete_all_wcps(printerlist)
    printerlist = outputmanager.delete_papersource_with_inspect(printerlist)
    outputmanager.create_output_excel_list_for_robot(path_with_filename=path_with_filename, title_of_worksheet=title_of_worksheet, list_with_header_names=list_with_header_names, printer_list=printerlist)
###############################################################

path_with_filename = "output/printerlist_serdar"
title_of_worksheet = "printers"
list_with_header_names = ["printername", "paperslots", "users", "users_to_windowsprinter", "users_combined", "pcs_to_default_windowsprinter"]

if GENERATE_OUTPUT_EXCELFILE_SERDAR:
    if LOAD_PICKLE:

       outputmanager.create_output_excel_list_for_serdar_new_version(path_with_filename=path_with_filename,
                                                                     title_of_worksheet="PRINTERS",
                                                                     list_with_header_names=list_with_header_names,
                                                                     printer_list=outputmanager.return_deep_copy_of_printermanger_printers(),
                                                                     delete_previous_file=True)
       outputmanager.create_output_excel_list_for_serdar_new_version(path_with_filename=path_with_filename,
                                                                     title_of_worksheet=f"PICKLE_{LOAD_PICKLE_NAME}",
                                                                     list_with_header_names=list_with_header_names,
                                                                     printer_list=printermanager.printers_from_loaded_pickle,
                                                                     delete_previous_file=False)
       outputmanager.create_output_excel_list_for_serdar_new_version(path_with_filename=path_with_filename,
                                                                     title_of_worksheet="ONLY_CHANGED_PRINTERS",
                                                                     list_with_header_names=list_with_header_names,
                                                                     printer_list=printermanager.printers_which_have_changed_compared_with_loaded_pickle,
                                                                     delete_previous_file=False)
       outputmanager.create_output_excel_list_for_serdar_new_version(path_with_filename=path_with_filename,
                                                                     title_of_worksheet="PRINTERS_PICKLE_DELTA",
                                                                     list_with_header_names=list_with_header_names,
                                                                     printer_list=printermanager.printers_which_have_changed_compared_with_loaded_pickle_only_changed_attributes,
                                                                     delete_previous_file=False)

    else:
        printerlist = outputmanager.return_deep_copy_of_printermanger_printers()
        list_of_dicts = []
        for printer in printerlist:
            #iterate through a copy of printermanger.printers and call for each printer a method which gets back a dict like {printername = pstva1769, paperslots = [s1, s2, s3], workspace = [AL-ZUL-PEZ1, AL_ZUL-PEZ2...], users = [B126SMP, B126IMD...]}
            list_of_dicts.append(printer.get_users_paperslots_workspaces(printermanager))

        outputmanager.create_output_excel_list_for_serdar(path_with_filename=path_with_filename, title_of_worksheet=title_of_worksheet, list_with_header_names=list_with_header_names, printer_list=list_of_dicts)

###############################################################

path_with_filename = "output/statistic/statistic.xlsx"
title_of_worksheet = "statistic"
list_with_header_names = ["username", "number_cari_workspaces", "number_printers", "number_printerslots"]

if GENERATE_STATISTICS:
    printerlist = outputmanager.return_deep_copy_of_printermanger_printers()
    list_of_dicts = []
    for printer in printerlist:
        # iterate through a copy of printermanger.printers and call for each printer a method which gets back a dict like {printername = pstva1769, paperslots = [s1, s2, s3], workspace = [AL-ZUL-PEZ1, AL_ZUL-PEZ2...], users = [B126SMP, B126IMD...]}
        list_of_dicts.append(printer.get_users_paperslots_workspaces(printermanager))

    outputmanager.create_output_excel_statistic(path_with_filename=path_with_filename,
                                                      title_of_worksheet=title_of_worksheet,
                                                      list_with_header_names=list_with_header_names,
                                                      printer_list=list_of_dicts,
                                                        delete_previous_file=True,
                                                       workspace_list=printermanager.workspaces)


###############################################################

path_with_filename = "output/bureau_lieugestion_list_gilles"
title_of_worksheet = None       #this will be defined in the code block below
list_with_header_names = None   #this will be defined in the code block below

if GENERATE_OUTPUT_EXCELFILE_GILLES:
    # check if the file exists and if it does delete it. the file will newly created the first time when outputmanager.create_output_excel_list_for_gilles is run
    if os.path.exists(f'{path_with_filename}.xlsx'):
        os.remove(f'{path_with_filename}.xlsx')
    #create worksheet bureau
    title_of_worksheet = "Bureau"
    list_with_header_names = ["root.Profiles.Bureau-id", "root.Profiles.Bureau-libelle", "root.Profiles.Bureau-value"]
    outputmanager.create_output_excel_list_for_gilles(path_with_filename=path_with_filename, title_of_worksheet=title_of_worksheet, list_with_header_names=list_with_header_names, printermanager=printermanager)
    #create worksheet lieugestion
    title_of_worksheet = "LieuGestion"
    list_with_header_names = ["root.Profiles.LieuGestion-id", "root.Profiles.LieuGestion-libelle", "root.Profiles.LieuGestion-value"]
    outputmanager.create_output_excel_list_for_gilles(path_with_filename=path_with_filename,
                                                      title_of_worksheet=title_of_worksheet,
                                                      list_with_header_names=list_with_header_names,
                                                      printermanager=printermanager)
    #create worksheet lienbureaugestion
    title_of_worksheet = "LienBureauLieuGestion"
    list_with_header_names = ["root.Profiles.LienBuerauLieuGestion-bureau", "root.Profiles.LienBuerauLieuGestion-lieuGestion"]
    outputmanager.create_output_excel_list_for_gilles(path_with_filename=path_with_filename,
                                                      title_of_worksheet=title_of_worksheet,
                                                      list_with_header_names=list_with_header_names,
                                                      printermanager=printermanager)
    #create worksheet bureauusers
    title_of_worksheet = "BureauUsers"
    list_with_header_names = ["root.Profiles.Bureau-id",
                              "root.Profiles.Bureau-libelle",
                              "users"]
    outputmanager.create_output_excel_list_for_gilles(path_with_filename=path_with_filename,
                                                      title_of_worksheet=title_of_worksheet,
                                                      list_with_header_names=list_with_header_names,
                                                      printermanager=printermanager)




###############################################################################
#                                     OUTPUT-DEBUG                            #
###############################################################################

# prints the last workspace-id - in case someone wants to add a new workspace, he knows with wich id to continue in the excelsheet

# Get the workspace with the highest id
workspace_with_highest_id = max(printermanager.workspaces, key=lambda x: x.id)
print(f"Last used ID is: {workspace_with_highest_id} ")
print(f"Next ID for new Workspace is: {len(printermanager.workspaces)+1}")


print()
