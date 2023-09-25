import pandas as pd
import os
from Printer import Printer
from Workspace import Workspace
import logging
from copy import deepcopy

class Printermanager:

    # the implemented __new__ methode makes sure this class can only be used as a singleton
    def __new__(cls):
        if not hasattr(cls, 'instance'):
            cls.instance = super(Printermanager, cls).__new__(cls)
        return cls.instance

    def __init__(self):
        # if old log file exists delete it
        if os.path.exists("log/logfile.log"):
            os.remove("log/logfile.log")
        # create log file
        logging.basicConfig(filename="log/logfile.log", level=logging.INFO)

        self.printers = []          #contains all printers with printerslots only once. for example pstva1139-s1 cannot exist more then once.
        self.workspaces = []        #contains all workspaces only once. for example AL-ZUL-FZZSpez1 cannot exist more then once.

        self.printers_from_loaded_pickle = None #contains all printers with printerslots only once. for example pstva1139-s1 cannot exist more then once.
        self.printers_which_have_changed_compared_with_loaded_pickle = [] #contains all printers which have a changed attributes if compared to the loaded_pickle_file
        self.printers_which_have_changed_compared_with_loaded_pickle_only_changed_attributes = [] #contains delta/difference between current self.printers and loaded printers from loaded pickle file


    def get_printer(self, printername: str):
        """returns the printer with the provided printername"""
        for printer in self.printers:
            if printername == printer.printername:
                return printer
        raise Exception(f"printermanger.get_printer() -> Printer {printername} not found in printermanger.printers (which means it the printer is most likely not defined in the printerlist of support")

    def get_printers(self, *, standort=None, buero=None, printername=None, ip=None, model=None, ):
        """returns all printers as a list
        you can also add filter arguments like this get_printers(standort = "AL", model = "Brother HL-L6250DN")"""
        filtered_printer_list = []
        filter_dict = dict(standort=standort, buero=buero, printername=printername, ip=ip, model=model)
        for printer in self.printers:
            match = True
            for key, value in filter_dict.items():
                if value is not None and (not hasattr(printer, key) or getattr(printer, key) != value):
                    match = False
                    break
            if match:
                filtered_printer_list.append(printer)
        return filtered_printer_list

    def add_wcps_to_matching_papersource_of_printers(self):
        """adds the wcps in printermanger.workspaces to the matching papersource of printers. it adds a tuple like (printername, printerslot, workspace_name, cariform, department)"""
        for workspace in self.workspaces:
            for wcps in workspace.wcps:
                printername = wcps.printername
                printerslot = wcps.printerslot
                workspace_name = wcps.workspace_name
                cariform = wcps.cariform
                department = wcps.department
                workspaceid = wcps.workspace_id
                workspace_userlist = wcps.workspace_users

                for printer in self.printers:
                    if printer.printername == printername:
                        printer.add_wcps(printername, printerslot, workspace_name, cariform, department, workspaceid, workspace_userlist)

    def load_printerlist_of_support(self, path_to_excel_file: str):
        """load the Excel file of support with all printers"""
        number_of_printers_add = 0
        check_list_if_printer_name_is_defined_more_then_once = set()
        check_list_if_same_ip_is_defined_more_then_once = set()

        excellist = pd.read_excel(path_to_excel_file)
        for index, row in excellist.iterrows():
            if isinstance(row['Standort'], str):
                printer = Printer(row['Standort'], row['Büro'], row['Druckername'], row['IP Adresse'], row['M0'],
                                  row['S1'], row['S2'], row['S3'], row['S4'], row['S5'], row['Drucker Modell'])

                #check if there is a space in Druckername or IP-Adresse
                if ' ' in row['Druckername'] or ' ' in row['IP Adresse']:
                    logging.warning(f"invalid char space detected in {row['Druckername']} or in {row['IP Adresse']} check {path_to_excel_file}")

                #check if printername has already been added and is defined more then once in the excellist
                if row['Druckername'] in check_list_if_printer_name_is_defined_more_then_once:
                    logging.warning(f"Warning, {row['Druckername']} is defined more then once in {path_to_excel_file}")
                check_list_if_printer_name_is_defined_more_then_once.add(row['Druckername'])

                # check if same ip has already been added and is defined more then once in the excellist
                if row['IP Adresse'] in check_list_if_same_ip_is_defined_more_then_once:
                    logging.warning(f"Warning, IP {row['IP Adresse']} is defined more then once in {path_to_excel_file}")
                check_list_if_same_ip_is_defined_more_then_once.add(row['IP Adresse'])

                self.printers.append(printer)
                number_of_printers_add += 1

        logging.info(f"{number_of_printers_add} printers from {path_to_excel_file} added")

    def load_printerlist_inspect(self, path_to_excel_file: str):
        """load the excel file printer inspect, checks if the printername and slot exists in printermanger.printers and updates the paperslots and wcps with cairform, department, inspect and active = 1 (which means
        for the roboter to do nothing. If the printer doesnt exist in printermanger.printers a error is thrown"""
        excellist = pd.read_excel(path_to_excel_file)

        error_set = set()
        printers_added = set()
        ignore_list = set()
        for index, row in excellist.iterrows():
            if isinstance(row['Druckername'], str):
                printername = row['Druckername']
                printerslot = row['Schacht Name']
                caridoc = row['Formular CARI / Prüfbahn / Parkplatz']
                paperformat = row['Format des Formulars']
                twosided = False
                department = row['Zuständige Fachabteilung']
                inspect = row['Inspect']
                active = row['Aktiv']
                name_for_ignorelist = f"{printername}-{printerslot}"

                try:
                    printer_from_selfprinters = self.get_printer(printername)
                    # checking if the printer has this printerslot and if the printerslot has the same format in the printerlist of support
                    printer_identical = printer_from_selfprinters.check_if_papersource_exists(printerslot=printerslot,
                                                                                              paperformat=paperformat)
                    if printer_identical:
                        #if printer is found in list of support with identical printerslot and paperformat update it
                        #but first check if twosied, inspect and active are none
                        printers_added.add(printername)
                        for papersource in printer_from_selfprinters.papersources:
                            if papersource.printerslot == printerslot:
                                checklist = [papersource.check_if_inspect_is_none(), papersource.check_if_twosided_is_none(), papersource.check_if_active_is_none()]
                                #if inspect, twosied and active is none, then if all(checklist) returns true and we can update the papersource with the values of inpsect excel list
                                if all(checklist):
                                    if inspect == "x":
                                        papersource.inspect = True
                                    if inspect == "CUT":
                                        papersource.inspect = "CUT"
                                    if twosided == False:
                                        papersource.twosided = False
                                    elif twosided == "2-sided":
                                        papersource.twosided = True
                                    else:
                                        papersource.twosided = None

                                    papersource.active = active
                                    ignore_list.add(name_for_ignorelist)

                                else:
                                    if name_for_ignorelist not in ignore_list:
                                        logging.warning(f"{printername}, {printerslot} has already set values in twosided, inspect or active. therefore values from {path_to_excel_file} cannot be set")
                                        ignore_list.add(name_for_ignorelist)
                    else:
                        error_set.add(f"{printername} has {printerslot}:{paperformat} in {path_to_excel_file} but different value in printerlist of support")
                except:
                    error_set.add(f"{printername} is defined in {path_to_excel_file} but has not been found in the printerlist of support")
        for item in error_set:
            logging.warning(item)

        logging.info(f"{len(printers_added)} printers has been updated from {path_to_excel_file}")

    def _load_workspaces_from_printerlist_of_department(self, path_to_excel_file: str):
        ######################################################################################
        # Load workspace-names and save it to self.workspaces - sheet Arbeitsplatz
        numbers_of_added_workspaces = 0
        excellist = pd.read_excel(path_to_excel_file, sheet_name='Arbeitsplatz')
        for index, row in excellist.iterrows():
            if isinstance(row['libelle'], str):
                assert row['Standort'] in (
                    'Albisgüetli', 'Winterthur', 'Regensdorf', 'Hinwil', 'Oberrieden', 'Bülach', 'Bassersdorf',
                    'AMA'), f"Invalid location in Excel {path_to_excel_file} in row ID {row['id']}"
                assert row['Fachabteilung'] in ('ZUL', 'IT', 'AAU', 'ADM', 'DIS', 'FIN', 'FZZ', 'PEZ',
                                                'TEC'), f"Invalid department in Excel {path_to_excel_file} in row ID {row['id']}"
                self.workspaces.append(Workspace(row['id'], row['libelle'], row['Standort'], row['Fachabteilung']))
                numbers_of_added_workspaces += 1

                # check if workspace name already exist and if so then throw an error
                list_of_workspace_names = []
                for workspace in self.workspaces:
                    if workspace.name not in (list_of_workspace_names):
                        list_of_workspace_names.append(workspace.name)
                    else:
                        logging.warning(
                            f"Error, the same workspace name ({workspace.name}) has been defined more then once in Excel ({path_to_excel_file})")

        logging.info(f"{numbers_of_added_workspaces} workspaces from {path_to_excel_file} sheet Arbeitsplatz added")
        #######################################################################################

    def _load_users_and_add_to_workspace_from_printerlist_of_department(self, path_to_excel_file: str, departmentname: str):
        #######################################################################################
        # Load users and add them to each workspace object in self.workspaces - sheet User zur Arbeitsplatz
        excellist = pd.read_excel(path_to_excel_file, sheet_name='User zu Arbeitsplatz')
        amount_of_users_added = 0
        for index, row in excellist.iterrows():

            if departmentname == "zul":
                if isinstance(row['Fachbereich'], str) and isinstance(row['Albisgüetli'], str):
                    for workspace in self.workspaces:
                        if workspace.department == departmentname.upper() and workspace.location == 'Albisgüetli':
                            amount_of_users_added += 1
                            workspace.add_user(row['User'])
                elif isinstance(row['Fachbereich'], str) and isinstance(row['Winterthur'], str):
                    for workspace in self.workspaces:
                        if workspace.department == departmentname.upper() and workspace.location == 'Winterthur':
                            amount_of_users_added += 1
                            workspace.add_user(row['User'])
                elif isinstance(row['Fachbereich'], str) and isinstance(row['Regensdorf'], str):
                    for workspace in self.workspaces:
                        if workspace.department == departmentname.upper() and workspace.location == 'Regensdorf':
                            amount_of_users_added += 1
                            workspace.add_user(row['User'])
                elif isinstance(row['Fachbereich'], str) and isinstance(row['Hinwil'], str):
                    for workspace in self.workspaces:
                        if workspace.department == departmentname.upper() and workspace.location == 'Hinwil':
                            amount_of_users_added += 1
                            workspace.add_user(row['User'])
                elif isinstance(row['Fachbereich'], str) and isinstance(row['Bülach'], str):
                    for workspace in self.workspaces:
                        if workspace.department == departmentname.upper() and workspace.location == 'Bülach':
                            amount_of_users_added += 1
                            workspace.add_user(row['User'])
                elif isinstance(row['Fachbereich'], str) and isinstance(row['Bassersdorf'], str):
                    for workspace in self.workspaces:
                        if workspace.department == departmentname.upper() and workspace.location == 'Bassersdorf':
                            amount_of_users_added += 1
                            workspace.add_user(row['User'])

            else:
                if isinstance(row['Fachbereich'], str):
                    for workspace in self.workspaces:
                        if workspace.department == departmentname.upper():
                            amount_of_users_added += 1
                            workspace.add_user(row['User'])


        logging.info(f"{amount_of_users_added} users added to workspaces from {path_to_excel_file}")
        #######################################################################################

    def _load_users_to_windowsprinter(self, path_to_excel_file: str):
        #######################################################################################
        excellist = pd.read_excel(path_to_excel_file, sheet_name='User zu Windowsprinter')
        for index, row in excellist.iterrows():
            if isinstance(row['User'], str) and isinstance(row['Printername1'], str):
                user = row['User']
                printer1 = row['Printername1']
                printer2 = None
                printer3 = None
                printer4 = None

                if isinstance(row['Printername2'], str):
                    printer2 = row['Printername2']

                if isinstance(row['Printername3'], str):
                    printer3 = row['Printername3']

                if isinstance(row['Printername4'], str):
                    printer4 = row['Printername4']

                #create a list of printers to iterate through
                printer_list = []
                printer_list.append(printer1)
                printer_list.append(printer2)
                printer_list.append(printer3)
                printer_list.append(printer4)

                for printer in printer_list:
                    if printer is not None:
                        a_printer = self.get_printer(printer)
                        a_printer.add_windowsuser(user) #calls the add_windowsuser-method of the printer class an adds user
        #######################################################################################

    def _self_load_pcs_to_default_windowsprinter(self, path_to_excel_file: str):
        #######################################################################################
        excellist = pd.read_excel(path_to_excel_file, sheet_name='PC zu Default-Windowsprinter')
        for index, row in excellist.iterrows():
            if isinstance(row['Printername'], str):
                pc = row['Name']
                printer = row['Printername']
                a_printer = self.get_printer(printer)
                a_printer.add_pc(pc)

        #######################################################################################

    def _add_cariform_to_printer_from_printerlist_of_department(self, path_to_excel_file: str):
        #####################################################################################################
        # Add CARi-Form to the printers in self.printers - sheet Arbeitsplatz-Formular-Drucker
        # Should a printer or a printerslot not exist (from the printer list of support) or the printerslot should have another paperformat, an error is thrown
        excellist = pd.read_excel(path_to_excel_file, sheet_name='Arbeitsplatz-Formular-Drucker')

        number_updated_printers_no_conflict = set()
        number_updated_printers_with_conflict = set()
        check_dict = {}
        warning_set = set()

        for index, row in excellist.iterrows():
            workspace_name = row['Arbeitsplatz']
            cariform = row['Formular']
            paperformat = row['Format']
            printername = row['Drucker']
            printerslot = row['Schacht']
            twosided = False
            inspect = False

            if isinstance(workspace_name, str):
                ###################################################################################################
                # check if printer printerslot to paperformat is not contradicting an other entry in the exelfile if it does log a warning
                new_key = printername + "_" + printerslot
                new_value = paperformat

                if new_key in check_dict:
                    orginal_value = check_dict[new_key]
                    check_dict[new_key] = new_value
                    if check_dict[new_key] != orginal_value:
                        logging.warning(
                            f"Warning printerslot {new_key} has contradicting values in excel {path_to_excel_file}")

                check_dict[new_key] = new_value
                ###################################################################################################
                # getting printer of the printerlist - if the printer is not on the list it will show an error
                # checking if printer exists in printerlist of support
                printer_from_selfprinters = self.get_printer(printername)
                # checking if the printer has this printerslot and if the printerslot has the same format in the printerlist of support
                printer_identical = printer_from_selfprinters.check_if_papersource_exists(printerslot=printerslot,
                                                                                          paperformat=paperformat)
                if printer_identical:
                    number_updated_printers_no_conflict.add(printername + "-" + printerslot)
                else:
                    number_updated_printers_with_conflict.add(printername + "-" + printerslot)
                    # add warning to a the warning_set that paper in a slot is not the same as in the support printer list
                    warning_set.add(
                        f"{printername} has {printerslot} : {paperformat} in {path_to_excel_file} but different value in printerlist of support")

                #####################################################################################################
                # updating papersource of printer with cariform and workspace
                for workspace in self.workspaces:
                    if workspace.name == workspace_name:
                        workspace.add_wcps(workspace.name, cariform, printername, printerslot)

                printer_from_selfprinters.update_papersource(printerslot=printerslot, paperformat=paperformat,
                                                             twosided=twosided, inspect=inspect,
                                                             active=None)
                #####################################################################################################

        # log the warnings and the number of updates printers with department-printer-list
        logging.info(
            f"{len(number_updated_printers_no_conflict)} papersources from {path_to_excel_file} updated with cariform=cariform, twosided=False, inspect=False, workspace=workspace")
        logging.warning(
            f"{len(number_updated_printers_with_conflict)} papersources from {path_to_excel_file} updated with paperformat=paperformat [CONFLICT], cariform=cariform, twosided=False, inspect=False, workspace=workspace")
        for warning in warning_set:
            logging.warning(warning)
        #####################################################################################################

    def _verify_workspaces(self):
        """loop for workspaces in self.workspaces and checks if all have users and wcps objects. if either one is missing it will log a warning in the logfile"""
        for workspace in self.workspaces:
            if len(workspace.wcps) == 0 or len(workspace.users) == 0:
                logging.warning(f"{workspace} has either no users or no wcps or both")

    def load_printerlist_of_department(self, path_to_excel_file: str, departname: str, load_user_to_windowsprinter: bool, load_pc_to_default_windowsprinter: bool):
        """load the excel file config-printers of each department with the workspace, users and cariforms"""
        #load workspace from printerlist of department in printermanager.workspaces
        self._load_workspaces_from_printerlist_of_department(path_to_excel_file)
        #load user from printerlist of deparment and add it to printermanaager.workspaces
        self._load_users_and_add_to_workspace_from_printerlist_of_department(path_to_excel_file, departname)
        self._add_cariform_to_printer_from_printerlist_of_department(path_to_excel_file)
        self._verify_workspaces()
        if load_user_to_windowsprinter:
            self._load_users_to_windowsprinter(path_to_excel_file)
        if load_pc_to_default_windowsprinter:
            self._self_load_pcs_to_default_windowsprinter(path_to_excel_file)

    def compare_two_printer_with_same_name_return_difference(self, printer1: Printer, printer2: Printer) -> tuple:
        """as input it accepts to printer and compares them. it returns a tuple. if all attributes of printer are identical it returns (true, [printer1]) in the form of a list. if there is a difference it returns a list with 3 printers. the
        first one is printer1, the second one is printer2 and the third one is a new printer object printer3 with the delta between printer1 and printer2 like (false, [printer1, printer2, deltaprinter])"""

        if printer1.printername == printer2.printername:
            standort = printer1.standort == printer2.standort #string
            if isinstance(printer1.buero, str):
                buero = printer1.buero == printer2.buero #string
            else:
                buero = True
            printername = printer1.printername == printer2.printername  #string
            ip = printer1.ip == printer2.ip #string
            papersources = True #TODO not implemented yet
            user_to_windowsprinter = printer1.user_to_windowsprinter == printer2.user_to_windowsprinter #set
            user_to_windowsprinter_for_cari = printer1.user_to_windowsprinter_for_cari == printer2.user_to_windowsprinter_for_cari #set
            pc_to_default_windowsprinter = printer1.pc_to_default_windowsprinter == printer2.pc_to_default_windowsprinter #set
            model = printer1.model == printer2.model    #string
            driver = printer1.driver == printer2.driver #string

            printer_identical = standort and buero and printername and ip and papersources and user_to_windowsprinter and user_to_windowsprinter_for_cari and pc_to_default_windowsprinter and model and driver
            returning_printer = [] #is a list of either one printer (if printers are identical) or a list of three printers like printer1, printer2, new_printer_object with delta between printer 1 and 2


            if printer_identical:
                printer1 = deepcopy(printer1)
                returning_printer.append(printer1)
            else:
                delta_user_to_windowsprinter = self._return_difference_of_two_sets(printer1.user_to_windowsprinter, printer2.user_to_windowsprinter)
                delta_user_to_windowsprinter_for_cari = self._return_difference_of_two_sets(printer1.user_to_windowsprinter_for_cari, printer2.user_to_windowsprinter_for_cari)
                delta_pc_to_default_windowsprinter = self._return_difference_of_two_sets(printer1.pc_to_default_windowsprinter, printer2.pc_to_default_windowsprinter)
                printer1 = deepcopy(printer1)
                printer2 = deepcopy(printer2)
                printer_delta = Printer(printer2.standort, printer2.buero, printer2.printername, printer2.ip, *printer2.papersource_for_pickle, printer2.model, delta_user_to_windowsprinter, delta_user_to_windowsprinter_for_cari, delta_pc_to_default_windowsprinter)

                printer_delta.papersources = deepcopy(printer2.papersources)

                returning_printer.append(printer1)
                returning_printer.append(printer2)
                returning_printer.append(printer_delta)

                logging.info(f"Printer {printer1.printername} has different attributes to last generated list")

            #print(f"standort:{standort}, buero:{buero}, printername:{printername}, ip:{ip}, papersources:{papersources}, user_to_windowsprinter:{user_to_windowsprinter}, user_to_windowsprinter_for_cari:{user_to_windowsprinter_for_cari}, pc_to_default_windowsprinter:{pc_to_default_windowsprinter}, model:{model}, driver:{driver}")

            return (printer_identical, returning_printer)

    def _return_difference_of_two_sets(self, set1: set, set2: set) -> set:
        present_in_set1_but_not_set2 = set1 - set2
        present_in_set2_but_not_set1 = set2 - set1
        return present_in_set1_but_not_set2.union(present_in_set2_but_not_set1)

    def compare_two_printerlists_and_return_difference(self, printerlist1: list, printerlist2: list) -> list:
        """the function returns a list of tuples like this [(true, printer2), (true, printer2), (false, printer1, printer2, printer_delta)....]
        when the printer is identical in printerlist1 and printerlist2 then (true, printer2) is returned. if it has not identical attributes
        then (false, printer1, printer2, printer_delta) is returned whre printer1 is the printer of printerlist1, printer2 is the printer of printerlist2 with changed attributes
        and printer_delta is the delta between them
        """
        printerlist = []
        for printer1 in printerlist1:
            for printer2 in printerlist2:
                compared_printer = self.compare_two_printer_with_same_name_return_difference(printer1, printer2)
                if printer1.printername == printer2.printername:
                    printerlist.append(compared_printer)
        return printerlist


    def add_cari_users_to_printer(self):
        """adds users from each papersource to the printer. for example printer pstva1111-s1 has b126kec und pastva1111-s2 has b126swm then both user gets added to pstva1111"""
        for printer in self.printers:
            for workspace in self.workspaces:
                for wcps in workspace.wcps:
                    if printer.printername == wcps.printername:
                        for user in wcps.workspace_users:
                            printer.add_windowsuser_for_cari(user)



