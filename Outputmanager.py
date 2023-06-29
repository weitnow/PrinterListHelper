import os
import openpyxl
from Printermanager import Printermanager
import copy

class Outputmanager:

    def __init__(self, printermanager: Printermanager):
        self.printermanager = printermanager

    def return_deep_copy_of_printermanger_printers(self):
        return copy.deepcopy(self.printermanager.printers) #make a deep-copy of the list. this makes sure that we don't change the orignal list

    def create_output_excel_list_for_robot(self, path_with_filename: str, title_of_worksheet: str, list_with_header_names: list, printer_list: list):
        file_path = f'{path_with_filename}.xlsx'

        # check if the file exists and if it does delete it
        if os.path.exists(file_path):
            os.remove(file_path)

        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.title = title_of_worksheet

        # get the instance variable names and write them as headers
        for i, variable in enumerate(list_with_header_names):
            worksheet.cell(row=1, column=i + 1, value=variable)

        # create a list of dictionarys of printername and papersources {printername: pstva1234, M0: papersourceobj1, S1: papersourceobj2}
        list_of_dict_with_printername_and_papersources = []

        #iterate through all printers in printermanager.printers
        for i, instance in enumerate(printer_list):
            list_of_dict = []

            papersources = getattr(instance, "papersources")

            for papersource in papersources:
                copies_papersource_for_each_wcps = []
                num_of_wcps = len(papersource.wcps)  # 0 = no wcps
                if num_of_wcps > 0:
                    for i in range(num_of_wcps): #generate as many copies of new dict as wcps in papersource
                        new_dict = {}   #generate a new_dict for each papersource
                        new_dict["Schacht Name"] = papersource.printerslot
                        new_dict["Druckername Printerserver"] = f"{instance.printername}-{papersource.printerslot}"
                        new_dict["Format des Formulars"] = papersource.paperformat
                        new_dict["Aktiv"] = papersource.active

                        if papersource.inspect == True:
                            new_dict["Inspect"] = "x"
                        elif papersource.inspect == "CUT":
                            new_dict["Inspect"] = "CUT"
                        else:
                            new_dict["Inspect"] = ""

                        if papersource.twosided == True:
                            new_dict["2-sided"] = "2-sided"
                        else:
                            new_dict["2-sided"] = "None"

                        copies_papersource_for_each_wcps.append(new_dict) #now we have for each wcps a copy of the papersource dict with all the same values



                    #loop through wcps of papersource. wcps contains a tuple (workspace_name, cari-doc, printername, printerslot, department)
                    temp_wcps_list = [] # now we have a list of tuples of wcps which have to be added to copies_papersource
                    for wcps in papersource.wcps:
                        temp_wcps_list.append(wcps)

                    for i, dict in enumerate(copies_papersource_for_each_wcps):
                        dict["Formular CARI / Prüfbahn / Parkplatz"] = temp_wcps_list[i][1]  #tuple(1) = caridoc
                        dict["Zuständige Fachabteilung"] = temp_wcps_list[i][4]              #tuple(4) = deparment
                        dict["Arbeitsplatz (Büro)"] = temp_wcps_list[i][0]                   #tuple(0) = workspace_name

                    for item in copies_papersource_for_each_wcps:
                        list_of_dict.append(item)
                else:
                    new_dict = {}  # generate a new_dict for each papersource
                    new_dict["Schacht Name"] = papersource.printerslot
                    new_dict["Druckername Printerserver"] = f"{instance.printername}-{papersource.printerslot}"
                    new_dict["Format des Formulars"] = papersource.paperformat
                    new_dict["Aktiv"] = papersource.active

                    if papersource.inspect == True:
                        new_dict["Inspect"] = "x"
                    elif papersource.inspect == "CUT":
                        new_dict["Inspect"] = "CUT"
                    else:
                        new_dict["Inspect"] = ""

                    if papersource.twosided == True:
                        new_dict["2-sided"] = "2-sided"
                    elif papersource.twosided == False:
                        new_dict["2-sided"] = "None"
                    else:
                        new_dict["2-sided"] = ""

                    list_of_dict.append(new_dict)

            for dict in list_of_dict:
                dict["Standort"] = instance.standort
                dict["Bemerkung"] = instance.buero
                dict["Printserver Link"] = "\\\\vstvacp100001\\"
                dict["Druckername"] = instance.printername
                dict["IP Drucker (nur zur Information für Vergleich QIP)"] = instance.ip
                dict["Portname"] = instance.printername
                dict["Drucker Modell"] = instance.model
                dict["Drucker Treiber"] = instance.driver

                list_of_dict_with_printername_and_papersources.append(dict)

        # write the instance variable values as rows
        for i, dictionary in enumerate(list_of_dict_with_printername_and_papersources):
            value_list = []
            for variable in list_with_header_names:
                value_list.append(dictionary.get(variable))
            for j, value in enumerate(value_list):
                worksheet.cell(row=i + 2, column=j + 1, value=value)
        workbook.save(file_path)

    def create_output_excel_list_for_serdar(self, path_with_filename: str, title_of_worksheet: str, list_with_header_names: list, printer_list: list):
        file_path = f'{path_with_filename}.xlsx'

        # check if the file exists and if it does delete it
        if os.path.exists(file_path):
            os.remove(file_path)

        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.title = title_of_worksheet

        # get the instance variable names and write them as headers
        for i, variable in enumerate(list_with_header_names):
            worksheet.cell(row=1, column=i + 1, value=variable)

        for i, dictionary in enumerate(printer_list):
            value_list = []
            for variable in list_with_header_names:
                value_list.append(dictionary.get(variable))
            for j, value in enumerate(value_list):
                if (isinstance(value, list)):
                    value = ','.join(value)
                worksheet.cell(row=i + 2, column=j + 1, value=value)
        workbook.save(file_path)



    def delete_printer_without_wcps(self, printer_list: list) -> list:
        """this function deletes all printerslots of a printer with no wcps. it means it is a printer used with
        cari. if a printer has not at least one printerslot with a wcps the whole printer is deleted"""
        temp_printers = []
        for printer in printer_list:
            temp_papersources = []
            for papersource in printer.papersources:
                #if papersource has no wcps objects do nothing
                if len(papersource.wcps) == 0:
                    continue
                #otherwise add the papersource to the temp_papersources_list
                else:
                    temp_papersources.append(papersource)

            #if temp_papersources has no items do nothing. it means there were no wcps present in papersource
            if temp_papersources == 0:
                continue
            #otherwise replace the papersources of the printer only with the papersources with a wcps
            else:
                printer.papersources = temp_papersources

        # now we ended up with printers without any papersources because for some printers there was not a single papersource with a wcps
        # we loop through the printers again and add all the printers to the temp_printers list which have at least one papersource (with a wcps)

        for printer in printer_list:
            # if the printer has no papersource left we dont add it to temp_printers
            if len(printer.papersources) == 0:
                continue
            # otherwise we add it
            else:
                temp_printers.append(printer)

        # no we switch self.printers with temp_printers
        printer_list = temp_printers

        return printer_list

    def delete_all_wcps(self, printer_list: list) -> list:
        """deletes all wcps of printers. papersource will remain. it means we end up with a printerlist without
        caridocs and workspaces attached"""
        for printer in printer_list:
            for papersource in printer.papersources:
                papersource.wcps = set()

        return printer_list

    def delete_papersource_with_inspect(self, printer_list: list) -> list:
        """deletes all papersource with inspect == true. if printer has no more papersource the whole printer gets deleted"""
        temp_printers = []
        for printer in printer_list:
            temp_papersources = []
            for papersource in printer.papersources:
                # if papersource has inspect == true do nothing
                if papersource.inspect == True or papersource.inspect == "CUT":
                    continue
                # otherwise add the papersource to the temp_papersources_list
                else:
                    temp_papersources.append(papersource)

            # if temp_papersources has no items do nothing. it means there were no papersources without inspect == true
            if temp_papersources == 0:
                continue
            # otherwise replace the papersources of the printer only with the papersources with inpsect == false
            else:
                printer.papersources = temp_papersources

        for printer in printer_list:
            # if the printer has no papersource left we dont add it to temp_printers
            if len(printer.papersources) == 0:
                continue
            # otherwise we add it
            else:
                temp_printers.append(printer)

        # no we switch self.printers with temp_printers
        printer_list = temp_printers

        return printer_list


















