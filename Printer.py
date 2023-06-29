from Papersource import Papersource

# PRINTERMODEL TO DRIVER MAPPING
# if the model starts with name of key then set it to driver name "value of dict"
# for example: if Model starts with "Brother" then set "Brother HL-6250DN series" as drivername
DRIVERMAPPING = {
    "Brother" : "Brother HL-6250DN series",
    "Canon" : "Canon Generic Plus PCL6"
}

class Printer:

    #init method will be used by printerlist of support
    def __init__(self, standort, buero, printername, ip, M0, S1, S2, S3, S4, S5, model):
        self.standort = standort
        self.buero = buero
        self.printername = printername
        self.ip = ip
        self.papersources = []
        self.model = model
        self.driver = None

        #setting driver according to mapping
        for key in DRIVERMAPPING.keys():
            if self.model.startswith(key):
                self.driver = DRIVERMAPPING[key]

        #adding papersource in self.papersources
        if isinstance(M0, str):
            self.papersources.append(Papersource(printerslot="M0", paperformat=M0))
        if isinstance(S1, str):
            self.papersources.append(Papersource(printerslot="S1", paperformat=S1))
        if isinstance(S2, str):
            self.papersources.append(Papersource(printerslot="S2", paperformat=S2))
        if isinstance(S3, str):
            self.papersources.append(Papersource(printerslot="S3", paperformat=S3))
        if isinstance(S4, str):
            self.papersources.append(Papersource(printerslot="S4", paperformat=S4))
        if isinstance(S5, str):
            self.papersources.append(Papersource(printerslot="S5", paperformat=S5))

    def __str__(self):
        outputstring = f"{self.printername}|Standort:{self.standort}|Buero:{self.buero}|Model:{self.model}|Driver:{self.driver}|IP:{self.ip}|"

        for papersource in self.papersources:
            outputstring += "["
            outputstring += str(papersource) + "|"
            outputstring += "]"

        return(outputstring)

    def get_users_paperslots_workspaces(self, printermanager) -> dict:
        if printermanager is None:
            raise Exception("Need printermanager to access printermanager.workspaces")

        my_dict = dict()

        my_dict["printername"] = self.printername

        my_dict["paperslots"] = []
        for papersource in self.papersources:
            my_dict["paperslots"].append(papersource.printerslot)

        my_dict["workspaces"] = set()
        my_dict["users"] = set()
        for workspace in printermanager.workspaces:
            for wcps in workspace.wcps:
                if wcps.printername == self.printername:
                    my_dict["workspaces"].add(workspace.name)

                    for users in workspace.users:
                        my_dict["users"].add(users)

        #cast sets to lists
        my_dict["workspaces"] = list(my_dict["workspaces"])
        my_dict["users"] = list(my_dict["users"])
        return my_dict




    def check_if_papersource_exists(self, *, printerslot: str, paperformat: str = None, twosided: bool = None, inspect: bool = None, active: int = None) -> bool:
        """checks if a papersource with the provided parameters already exist. printerslot mus be provided, everyhting else is optional. if its not provided or None, then this
        parameter will not be checked if already set and if they are the same as the provided values. it returns True or False"""
        papersource_found = False
        paperformat_is_same = False
        twosided_is_same = False
        inspect_is_same = False
        active_is_same = False
        for papersource in self.papersources:
            if papersource.printerslot == printerslot:
                papersource_found = True
                if paperformat is not None:
                    paperformat_is_same = papersource.paperformat == paperformat
                else:
                    paperformat_is_same = True


                if twosided is not None:
                    twosided_is_same = papersource.twosided == twosided
                else:
                    twosided_is_same = True

                if inspect is not None:
                    inspect_is_same = papersource.inspect == inspect
                else:
                    inspect_is_same = True

                if active is not None:
                    active_is_same = papersource.active == active
                else:
                    active_is_same = True

        if all([papersource_found, paperformat_is_same, twosided_is_same, inspect_is_same, active_is_same]):
            return True
        else:
            return False


    def update_papersource(self, *, printerslot: str, paperformat: str = None, twosided: bool = None, inspect: bool = None, active: int = None):

        for papersource in self.papersources:
            if printerslot == papersource.printerslot:

                if paperformat is not None:
                    assert paperformat in (
                    "A3", "A4", "A5", "A6"), "invalid paperformat, format must be A3, A4, A5 or A6"
                    papersource.paperformat = paperformat

                if twosided is not None:
                    assert twosided in (
                    True, False), "twosided must be either True or False. True = Duplexprint, False = Simplexprint"
                    papersource.twosided = twosided

                if inspect is not None:
                    assert inspect in (
                    True, False), "inspect must bei either True or False. True = Printer used for Inspect"
                    papersource.inspect = inspect

                if active is not None:
                    assert active in (1, 2, 3, 4), "active must be 1, 2, 3, or 4"
                    papersource.active = active


    def set_twosided_from_None_to_True(self):
        """set the the attribute twosided of every papersource of the printer from None to True"""
        for papersource in self.papersources:
            if papersource.twosided is None:
                papersource.twosided = True

    def set_inspect_from_None_to_False(self):
        """set the the attribute inspect of every papersource of the printer from None to True"""
        for papersource in self.papersources:
            if papersource.inspect is None:
                papersource.inspect = False

    def set_active_from_None_to_1(self):
        """set the the attribute inspect of every papersource of the printer from None to True"""
        for papersource in self.papersources:
            if papersource.active is None:
                papersource.active = 1

    def add_wcps(self, printername: str, printerslot: str, workspace_name: str, cariform: str, department: str = None, workspace_id: str = None, workspace_user_list: list = None):
        if printername == self.printername:
            printerslot_not_found = True
            for papersource in self.papersources:
                if papersource.printerslot == printerslot:
                    printerslot_not_found = False
                    papersource.wcps.add((workspace_name, cariform, printername, printerslot, department, workspace_id))

            if printerslot_not_found:
                raise Exception(f"{printername} has no printerslot {printerslot}")
        else:
            raise Exception(f"add_wcps: received {printername} as printername but self.printername is {self.printername}")



