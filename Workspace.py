import logging

from WCPS import WCPS

class Workspace:

    def __init__(self, id: int, name: str, location: str = None, department: str = None):
        self.users = set()
        self.wcps = set()

        self.id = id
        self.name = name
        self.location = location
        self.department = department

    def __str__(self):
        show_string = f"Workspace ID {self.id}, {self.name}, Location: {self.location}, Department: {self.department}"
        if len(self.users) > 0:
            show_string += f" users {len(self.users)}"
        if len(self.wcps) > 0:
            show_string += f" wcps {len(self.wcps)}"
        return show_string

    def add_user(self, username: str):
        self.users.add(username)

    def remove_user(self, username: str):
        self.users.remove(username)

    def show_users(self):
        print(self.users)

    def add_wcps(self, workspace_name: str, cariform: str, printername:str, printerslot: str):
        if self.name == workspace_name:
            userlist = list(self.users)
            # before adding a new wpcs I have to check, if there is not a existing workspace with the same name and the same caridoc which points to a diffrent printer and printerslot
            for wcps in self.wcps:
                if wcps.workspace_name == workspace_name and wcps.cariform == cariform:
                    if wcps.printername == printername and wcps.printerslot == printerslot:
                        continue
                    else:
                        logging.warning(f"Warning there is already an existing {workspace_name} with {cariform} pointing to annother printerslot")

            self.wcps.add(WCPS(workspace_name, cariform, printername, printerslot, self.department, self.id, userlist))

        else:
            raise Exception(f"self.name is {self.name} and the provided workspace_name is {workspace_name}")




