class WCPS:

    def __init__(self, workspace_name: str, cariform: str, printername:str, printerslot: str, department: str = None, workspace_id: str = None, workspace_user_list: list = None):
        self.workspace_name = workspace_name
        self.cariform = cariform
        self.printername = printername
        self.printerslot = printerslot

        if department is not None:
            self.department = department

        if workspace_id is not None:
            self.workspace_id = workspace_id

        if workspace_user_list is not None:
            self.workspace_users = workspace_user_list


    def __str__(self):
        return f"{self.cariform} -> {self.printername}:{self.printerslot}"

