class User():
    def __init__(self, username: str):
        self.username = username
        self.number_cari_workspaces = 0
        self.number_printers = 0
        self.number_printerslots = 0


    def count_cari_workspace(self):
        self.number_cari_workspaces += 1

    def count_printer(self):
        self.number_printers += 1

    def count_printerslot(self):
        self.number_printerslots += 1

