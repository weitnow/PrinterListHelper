
class Papersource():

    def __init__(self, printerslot: str, paperformat: str, twosided: bool = None, inspect: bool = None, active: int = None):

        assert printerslot in (
        "M0", "S1", "S2", "S3", "S4", "S5"), "invalid papersource name, name must be MO, S1, S2, S3, S4 or S5"
        self.printerslot = printerslot

        assert paperformat in ("A3", "A4", "A5", "A6"), "invalid paperformat, format must be A3, A4, A5 or A6"
        self.paperformat = paperformat

        self.twosided = twosided
        self.inspect = inspect
        self.active = active

        self.wcps = set() #contains a tuple (workspace_name, cari-doc, printername, printerslot, department, workspace_id)

    def __str__(self):
        output_string = f"{self.printerslot}:{self.paperformat}|2sided {self.twosided}|active {self.active}"
        if len(self.wcps) > 0:
            output_string += f" wcps:{len(self.wcps)}"
        return output_string

    def __eq__(self, other):
        return self.printerslot == other.printerslot and self.paperformat == other.paperformat

    def check_if_twosided_is_none(self) -> bool:
        return self.twosided == None

    def check_if_inspect_is_none(self) -> bool:
        return self.inspect == None

    def check_if_active_is_none(self) -> bool:
        return self.active == None

    def check_if_wcps_is_none(self) -> bool:
        return len(self.wcps) == 0




