from lark.visitors import Interpreter

class FormulaFunctionHandler(Interpreter):

    def __init__(self, default_sheet):
        self.children = set()
        self.default_sheet = default_sheet

    def cell(self, args):
        if len(args) == 2:
            cell_sheet = str(args[0])
            cell_ref = str(args[1])
        else:
            cell_sheet = self.default_sheet
            cell_ref = str(args[0])

        self.children.add((cell_sheet, cell_ref))

        self.visit_children(args)

    def function_expr(self, args):
        if 
