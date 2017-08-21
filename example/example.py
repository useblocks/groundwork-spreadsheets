from groundwork_spreadsheets import ExcelValidationPattern
from groundwork import App


def Application():
    app = App(plugins=[], strict=True)
    return app


class ReadCustomExcel(ExcelValidationPattern):
    def __init__(self, app, name=None, *args, **kwargs):
        self.name = name or self.__class__.__name__
        super(ReadCustomExcel, self).__init__(app, *args, **kwargs)

    def activate(self):
        pass

    def deactivate(self):
        pass


if __name__ == '__main__':
    app = App(plugins=[], strict=True)
    plugin = ReadCustomExcel(app)
    data = plugin.excel_validation.read_excel('config.json', 'example.xlsx')
    for row in data:
        headers = data[row]
        for header in headers:
            print("Row {0}, Header '{1}': {2}".format(row, header, data[row][header]))
