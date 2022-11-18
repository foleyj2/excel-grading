import xlwings as xw


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    if sheet["A1"].value == "Hello xlwings!":
        sheet["A1"].value = "Bye xlwings!"
    else:
        sheet["A1"].value = "Hello xlwings!"

    # call
    macro_call = wb.macro("PrepareAndPrint")
    macro_call()#put arguments here if needed

        

@xw.func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("excel_makepdfs.xlsm").set_mock_caller()
    main()
