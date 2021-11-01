import xlwings


def Min():
    excel_app = xlwings.App(visible=False)
    excel_book = excel_app.books.open('Dane_S2_50_10.xlsx')
    min = int(excel_book.sheets['Arkusz1'].range('W51').value)
    excel_book.save()
    excel_book.close()
    excel_app.quit()
    return min
