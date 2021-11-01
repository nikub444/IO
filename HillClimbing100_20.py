import xlwings
import random

excel_app = xlwings.App(visible=False)
excel_book = excel_app.books.open('Dane_S2_100_20.xlsx')
tasks = 100


def reset():
    for x in range(tasks):
        excel_book.sheets['Arkusz1'].range(x+2, 23).value = x+1


def HillClimbing(iterations):
    for x in range(iterations):
        globalMin = int(excel_book.sheets['Arkusz1'].range('AQ101').value)
        # losujemy dwie różne liczby
        while True:
            rand1 = int((tasks - 1 + 1) * random.random() + 1) + 1
            rand2 = int((tasks - 1 + 1) * random.random() + 1) + 1
            # print(rand1 != rand2)
            # print(rand1, rand2)
            if (rand1 != rand2):
                # print(rand1, rand2)
                break
        cell1 = excel_book.sheets['Arkusz1'].range(rand1, 23).value
        cell2 = excel_book.sheets['Arkusz1'].range(rand2, 23).value
        excel_book.sheets['Arkusz1'].range(rand1, 23).value = cell2
        excel_book.sheets['Arkusz1'].range(rand2, 23).value = cell1
        # wb.save('Dane_S2_50_10.xlsx')
        localMin = int(excel_book.sheets['Arkusz1'].range('AQ101').value)
        print(globalMin, localMin, globalMin < localMin)
        if(globalMin < localMin):
            excel_book.sheets['Arkusz1'].range(rand1, 23).value = cell1
            excel_book.sheets['Arkusz1'].range(rand2, 23).value = cell2
        excel_book.save()


reset()
# HillClimbing(100)


excel_book.save()
excel_book.close()
excel_app.quit()
