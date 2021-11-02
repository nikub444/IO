import xlwings
import random

excel_app = xlwings.App(visible=False)
excel_book = excel_app.books.open('Dane_S2_100_20.xlsx')
tasks = 100
tabulist1 = []
tabulist2 = []
tabuWait = []


def reset():
    for x in range(tasks):
        excel_book.sheets['Arkusz1'].range(x+2, 23).value = x+1


def TabuSearch(iterations, wait):
    for x in range(iterations):
        globalMin = int(excel_book.sheets['Arkusz1'].range('AQ101').value)
        inTabu = False
        # losujemy dwie różne liczby
        while True:
            rand1 = int((tasks - 1 + 1) * random.random() + 1) + 1
            rand2 = int((tasks - 1 + 1) * random.random() + 1) + 1
            # print(rand1 != rand2)
            # print(rand1, rand2)
            if (rand1 != rand2):
                # print(rand1, rand2)
                break

        for i in range(len(tabulist1)):
            if((tabulist1[i] == rand1 - 1 and tabulist2[i] == rand2 - 1) or ((tabulist1[i] == rand2 - 1 and tabulist2[i] == rand1 - 1))):
                inTabu = True
        print(inTabu)
        if (inTabu == True):
            continue

        cell1 = excel_book.sheets['Arkusz1'].range(rand1, 23).value
        cell2 = excel_book.sheets['Arkusz1'].range(rand2, 23).value
        excel_book.sheets['Arkusz1'].range(rand1, 23).value = cell2
        excel_book.sheets['Arkusz1'].range(rand2, 23).value = cell1

        # wb.save('Dane_S2_50_10.xlsx')
        localMin = int(excel_book.sheets['Arkusz1'].range('AQ101').value)

        print(globalMin, localMin, localMin <
              globalMin or localMin == globalMin)

        if(localMin < globalMin or localMin == globalMin):
            print(tabulist1, tabulist2, tabuWait)
            for i in range(len(tabulist1)):
                tabuWait[i] = tabuWait[i] - 1
            tabulist1.append(rand1 - 1)
            tabulist2.append(rand2 - 1)
            tabuWait.append(wait)
        elif (localMin > globalMin):
            excel_book.sheets['Arkusz1'].range(rand1, 23).value = cell1
            excel_book.sheets['Arkusz1'].range(rand2, 23).value = cell2
            if len(tabulist1) != 0:
                for i in range(len(tabulist1)):
                    tabuWait[i] = tabuWait[i] - 1
            else:
                continue
        if len(tabulist1) != 0:
            for i in range(len(tabulist1)):
                if tabuWait[i] == 0:
                    tabuWait.pop(0)
                    tabulist1.pop(0)
                    tabulist2.pop(0)
                    break

        print(tabulist1, tabulist2, tabuWait)
        excel_book.save()


reset()
TabuSearch(100, 5)


excel_book.save()
excel_book.close()
excel_app.quit()
