import xlwings
import random

excel_app = xlwings.App(visible=False)
excel_book = excel_app.books.open('Dane_S2_50_10.xlsx')
tasks = 50


def reset():
    for x in range(tasks):
        excel_book.sheets['Arkusz1'].range(x+2, 13).value = x+1


def TabuSearch(iterations):
    for x in range(iterations):
        globalMin = int(excel_book.sheets['Arkusz1'].range('W51').value)
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
        tabulist1 = []
        tabulist2 = []
        tabuWait = []
        for i in tabulist1:
            if((tabulist1[i] == rand1 - 1 and tabulist2[i] == rand2 - 1) or ((tabulist1[i] == rand2 - 1 and tabulist2[i] == rand1 - 1))):
                inTabu = True
        if (inTabu == True):
            continue

        cell1 = excel_book.sheets['Arkusz1'].range(rand1, 13).value
        cell2 = excel_book.sheets['Arkusz1'].range(rand2, 13).value
        excel_book.sheets['Arkusz1'].range(rand1, 13).value = cell2
        excel_book.sheets['Arkusz1'].range(rand2, 13).value = cell1

        # wb.save('Dane_S2_50_10.xlsx')
        localMin = int(excel_book.sheets['Arkusz1'].range('W51').value)

        print(globalMin, localMin, globalMin < localMin)

        if(localMin < globalMin or localMin == globalMin):
            for i in tabulist1:
                tabuWait[i] = tabuWait[i] - 1
            tabulist1.append(rand1 - 1)
            tabulist2.append(rand2 - 1)
            tabuWait.append(5)
        elif (localMin > globalMin):
            excel_book.sheets['Arkusz1'].range(rand1, 13).value = cell1
            excel_book.sheets['Arkusz1'].range(rand2, 13).value = cell2
            if not tabulist1:
                for i in tabulist1:
                    tabuWait[i] = tabuWait[i] - 1
            else:
                continue
        if not tabulist1:
            for i in tabulist1:
                print(tabuWait[i] == 0)
                if tabuWait[i] == 0:
                    tabuWait.pop(i)
                    tabulist1.pop(i)
                    tabulist2.pop(i)
        print(tabulist1, tabulist2, tabuWait)
        excel_book.save()


# reset()
TabuSearch(100)


excel_book.save()
excel_book.close()
excel_app.quit()