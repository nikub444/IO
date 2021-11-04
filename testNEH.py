import xlwings
import random

excel_app = xlwings.App(visible=False)
excel_book = excel_app.books.open('Dane_S2_50_10NEH.xlsx')
sortedJobs = []
for i in range(2, 52):
    sortedJobs.append(int(excel_book.sheets['Arkusz2'].range(i, 1).value))
# print(sortedJobs)

for i in range(2, 52):
    excel_book.sheets['Arkusz1'].range(i, 13).value = sortedJobs[i-2]

permutations = []
makespan = []
permutations.append(int(excel_book.sheets['Arkusz1'].range(2, 13).value))
# len perm 1
for x in range(51):
    for i in range(len(permutations)+1):
        newTask = int(excel_book.sheets['Arkusz1'].range(
            2+len(permutations), 13).value)
        print("New task=", newTask)
        permutations.insert(i, newTask)
        print("Perm=", permutations)
        for j in range(len(permutations)):
            excel_book.sheets['Arkusz1'].range(j+2, 13).value = permutations[j]
        makespan.append(
            int(excel_book.sheets['Arkusz1'].range(1+len(permutations), 23).value))
        permutations.pop(i)
        for k in range(len(permutations)):
            excel_book.sheets['Arkusz1'].range(k+2, 13).value = permutations[k]
        excel_book.sheets['Arkusz1'].range(
            len(permutations)+2, 13).value = newTask
        print("Perm after calc=", permutations)
        print("Makespan=", makespan)

        # for i in range(2, 5):
        #     for j in range(2, i+2):
        #         permutations.append(
        #             int(excel_book.sheets['Arkusz1'].range(j, 13).value))
        #         print("i=", i, "j=", j)
    minMakespan = min(makespan)
    index = makespan.index(minMakespan)
    makespan.clear()
    permutations.insert(i, newTask)
    print(permutations)
    print(excel_book.sheets['Arkusz1'].range(len(permutations)+1, 23).value)


excel_book.save()
excel_book.close()
excel_app.quit()
