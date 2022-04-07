import gspread
import time

sa = gspread.service_account()
sh = sa.open("Оценка занятия (Ответы)")

wks = sh.worksheet("Ответы на форму (1)")
wks2 = sh.worksheet('Формативные сообщения')
alf = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'


def fun(l1, l2):
    feedback = []
    i = l1
    dig = str(i)

    m = 1

    name = wks.acell('B'+dig).value
    lesson_link = wks.acell('D'+dig).value
    fb = wks.acell('E'+dig).value

    grades = []
    a = 5

    el = wks.acell(alf[a]+dig).value

    while el != None:
        grades.append(int(el))
        a += 1
        el = wks.acell(alf[a]+dig).value

    k = 5

    grades_str = ''
    for i in grades:
        grades_str += (wks.acell(alf[k]+'1').value + ": " + str(i) + '\n')
        k += 1


    full_fb = fb + '\n\n' + grades_str + '\n\n' + "Ссылка на урок: " + lesson_link + '\n\n' + "Средний балл за урок: " + str(round(sum(grades)/len(grades), 2))
    print(name)
    wks2.update('A'+str(l2), name)
    wks2.update('B'+str(l2), full_fb)


from gspread_formatting  import *
import gspread
import time

sa = gspread.service_account()
sh = sa.open("Оценка занятия (Ответы)")

wks = sh.worksheet("Ответы на форму (1)")
wks2 = sh.worksheet('Формативные сообщения')
alf = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'


def fun(l1, l2):
    feedback = []
    i = l1
    dig = str(i)

    m = 1

    name = wks.acell('B'+dig).value
    lesson_link = wks.acell('D'+dig).value
    fb = wks.acell('E'+dig).value

    grades = []
    a = 5

    el = wks.acell(alf[a]+dig).value

    while el != None:
        grades.append(int(el))
        a += 1
        el = wks.acell(alf[a]+dig).value

    k = 5

    grades_str = ''
    for i in grades:
        grades_str += (wks.acell(alf[k]+'1').value + ": " + str(i) + '\n')
        k += 1

    full_fb = str(fb) + '\n\n' + grades_str + '\n\n' + "Ссылка на урок: " + lesson_link + '\n\n' + "Средний балл за урок: " + str(round(sum(grades)/len(grades), 2)).replace('.', ',')
    print(name)
    wks2.update('A'+str(l2), name)
    wks2.update('B'+str(l2), full_fb)


start = input("С какой строки начинать? ")

while wks.acell('A'+start)!=None:
    print("loading...")
    fun(int(start), int(start)-1)
    start = str(int(start)+1)
    time.sleep(60)


