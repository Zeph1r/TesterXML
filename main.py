import os.path
import random
import pandas as pd
from openpyxl import load_workbook

def test(text, var_true, var1, var2, var3=None):
    '''Функция тестирования

    На вход получает лишь текст вопроса и максимум 4 варианта ответа.
    Первый вариант должен быть обязательно правильный.
    4-й вариант не обязательно вводить если 4-го варианта нету.'''
    print(text)
    ansver = [var_true, var1, var2, var3]
    random.shuffle(ansver)
    for r in ansver:
        if r == None:
            ansver.remove(r)
    Number = 1
    for i in ansver:
        print(str(Number) + "." + str(i))
        Number += 1
    ans = input("Введите правильный ответ: (1,2,3,4):")
    if ans == "1":
        if ansver[0] == var_true:
            print("Правильно")
            return True
        else:
            print("Не правильно")
            return False
    elif ans == "2":
        if ansver[1] == var_true:
            print("Правильно")
            return True
        else:
            print("Не правильно")
            return False
    elif ans == "3":
        if ansver[2] == var_true:
            print("Правильно")
            return True
        else:
            print("Не правильно")
            return False
    elif len(ansver) == 4 and ans == "4":
        if ansver[3] == var_true:
            print("Правильно")
            return True
        else:
            print("Не правильно")
            return False
    else:
        print("Ответ неверен или неправильно введён.")

def create_excel():
    '''Создаёт excel шаблон для заполнения'''
    df = pd.DataFrame({'Вопрос': [],
                       'Ответ1*Правильный*': [],
                       'Ответ2': [],
                       'Ответ3': [],
                       'Ответ4(не обязательно)': []})
    df.to_excel('./test.xlsx')

def read_excel():
    '''Считывает вопрос с вариантами ответа из excel листа'''
    question = []
    quest = "B"
    ask_true = "C"
    ask2 = "D"
    ask3 = "E"
    ask4 = "F"
    count = 2
    CountQuestion = 0
    wb = load_workbook('./test.xlsx')
    sheet = wb.get_sheet_by_name('Sheet1')
    work = True
    while work:
        if sheet[quest+str(count)].value:
            question.append([])
            question[CountQuestion].append(sheet[quest + str(count)].value)
            question[CountQuestion].append(sheet[ask_true + str(count)].value)
            question[CountQuestion].append(sheet[ask2 + str(count)].value)
            question[CountQuestion].append(sheet[ask3 + str(count)].value)
            question[CountQuestion].append(sheet[ask4 + str(count)].value)
        else:
            work = False
        count += 1
        CountQuestion += 1
    return question

work = True
while work:
    print("\n\nДобро пожаловать в программу по тестированию знаний!")
    print("Выберите чтобы вы хотели сделать:\n1. Создать шаблон для заполнения\n2. Провести тестирование\n3. Выход")
    ans = input()
    if ans == "1":
        create_excel()
    elif ans == "2":
        if os.path.exists('test.xlsx'):
            Test = read_excel()
            count_bad = 0
            for i in range(len(Test)):
                if test(Test[i][0], Test[i][1], Test[i][2], Test[i][3], Test[i][4]) == False:
                    count_bad += 1
            if count_bad > 0:
                print("Всего было допущено ошибок ", count_bad)
            else:
                print("Поздравляю, все ответы верны!")
        else:
            print("Не найден файл с вопросами, пожалуйста для начала создайте его и заполните")
    elif ans == "3":
        work = False
    else:
        print('Неверный вариант выбора')