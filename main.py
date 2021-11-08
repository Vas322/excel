import openpyxl

wb = openpyxl.reader.excel.load_workbook(filename='test.xlsx', data_only=True)
count_of_lists = len(wb.sheetnames)  # количество листов в книге


def get_count_of_elements(list):
    count = 0
    for i in list:
        count += 1
    return count


def input_lists(msg):
    input_data = input(msg)
    if not input_data.isdigit(): return input_lists("Вы ввели не число. Введите целое положительное число: ")
    if 0 <= int(input_data) <= count_of_lists: return int(input_data) - 1
    return input_lists(f'Всего листов в книге: {count_of_lists}.'
                       f' Введите число от 1 до {count_of_lists}. ')


def win_rate():
    projects = []
    projects_win = []
    list_excel = input_lists("Введите номер листа книги Ecxel: ")
    sheet = wb.worksheets[list_excel]  # Обращение к определенному листу. 0 - это 1-й лист
    for row in range(3, sheet.max_row + 1):  # max_row - проходит до конца значений по столбцу
        id_project = sheet[row][0].value  # 0,1,2...n - идентификаторы столбцов слева направо
        if id_project is not None:
            projects.append(id_project)
        id_project_win = sheet[row][1].value
        if id_project_win is not None:
            projects_win.append(id_project_win)
    projects = set(projects)
    projects_win = set(projects_win)
    count_projects = get_count_of_elements(projects)
    count_projects_win = get_count_of_elements(projects_win)
    if count_projects != 0:
        projects_win_rate = round(count_projects_win / count_projects * 100, 1)  # round - округление до одного знака
    else:
        projects_win_rate = 0
        return print(f'Количество зарегистрированных проектов равно {projects_win_rate}. Проверьте файл!')
    print(f'Проекты: {projects},\nЗарегистрированных проектов: {count_projects}')
    print(f'Отгруженные проекты: {projects_win},\nВсего выигранных проектов: {count_projects_win}')
    print(f'WinRate равен: {projects_win_rate} %')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    win_rate()
