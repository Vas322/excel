import openpyxl


def get_count_of_elements(list):
    count = 0
    for i in list:
        count += 1
    return count


def win_rate():
    wb = openpyxl.reader.excel.load_workbook(filename='test.xlsx', data_only=True)
    projects = []
    projects_win = []
    sheet = wb.worksheets[0]  # Обращение к определенному листу. 0 - это 1-й лист
    for row in range(2, sheet.max_row + 1):
        # max_row - проходит до конца значений по столбцу
        id_project = sheet[row][0].value  # 0,1,2...n - идентификаторы столбцов слева направо
        if id_project is not None:
            projects.append(id_project)
        id_project_win = sheet[row][1].value
        if id_project_win is not None:
            projects_win.append(id_project_win)
    count_projects = get_count_of_elements(projects)
    count_projects_win = get_count_of_elements(projects_win)
    if count_projects != 0:
        projects_win_rate = round(count_projects_win / count_projects * 100, 1)  # round - округление до одного знака
    else:
        projects_win_rate = 0
        print(f'Количество зарегистрированных проектов равно {projects_win_rate}. Проверьте файл!')
    print(f'Проекты: {projects},\nЗарегистрированных проектов: {count_projects}')
    print(f'Отгруженные проекты: {projects_win},\nВсего выигранных проектов: {count_projects_win}')
    print(f'WinRate равен: {projects_win_rate} %')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    win_rate()
