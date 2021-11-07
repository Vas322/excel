import openpyxl


def get_count_of_elements(list):
    count = 1
    for element in list:
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

    print(f'Проекты: {projects},\nВсего проектов: {get_count_of_elements(projects)}')
    print(f'Отгруженные проекты: {projects_win},\nВсего выигранных проектов: {get_count_of_elements(projects_win)}')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    win_rate()
