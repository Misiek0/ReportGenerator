import word as w
import excel as x

failure_solution_dict = {
    'fault1': 'solution1',
    'fault2': 'solution2',
    'fault3': 'solution3',
    'fault4': 'solution4',
    'fault5': 'solution5',
    'fault6': 'solution6',
    'fault7': 'solution7',
    'fault8': 'solution8',
    'fault9': 'solution9',
    'fault10': 'solution10',
    'fault11': 'solution11',
    'fault12': 'solution12',
    'fault13': 'solution13',
    'fault14': 'solution14',
    'fault15': 'solution15',
    'fault16': 'solution16'
}

word_stationary = w.open_docx("Raport odnośnie urządzeń umieszczonych w terenie.docx")
word_mobile = w.open_docx("Raport odnośnie urządzeń mobilnych umieszczonych w pociągach.docx")

table_stationary = word_stationary.tables[0]
stationary_fail_col = w.find_col_index("Usuwanie nieprawidłowości w funkcjonowaniu urządzeń",table_stationary)



table_mobile = word_mobile.tables[0]
mobile_fail_col = w.find_col_index("Usuwanie nieprawidłowości w funkcjonowaniu urządzeń",table_mobile)

excel = x.open_excel("transformed_data.xlsx")

failures_dictionary = x.count_failures(excel)





