import word as w
import excel as x

month = input("Podaj miesiąc, którego ma dotyczyć raport: ")
year = input("Podaj rok, którego ma dotyczyć raport: ")

replacement_dict = {
    '{month}': month,
    '{year}': year,
}

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

word_stationary = w.open_docx("Raport odnośnie urządzeń umieszczonych w terenie.docx") #otwarcie pliku word stacjonarne automaty
word_mobile = w.open_docx("Raport odnośnie urządzeń mobilnych umieszczonych w pociągach.docx") #otwarcie pliku word mobilne automaty

w.replace_text(word_mobile, replacement_dict)

table_stationary = word_stationary.tables[0] #przypisanie tabeli do zmiennej (stacjonarne)
stationary_fail_col = w.find_col_index("Usuwanie nieprawidłowości w funkcjonowaniu urządzeń",table_stationary) #wyszukanie kolumny w której będziemy zmieniać wartości

table_mobile = word_mobile.tables[0] #przypisanie tabeli do zmiennej (mobilne)
mobile_fail_col = w.find_col_index("Usuwanie nieprawidłowości w funkcjonowaniu urządzeń",table_mobile)

excel = x.open_excel("transformed_data.xlsx")
failures_dictionary = x.count_failures(excel) #utworzenie słownika z kluczem automat_id i wartościami failures, count


for automat_id, failures in failures_dictionary.items():
    if 'EN' in automat_id:
        w.insert_failures(automat_id,mobile_fail_col, failures_dictionary,table_mobile, failure_solution_dict)
    #else:
     #   w.insert_failures(automat_id,stationary_fail_col, failures_dictionary,table_stationary, failure_solution_dict)

w.save_docx(f"Raport odnośnie urządzeń mobilnych umieszczonych w pociągach za {month} {year}.docx", word_mobile)
#w.save_docx(f"Raport odnośnie urządzeń umieszczonych w terenie za {month} {year}.docx", word_stationary)





