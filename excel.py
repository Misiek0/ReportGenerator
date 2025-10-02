import pandas as pd

def open_excel(excel_file):
    excel = pd.read_excel(excel_file, sheet_name = 'Your Jira Issues', skiprows=1)
    return excel

def count_failures(excel):
    fail_count_dict = {} #inicjalizacja słownika z usterkami automatow
    for _, row in excel.iterrows():
        if row['Status'] != 'Zrealizowane':
            continue

        automat_id = row['Numer Automatu']
        failures = row['Rodzaj usterki'].split(';')

        if automat_id not in fail_count_dict.keys():  # inicjalizacja klucza automatID dla słownika
            fail_count_dict[automat_id] = {}

        for failure in failures:
            if failure not in fail_count_dict[automat_id]:
                fail_count_dict[automat_id][failure] = 0
            fail_count_dict[automat_id][failure] += 1
    return fail_count_dict

def merge_failure_dicts(dict_list):
    merged_dict = {}
    for fail_dict in dict_list:
        for automat_id, failures in fail_dict.items():
            if automat_id not in merged_dict:
                merged_dict[automat_id] = {}
            for failure, count in failures.items():
                if failure not in merged_dict[automat_id]:
                    merged_dict[automat_id][failure] = 0
                merged_dict[automat_id][failure] += count
    return merged_dict
