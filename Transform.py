import pandas as pd


conversion_dict = {
            "secret_data1": "Fault1",
            "secret_data2": "Fault2",
            "secret_data3": "Fault3",
            "secret_data4": "Fault4",
            "secret_data5": "Fault5",
            "secret_data6": "Fault6",
            "secret_data7": "Fault7",
            "secret_data8": "Fault8",
            "secret_data9": "Fault9",
            "secret_data10": "Fault10",
            "secret_data11": "Fault11",
            "secret_data12": "Fault12",
            "secret_data13": "Fault13",
            "secret_data14": "Fault14",
            "secret_data15": "Fault15",
            "secret_data16": "Fault16",
        }

'''this function was used to convert xlsx file 
removing secret data to prepare xlsx file for next report generating processes'''


def transform_confidential_data(data_file, data_output_file):
    excel = pd.read_excel(data_file, sheet_name='Your Jira Issues', skiprows=1)
    #column = excel[["Rodzaj usterki"]]
    new_column = []
    for index, row in excel.iterrows():
        cell = row['Rodzaj usterki']
        failures = cell.split(";")
        mapped = []

        for failure in failures:
            if failure in conversion_dict:
                mapped.append(conversion_dict[failure])
        new_column.append("; ".join(mapped))
    #print(new_column)
    excel['Rodzaj usterki'] = new_column
    excel.to_excel(data_output_file)


transform_confidential_data("data.xlsx", "transformed_data.xlsx")
