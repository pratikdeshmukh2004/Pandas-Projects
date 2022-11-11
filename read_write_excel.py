import pandas as pd

# Reading all worksheets from xlsx files.
xl = pd.ExcelFile("data.xlsx")
for sheet in xl.sheet_names:
    print("----------",sheet,"--------")
    dataframe = xl.parse(sheet).to_dict(orient="record")
    for d in dataframe[:50]:
        print(tuple(d.values()))

# Writing xlsx files.
students = pd.DataFrame({'Name': ["Pratik", "Satish"], "Surname": ["Deshmukh", "Mare"]})
xl_writer = pd.ExcelWriter('data.xlsx', engine="xlsxwriter")
status = students.to_excel(xl_writer, sheet_name="Sheet1", index=False)
xl_writer.close()