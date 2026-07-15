import pandas as pd

# تجهيز البيانات الجاهزة للتجربة
data = {
    'CompanyName': ['Tech Solutions', 'DataCorp', 'Global Systems'],
    'Email': ['test1@domain.com', 'test2@domain.com', 'test3@domain.com'],
    'Notes': [
        'Looking for a .NET Backend Developer with SQL Server expertise',
        'Needs a Full-Stack developer experienced in Angular and Python',
        'Urgent need for a Database Administrator with Oracle and PL/SQL experience'
    ]
}

# تحويل البيانات إلى dataframe
df = pd.DataFrame(data)

# حفظ الملف كـ Excel بصيغة صحيحة
file_name = "companies.xlsx"
df.to_excel(file_name, index=False)

print(f"🎉 Success! '{file_name}' has been created with ready-to-use data.")