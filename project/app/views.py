from django.shortcuts import render
import os
import pandas as pd

def home(request):
    file_path = r'C:\Users\Ashish Mishra\Desktop\Excel_Data_Read from Django\project\app\Sheets.xlsx'
    sheet_name = 'Sheet1'
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    if request.method == 'POST':
        name = request.POST.get('name')
        address = request.POST.get('address')
        mob = request.POST.get('mob')

        # Check if the name already exists in the DataFrame
        if name in df['NAME'].values:
            # Update the existing data
            df.loc[df['NAME'] == name, 'Address'] = address
            df.loc[df['NAME'] == name, 'Mob'] = mob
        else:
            # Add new data if the name doesn't exist
            new_data = pd.DataFrame({'NAME': [name], 'Address': [address], 'Mob': [mob]})
            df = pd.concat([df, new_data], ignore_index=True)

        try:
            df.to_excel(file_path, sheet_name=sheet_name, index=False)
            print("Data updated successfully.")
        except Exception as e:
            print(f"Error writing to Excel file: {e}")

    return render(request, "app/index.html", {"df": df})
