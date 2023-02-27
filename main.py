import pandas as pd
import datetime
from fastapi import FastAPI, UploadFile,File
import calendar

app = FastAPI()

@app.post("/timesheet")
def generate_timesheet(file: UploadFile = File(...)):
    with open(file.filename, "wb") as f:
        f.write(file.file.read())
    df = pd.read_excel(file.filename, sheet_name='in', keep_default_na=False, parse_dates=True)
    billable_data = df[df["Billing Action"] == "Billable"]
    billable_data["Date"] = pd.to_datetime(billable_data["Date"], format="%Y-%m-%d")

    # get the year, month, date
    year = billable_data["Date"].dt.year.unique()[0]
    month = billable_data["Date"].dt.month.unique()[0]
    num_days = calendar.monthrange(year, month)[1]

    # create a column list format
    dates = []
    for day in range(1, num_days+1):
       date = datetime.datetime(year, month, day)
       if date.strftime("%b-%d") not in dates:
            dates.append(date.strftime("%b-%d"))
    
    # weekend for to give css
    weekend_dates = []
    for day in range(1, num_days + 1):
        if datetime.datetime(year, month, day).weekday() in [5, 6]:
            weekend_dates.append(datetime.datetime(year, month, day).strftime("%b-%d"))
    
    # unique values of employee
    unique_employee_ids = billable_data["Employee ID"].unique()
    
    # Create a new dataframe 
    result = pd.DataFrame(columns=["Sl No",
                                   "Project",
                                   "Project Name",
                                   "Employee ID", 
                                   "Name", 
                                   "Location", 
                                   "Total_Worked_Days",
                                   "Total_billable_hours"] + dates)
    
    # Loop through each unique "Employee ID"
    for i, employee_id in enumerate(unique_employee_ids):
        employee_data = billable_data[billable_data["Employee ID"] == employee_id]
        num_worked_days = len(employee_data["Date"].unique())
        total_hours=employee_data["Time Quantity"].sum()
        
        # content in the row
        first_row = employee_data.iloc[0]
        result_row = {"Sl No": i+1,
                      "Project": first_row["Project"],
                      "Project Name": "AGERO FULL STACK DEV T&M",
                      "Employee ID": first_row["Employee ID"],
                      "Name": first_row["Name"],
                      "Location": first_row["ON / OF"],
                      "Total_Worked_Days": num_worked_days,
                      "Total_billable_hours":total_hours}
      
        # Add the result row to the result dataframe
        result = result.append(result_row, ignore_index=True)
        
        # date fetch time quantity
        for date in dates:
            grouped_data = employee_data.groupby(["Date", "Employee ID"])["Time Quantity"].sum().reset_index()
            date_data = grouped_data[grouped_data["Date"].dt.strftime("%b-%d") == date]
            if not date_data.empty:
                time_quantity = date_data["Time Quantity"].iloc[0]
                result.loc[result["Employee ID"] == employee_id, date] = time_quantity
            else:
                result_row[date] = " "
    
    # CSS -START
    # Function to apply background color to cells
    def color_background(value):
       if value==9:
        color = 'None'
       elif value == 4.5:
          color = 'yellow'
       elif value==8:
          color='None'
       elif value==10:
          color='None'
       else:
          color = 'red'
       return f'background-color: {color}'   
    styler = result.style.set_properties(**{'text-align': 'center'}).applymap(color_background, subset=pd.IndexSlice[:, dates])
    for date in weekend_dates:
        styler = styler.set_properties(**{'background-color': 'None'}, subset=date)

        
    #  result dataframe to a new excel file
    styler.to_excel("time.xlsx", index=False)
    return ("Timesheet has been created successfully")

