#import libraries
import pandas as pd
import numpy as np
from datetime import timedelta, time
from datetime import datetime, timedelta


#import data file paths named after departments
production=r"C:\Users\Admin\bkg.nl(1)\OLT-Power BI - Human Resource - Human Resource\Evalyne Lenku's files - actual activities\Actual Activities\Production Actual Activities.xlsx"
groups_file_path=r"C:\Users\Admin\bkg.nl(1)\OLT-Power BI - Documents\General\Speed file\New quality check 2.xlsx"
plumbing= r"C:\Users\Admin\bkg.nl(1)\OLT-Power BI - Human Resource - Human Resource\Evalyne Lenku's files - actual activities\Actual Activities\Plumbing Actual Activities.xlsx"
technical=r"C:\Users\Admin\bkg.nl(1)\OLT-Power BI - Human Resource - Human Resource\Evalyne Lenku's files - actual activities\Actual Activities\Technical Actual Activities.xlsx"
electricals=r"C:\Users\Admin\bkg.nl(1)\OLT-Power BI - Human Resource - Human Resource\Evalyne Lenku's files - actual activities\Actual Activities\Electricals Actual Activities.xlsx"


production=pd.read_excel(production, sheet_name='Activities')
groups=pd.read_excel(groups_file_path, sheet_name='Groups')
plumbing=pd.read_excel(plumbing, sheet_name= "Activities")
technical=pd.read_excel(technical, sheet_name="Activities")
electricals=pd.read_excel(electricals, sheet_name="Activities")



#function to clean data
def clean_data(df):
    # Strip whitespace from the whole column first
    df['Actual Time'] = df['Actual Time'].astype(str)
    df["Actual Time"] = df["Actual Time"].str.replace(";", ":", regex=False)


    #change datatype to string
    df['Week']=df['Week'].astype(str)
    df['Day']=df['Day'].astype(str).str.strip()
    df['Actual Time']=df['Actual Time'].astype(str).str.strip()


    #first convert 'nan' strings to actual NaN values
    df['Week'].replace('nan', pd.NA, inplace=True)
    df['Day'].replace('nan', pd.NA, inplace=True)
    df['Year'].replace('nan', pd.NA, inplace=True)

    #convert empty strings to NaN
    df['Week'].replace('', pd.NA, inplace=True)
    df['Day'].replace('', pd.NA, inplace=True)

    #fill null values
    df['Week'].fillna(method='ffill', inplace=True)
    df['Day'].fillna(method='ffill', inplace =True)
    df['Year'].fillna(method='ffill', inplace =True)

    #convert back to int
    df['Week']=df['Week'].astype(float).astype(int)
    df['Day']=df['Day'].astype(float).astype(int)


    #delete row where actual time is missing. because that mean the activity wasnt done because there was no time recorded
    df = df.dropna(subset=["Actual Time"])          # removes NaN
    df = df[df["Actual Time"].str.strip() != ""]  # removes empty strings
    

    #split start and stop time
    #df[['Start_Time', 'Stop_Time']] = df['Actual Time'].str.split('-', expand=True)
    split_times = df['Actual Time'].str.split('-', n=1, expand=True)
    df['Start_Time'] = split_times[0]
    df['Stop_Time']  = split_times[1]

    df['Start_Time'] = df['Start_Time'].str.strip()
    df['Stop_Time']  = df['Stop_Time'].str.strip()


    df['Start_Time'] = pd.to_datetime(df['Start_Time'], format='%H:%M', errors='coerce')
    df['Stop_Time'] = pd.to_datetime(df['Stop_Time'], format='%H:%M', errors='coerce')

    # If Stop_Time is earlier than Start_Time, assume it's PM and add 12 hours
    df.loc[df['Stop_Time'].dt.hour < df['Start_Time'].dt.hour, 'Stop_Time'] += pd.Timedelta(hours=12)

    return df


#convertime to 12hours: anytime between 1pm and 4pm is converted to 12hrs
def adjust_pm(t):
    if 1 <= t.hour <= 6:
        return t + timedelta(hours=12)
    return t


def find_time_taken(df):  
    df['Start_Time'] = df['Start_Time'].apply(adjust_pm)
    df['Stop_Time'] = df['Stop_Time'].apply(adjust_pm)

    #subtract to find time spent on activity
    df['Time_Taken']=df['Stop_Time']-df['Start_Time']

    #remove dates from the start and stop time column so it is left with time only
    df['Start_Time'] = df['Start_Time'].dt.strftime('%H:%M')
    df['Stop_Time'] = df['Stop_Time'].dt.strftime('%H:%M')
    #df['Time_taken']=df['Time_taken'].dt.strftime('%H:%M')

    df['Time_Taken'].fillna(pd.Timedelta(0), inplace=True)
    df['Time_Taken']=df['Time_Taken'].astype(str).str.split(' ', expand=True)[2]


    # Convert Start_Time and Stop_Time to time objects
    df['Start_Time'] = pd.to_datetime(df['Start_Time'], format='%H:%M').dt.time
    df['Stop_Time'] = pd.to_datetime(df['Stop_Time'], format='%H:%M').dt.time
    return df



#apply duration adjustment rules
def duration_adjust(row):
    start = row['Start_Time']
    stop = row['Stop_Time']

    if pd.isna(start) or pd.isna(stop):
        return timedelta(0)

    # Rule 1: No deduction if start is 2 PM or later
    if start >= time(14, 0):
        return timedelta(0)
    
    # Rule 1b: No deduction if start is between 1 PM and 2 PM
    if time(13, 0) <= start < time(14, 0):
        return timedelta(0)

    # Rule 1: No deduction if both times are between 11 AM and 1 PM
    if time(11, 0) <= start <= time(13, 0) and time(11, 0) <= stop <= time(13, 0):
        return timedelta(0)

    # Apply deductions
    deduction = timedelta(0)

    if stop > time(11, 0):
        deduction += timedelta(minutes=30)

    if stop > time(14, 0):
        deduction += timedelta(hours=1)

    return deduction



#remove days from duration adjustment
def clean_time(df):
    df['Duration_Adjust']=df['Duration_Adjust'].astype(str).str.split(' ', expand=True)[2]
    df['Net_Duration']=pd.to_timedelta(df['Time_Taken']) - pd.to_timedelta(df['Duration_Adjust'])
    df['Net_Duration']=df['Net_Duration'].astype(str).str.split(' ', expand=True)[2]

    #remove all the white spaces in the dataframe
    df=df.applymap(lambda x:x.strip() if  isinstance(x, str) else x)

    #drop actual time column because its problematic
    df.drop(columns=['Actual Time'], inplace =True)

    # Convert Start_Time and Stop_Time to time objects
    df['Start_Time'] = df['Start_Time'].apply(lambda x: x.strftime('%H:%M:%S') if pd.notnull(x) else '')
    df['Stop_Time'] = df['Stop_Time'].apply(lambda x: x.strftime('%H:%M:%S') if pd.notnull(x) else '')
    return df




#apply the functions
production=clean_data(production)
production=find_time_taken(production)
# Apply adjustment function to each row
production['Duration_Adjust'] = production.apply(duration_adjust, axis=1)
production=clean_time(production)



print(production.info())
print(production.head())


#find number of staff by group
production['Group'].fillna(method='ffill', inplace=True)
groups = groups[['HARVEST NO', 'GRP']]
groups.drop(
    columns=[col for col in groups.columns if col not in ['HARVEST NO', 'GRP']],
    inplace=True
)
Group_size = groups.groupby('GRP')['HARVEST NO'].nunique().reset_index()
print(Group_size)

#marge df and and group size
production = pd.merge(production, Group_size, left_on='Group', right_on='GRP', how='left')
production=production.rename(columns={'HARVEST NO': 'Group_Size'})
print(production)

from datetime import datetime, timedelta

def date_from_week(year, week, day):
    # ISO defines week 1 as the week with Jan 4th
    jan4 = datetime(year, 1, 4)
    week1_monday = jan4 - timedelta(days=jan4.isoweekday() - 1)
    # day=1 means Monday, day=7 means Sunday
    return week1_monday + timedelta(weeks=week-1, days=day-1)


print(date_from_week(2025, 47, 3))  # week 10, day 3 (Wednesday)
production['Date'] = production.apply(
    lambda row: date_from_week(int(row['Year']), int(row['Week']), int(row['Day'])),
    axis=1
)
print(production)
print(production.info())


#cleaning and Technical and plumbing
Technical=pd.concat([plumbing, technical, electricals], ignore_index=True, axis =0)
print(Technical.info())


#apply the functions
Technical=clean_data(Technical)
Technical=find_time_taken(Technical)
# Apply adjustment function to each row
Technical['Duration_Adjust'] = Technical.apply(duration_adjust, axis=1)
Technical=clean_time(Technical)
Technical['Date'] = production.apply(
    lambda row: date_from_week(int(row['Year']), int(row['Week']), int(row['Day'])),
    axis=1
)


Technical.to_csv(r"C:\Users\Admin\bkg.nl(1)\OLT-Power BI - Human Resource - Human Resource\Evalyne Lenku's files - actual activities\Clean Data\clean_data_technical & Plumbing.csv", index=False)
production.to_csv(r"C:\Users\Admin\bkg.nl(1)\OLT-Power BI - Human Resource - Human Resource\Evalyne Lenku's files - actual activities\Clean Data\clean_data Production.csv", index=False)
