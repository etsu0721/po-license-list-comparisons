"""
Created on Mon Nov  2 10:02:51 2020
@author: elijah.sutton

For run instructions, see README.md.
"""

import pandas as pd
import sys

# Define path variable to load appropriate license lists
DATE = sys.argv[1]
PATH = r'C:\Users\elijah.sutton\Documents\OneDrive - Commonwealth of Kentucky\Division of Governance and Strategy\Project Management Branch\License Lists'

def load_Billie_license_list(file_name):
    """ 
    This function loads an Excel file into a pandas DataFrame. The Excel file
    name is specified as a function parameter, and a path defined globally is used.
    
    Parameters
    ------
    file_name : Name of an Excel file (string)

    Returns
    -------
    df : pandas DataFrame
    """
    df = pd.read_excel(r'{}\{}\{}'.format(PATH, DATE, file_name), engine='openpyxl')
    return df

def subset_tracker_data_by_license(license_type):
    """
    This function subsets a pandas DataFrame based on the license type 
    argument. The DataFrame is also subset to include only records where the
    'Removed' column is False.

    Parameters
    ----------
    license_type : string

    Returns
    -------
    df : pandas DataFrame

    """
    condition_1 = ~tracker_df[license_type].isna()
    condition_2 = tracker_df['Removed'] == False
    df = tracker_df[(condition_1 & (condition_2))]
    return df

def clean_user_principal_name(df):
    """
    This function cleans the 'User principal name' column by making all
    characters lowercase.

    Parameters
    ----------
    df : TYPE
        DESCRIPTION.

    Returns
    -------
    df : TYPE
        DESCRIPTION.

    """
    df['User principal name'] = df['User principal name'].str.lower()
    return df

def find_differences(df1, df2, license_type, df1_name="O365 list", df2_name="OPM list"):
    """
    This function takes two pandas Series, one containing licensed emails per O365's
    list and the other containing licensed emails per OPM's list, a license type, 
    and names for each of the series (which have default values) and produces a 
    report of the differences between the two Series. The differences found are
    written to CSV files with names identifying the license type and the order 
    the difference was taken.
    
    Parameters
    ----------
    s1 : pandas DataFrame
    s2 : pandas DataFrame
    license_type : string
    df1_name : string, optional. The default is "O365 list".
    df2_name : string, optional. The default is "OPM list".

    Returns
    -------
    None.

    """
    s1 = df1['User principal name']
    s2 = df2['Email']
    
    print('Comparison of {} licenses:'.format(license_type))
    
    # Display count of total UPNs/Emails
    print('\t{} has {} items.'.format(df1_name, len(s1)))
    print('\t{} has {} items.'.format(df2_name, len(s2)))
    print()
    
    # Display count of unique UPNs/Emails
    print('\t{} has {} unique items.'.format(df1_name, len(s1.unique())))
    print('\t{} has {} unique items.'.format(df2_name, len(s2.unique())))
    print()
    
    # Display summary of set differences 
    
    # series 1 - series 2
    s1_minus_s2 = set(s1).difference(set(s2))
    if len(s1_minus_s2) == 0:
        print('\t No differences.')
    else:
        print('\t{} minus {} differences: {}'.format(df1_name, df2_name, len(s1_minus_s2)))
        
    # series 2 - series 1
    s2_minus_s1 = set(s2).difference(set(s1))
    if len(s2_minus_s1) == 0:
        print('\t No differences.')
    else:
        print('\t{} minus {} differences: {}'.format(df2_name, df1_name, len(s2_minus_s1)))
    print()
    
    # Write differences to CSV files (and include all data elements of the DataFrame)
    df1[ df1['User principal name'].isin(s1_minus_s2) ].sort_values(by='Last name').to_csv(r'{}\{}\{}-O365-minus-OPM.csv'.format(PATH, DATE, license_type), index=False)
    df2[ df2['Email'].isin(s2_minus_s1) ].sort_values(by='Last Name').to_csv(r'{}\{}\{}-OPM-minus-O365.csv'.format(PATH, DATE, license_type), index=False)
    #pd.Series(list(s1_minus_s2)).drop(labels=['', ' '], errors='ignore').sort_values().to_csv(r'{}\{}\{}-O365-minus-OPM.csv'.format(PATH, DATE, license_type), header=['Licensed Email'], index=False)
    #pd.Series(list(s2_minus_s1)).drop(labels=['', ' '], errors='ignore').sort_values().to_csv(r'{}\{}\{}-OPM-minus-O365.csv'.format(PATH, DATE, license_type), header=['Licensed Email'], index=False)

    return

# Read in each of the license lists
pbi_pro = load_Billie_license_list('PBI.xlsx')
po_essential = load_Billie_license_list('P1.xlsx')
po_pro = load_Billie_license_list('P3.xlsx')
po_premium = load_Billie_license_list('P5.xlsx')

# Clean 'User principal name' column in each Datafram
pbi_pro = clean_user_principal_name(pbi_pro)
po_essential = clean_user_principal_name(po_essential)
po_pro = clean_user_principal_name(po_pro)
po_premium = clean_user_principal_name(po_premium)

# Load Granted Licenses sheet from PWA License Tracker
pwa_tracker_fname = 'PWA Licenses Tracker.xlsx'
tracker_df = pd.read_excel(r'{}\{}'.format(PATH, pwa_tracker_fname), sheet_name='Granted Licenses', engine='openpyxl')

# Clean 'Email' column by making lowercase.
tracker_df['Email'] = tracker_df['Email'].str.lower()

# Subset tracker data by license type
tracker_pbi_pro = subset_tracker_data_by_license('Power BI')
tracker_po_essential = subset_tracker_data_by_license('Essentials License (Project Plan Essential)')
tracker_po_pro = subset_tracker_data_by_license('Professional (Project Plan 3)')
tracker_po_premium = subset_tracker_data_by_license('Premium (Project Plan 5)')

# Find differences between O365 lists and OPM lists (and vice versa)
find_differences(po_essential, tracker_po_essential, 'P1')
find_differences(po_pro, tracker_po_pro, 'P3')
find_differences(po_premium, tracker_po_premium, 'P5')
find_differences(pbi_pro, tracker_pbi_pro, 'PBI')