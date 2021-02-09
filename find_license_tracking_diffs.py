# For run instructions, see README.md.

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

def clean_o365_license_list(df):
    """
    This function takes in an Office 365 license list (as a DataFrame) and cleans various aspects and then returns the cleaned list.
    
    Parameters
    ------
    df :  pandas DataFrame

    Returns
    -------
    df : pandas DataFrame
    """
    # Cast 'User principal name' (i.e., email address) to all lowercase
    df['User principal name'] = df['User principal name'].str.lower()
    return df

def read_in_o365_licenses(*licenses):
    """ 
    This function takes a variable number of license types as input and reads in a 
    corresponding license list into a dictionary of DataFrames.
    
    Parameters
    ------
    *licenses : variable argument 

    Returns
    -------
    license_lists_dict : Python dictionary of pandas DataFrames
    """
    license_lists_dict = dict()
    for license in licenses:
        license_lists_dict[license] = pd.read_excel(r'{}\{}\{}.xlsx'.format(PATH, DATE, license), engine='openpyxl')
        license_lists_dict[license] = clean_o365_license_list(license_lists_dict[license])
    return license_lists_dict

def clean_opm_license_list(df):
    """This function takes the OPM Granted Licenses list (as a DataFrame) as input and cleans various aspects of it before returning it. 

    Args:
        df (DataFrame): OPM Granted Licenses list

    Returns:
        DataFrame: df after being cleaned.
    """
    # Include only rows where 'Removed' is False
    df = df[ df['Removed'] == False ]

    # Cast 'Email' to all lowercase
    df['Email'] = df['Email'].str.lower()
    return df

def partition_opm_license_list(df):
    """This function partitions the OPM Granted Licenses list by license type, storing the result in a Python dictionary.

    Args:
        df (DataFrame): OPM Granted Licenses list

    Returns:
        dictionary: OPM Granted Licenses list partitioned by license type
    """
    license_list_dict = dict()
    license_list_dict['PBI'] = df[ ~df['Power BI'].isna() ]
    license_list_dict['P1'] = df[ ~df['Essentials License (Project Plan Essential)'].isna() ]
    license_list_dict['P3'] = df[ ~df['Professional (Project Plan 3)'].isna() ]
    license_list_dict['P5'] = df[ ~df['Premium (Project Plan 5)'].isna() ]
    return license_list_dict

def read_in_opm_license_list():
    """This function reads in the OPM Granted Licenses list from an Excel Workbook and returns a dictionary of license lists with license types as keys.

    Returns:
        dictionary: OPM Granted Licenses list partitioned by license type (and the keys are the license types matching the keys of O365's license list dict)
    """
    fname = 'PWA Licenses Tracker.xlsx'
    df = pd.read_excel(r'{}\{}'.format(PATH, fname), sheet_name='Granted Licenses', engine='openpyxl')
    df = clean_opm_license_list(df)
    license_list_dict = partition_opm_license_list(df)
    return license_list_dict

def compare_license_lists(dict1, dict2, dict1_name='O365 list', dict2_name='OPM list'):
    """This function generates a summary of the comparison of license lits for each license type and displays it in the terminal.
    This function also writes the differences (in both directions) for each license type to separate CSV files.

    Args:
        dict1 (dictionary): O365's license lists dictionary
        dict2 (dictionary): OPM's license lists dictionary
        dict1_name (str, optional): A name to reference to dict1 in the output. Defaults to 'O365 list'.
        dict2_name (str, optional): A name to reference to dict2 in the output. Defaults to 'OPM list'.
    """
    for k in dict1.keys():
        o365_df = dict1[k]
        opm_df = dict2[k]
        s1 = dict1[k]['User principal name']
        s2 = dict2[k]['Email']

        print('Comparison of {} licenses:'.format(k))

        # Display count of licenses of type k
        print('\t{} has {} items.'.format(dict1_name, len(o365_df)))
        print('\t{} has {} items.'.format(dict2_name, len(opm_df)))
        print()

        # Display count of unique UPNs/Emails
        print('\t{} has {} unique items.'.format(dict1_name, len(s1.unique())))
        print('\t{} has {} unique items.'.format(dict2_name, len(s2.unique())))
        print()

        # Display summary of set differences 
    
        # O365 minus OPM
        s1_minus_s2 = set(s1).difference(set(s2))
        if len(s1_minus_s2) == 0:
            print('\tNo differences.')
        else:
            print('\t{} minus {} differences: {}'.format(dict1_name, dict2_name, len(s1_minus_s2)))
            
        # OPM minus O365
        s2_minus_s1 = set(s2).difference(set(s1))
        if len(s2_minus_s1) == 0:
            print('\tNo differences.')
        else:
            print('\t{} minus {} differences: {}'.format(dict2_name, dict1_name, len(s2_minus_s1)))
        print()
        
        # Write differences to CSV files (and include all data elements of the DataFrame)
        o365_df[ o365_df['User principal name'].isin(s1_minus_s2) ].sort_values(by='Last name').to_csv(r'{}\{}\{}-O365-minus-OPM.csv'.format(PATH, DATE, k), index=False)
        opm_df[ opm_df['Email'].isin(s2_minus_s1) ].sort_values(by='Last Name').to_csv(r'{}\{}\{}-OPM-minus-O365.csv'.format(PATH, DATE, k), index=False)
    return

def main():
    # Read in license lists provided by O365 Team (o365_license_lists)
    o365_license_lists = read_in_o365_licenses('P1', 'P3', 'P5', 'PBI')

    # Read OPM's Granted Licenses list
    opm_license_lists = read_in_opm_license_list()

    # Find differences between the O365 Team's and OPM's license lists
    compare_license_lists(o365_license_lists, opm_license_lists)
    return

main()