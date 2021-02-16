# For run instructions, see README.md.

import pandas as pd
import sys

# Define path variable to load appropriate license lists
DATE = sys.argv[1]
PATH = r'C:\Users\elijah.sutton\Documents\OneDrive - Commonwealth of Kentucky\Division of Governance and Strategy\Project Management Branch\License Lists'

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