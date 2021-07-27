# For run instructions, see README.md.

import pandas as pd
import sys

# Store command-line argument (DATE)
DATE = sys.argv[1]

def clean_o365_license_list(df):
    """
    This function takes in an Office 365 license list (as a DataFrame) and cleans various aspects and then returns the cleaned list.
    
    Parameters
    ----------
    df : pandas DataFrame

    Returns
    -------
    df : pandas DataFrame
    """
    # Cast 'User principal name' (i.e., email address) to all lowercase
    df['User principal name'] = df['User principal name'].str.lower()
    return df

def clean_opm_license_list(df):
    """This function takes the OPM Granted Licenses list (as a DataFrame) as input and cleans various aspects of it before returning it. 

    Parameters
    ----------
    df (DataFrame) : OPM Granted Licenses list

    Returns
    -------
    df (DataFrame) : df after being cleaned.
    """
    # Include only rows where 'Removed' is False
    df = df[ df['Removed'] == False ]

    # Cast 'Email' to all lowercase
    df['Email'] = df['Email'].str.lower()
    return df

def partition_opm_license_list(df):
    """This function partitions the OPM Granted Licenses list by license type, storing the result in a Python dictionary.

    Parameters
    ----------
    df (DataFrame): OPM Granted Licenses list

    Returns
    -------
    license_list_dict (dict): OPM Granted Licenses list partitioned by license type
    """
    license_list_dict = {}
    license_list_dict['PBI'] = df[ ~df['Power BI'].isna() ]
    license_list_dict['P1'] = df[ ~df['Essentials License (Project Plan Essential)'].isna() ]
    license_list_dict['P3'] = df[ ~df['Professional (Project Plan 3)'].isna() ]
    license_list_dict['P5'] = df[ ~df['Premium (Project Plan 5)'].isna() ]
    return license_list_dict

def read_in_users_to_ignore():
    """This function reads in a list of emails to ignore during the comparison and returns the list as a DataFrame.

    Returns
    -------
    df (DataFrame) : Emails to ignore during license list comparison
    """
    df = pd.read_csv('licensed_users_to_ignore.csv', usecols=['User principal name'])
    df['email'] = df['User principal name'].str.lower()
    return df

def drop_users_to_ignore(ignore, license_lists):
    """This function drops the users to ignore during the comparison from each license type list.

    Parameters
    ----------
    ignore (DataFrame) : Users to ignore during comparison
    license_lists (dict) : dictionary of DataFrames, one for each license type 

    Returns
    -------
    license_lists (dict) : same as input minus DataFrame records whose email matched an email in the *ignore* DataFrame
    """
    for license in license_lists.keys():
        license_lists[license] = license_lists[license][ ~license_lists[license]['User principal name'].isin(ignore['email']) ]
    return license_lists

def read_in_o365_licenses(*licenses):
    """ 
    This function takes a variable number of license types as input and reads in a 
    corresponding license list into a dictionary of DataFrames.
    
    Parameters
    ------
    *licenses : variable argument 

    Returns
    -------
    license_lists_dict (dict) : Python dictionary of pandas DataFrames
    """
    license_lists_dict = dict()
    for license in licenses:
        license_lists_dict[license] = pd.read_excel(r'{}\{}.xlsx'.format(DATE, license), engine='openpyxl')
        license_lists_dict[license] = clean_o365_license_list(license_lists_dict[license])
    return license_lists_dict

def read_in_opm_license_list():
    """This function reads in the OPM Granted Licenses list from an Excel Workbook and returns a dictionary of license lists with license types as keys.

    Returns
    -------
    license_list_dict (dict) : OPM Granted Licenses list partitioned by license type (and the keys are the license types matching the keys of O365's license list dict)
    """
    fname = 'PWA Licenses Tracker.xlsx'
    df = pd.read_excel(fname, sheet_name='Granted Licenses', engine='openpyxl')
    df = clean_opm_license_list(df)
    license_list_dict = partition_opm_license_list(df)
    return license_list_dict

def compare_license_lists(o365_dict, opm_dict):
    """This function computes a symmetric difference for each license list type and writes the results to Excel files

    Parameters
    ----------
    o365_dict (dictionary): O365's license lists dictionary
    opm_dict (dictionary): OPM's license lists dictionary
    """
    # Open text file to store comparison summary
    f = open(r'{}\comparison_summary.txt'.format(DATE), 'w')

    for k in o365_dict.keys():
        # Store license type DataFrame ina variable (for readability)
        o365_df = o365_dict[k]
        opm_df = opm_dict[k]
        
        # Compute symmetric difference        
        res = o365_df.merge(opm_df, 
                            how='outer', 
                            left_on='User principal name', 
                            right_on='Email', 
                            suffixes=('_o365', '_opm'), 
                            indicator='source'
                            )[['Department', 'Display name', 'First name_o365', 'Last name_o365', 'Licenses', 'User principal name', 
                            'When created', 'Cabinet', 'Last name_opm', 'First name_opm', 'Email', 'Essentials License (Project Plan Essential)',
                            'Professional (Project Plan 3)', 'Power BI', 'Premium (Project Plan 5)', 'Owner', 'Issue Date', 'source']]
        symmetric_diff_df = res.loc[res['source'] != 'both']

        # Make clear which license list the difference is from
        symmetric_diff_df.replace({'source': {'left_only': 'o365', 'right_only': 'opm'}}, inplace=True)

        # Write summary of comparison to file object for comparison_summary
        o365_minus_opm_len = symmetric_diff_df[symmetric_diff_df['source'] == 'o365'].shape[0]
        opm_minus_o365_len = symmetric_diff_df[symmetric_diff_df['source'] == 'opm'].shape[0]
        str_2_write = '{}:\n\t O365 - OPM = {}\n\t OPM - O365 = {}\n'.format(k, o365_minus_opm_len, opm_minus_o365_len)
        f.write(str_2_write)

        # Write symmetric_diff_df to Excel file
        symmetric_diff_df.to_excel(r'{}\{}_diffs.xlsx'.format(DATE, k), index=False)
    
    # Close file object for comparison_summary
    f.close()

    return

def main():
    # Read in license lists provided by O365 Team (o365_license_lists)
    o365_license_lists = read_in_o365_licenses('P1', 'P3', 'P5', 'PBI')

    # Read in list of users to ignore during comparison
    users_to_ignore = read_in_users_to_ignore()

    # Drop license users in users_to_ignore from DataFrames in o365_license_lists
    drop_users_to_ignore(users_to_ignore, o365_license_lists)

    # Read OPM's Granted Licenses list
    opm_license_lists = read_in_opm_license_list()

    # Compare license lists: O365 Team's versus OPM's
    compare_license_lists(o365_license_lists, opm_license_lists)

    return

main()