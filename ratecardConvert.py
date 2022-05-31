from asyncio.windows_events import NULL
from pickle import NONE
import sys
import glob, os
import re

import pandas as pd

"""
TODO: 
1. Create a comparison tool
"""

# amount of discounts we need to run through for business - Change this if there are ever more or less discounts available
# This deals with 9 total tariffs (no discount and amount of discounts stated below)
DISCOUNTS = 8

input_folder = "\\Rate_Cards\\"
output_folder = "Output_Rate_Cards\\"
static_folder = "\\Static\\"

# TO DO: include a comparison tool

def main():
    files = []


    # ensures the correct directories exist
    os.makedirs(os.getcwd() + input_folder, exist_ok=True)
    os.makedirs(os.getcwd() + output_folder, exist_ok=True)
    os.makedirs(os.getcwd() + static_folder, exist_ok=True)

    # ensure files are in folder
    print(
        "Please ensure that all Rate Cards that need to be converted are in the Rate_Cards folder and you have included the static ratecard in the Static folder before you continue!")
    run = input("Please enter Y to continue: ")

    # user must confirm to continue
    if run.upper() != "Y":
        sys.exit("User not confirmed.")

    # stores list of files in folder also changes current directory down a level
    os.chdir(os.getcwd() + input_folder)
    for file in glob.glob("*.xlsx"):
        files.append(file)

    # set the current directory back to the origin
    os.chdir(os.path.dirname(os.getcwd()))

    # ensures a min of 1 and max of 2 files in folder
    if not files and len(files) < 3:
        sys.exit("There are currently " + str(len(files)) + " files in the folder. Expecting Minimum of 1 and Maximum of 2")

    # prepares any business rate cards
    business_rates = prepare_business(files)

    # we now are going to run this script until all files have been processed
    runtime = len(files)
    print(files)
    while runtime != 0:
        process_rate_card(runtime, files, business_rates)
        runtime -= 1
    return


def prepare_business(files):
    """
    This function checks whether there is a business rate card to add. 
    It will process it ready to be concatenated later in the program. 
    """
    # checks files for business rate card
    pattern = r'(?i)(.*Business.*)'
    for file in files:
        if re.search(pattern, file):
            disc_tot = NONE
            # sets up pattern to check for franchise in sheet
            pattern = r'(?i)(.*Franchise.*)'
            sheets = pd.ExcelFile(os.getcwd() + input_folder + file, engine='openpyxl')
            for sheet in sheets.sheet_names:
                if re.search(pattern, sheet):
                    df = sheets.parse(sheet, parse_dates=True, skiprows=1)
                    break
                else:
                    #if sheet doesn't have the word 'Franchise' we move to the next one.
                    continue
        else:
            #if File doesn't have the word 'BIR' we move to the next one.
            continue

        try:        
            # deals with the first non discounted sku's
            df = df.rename(columns={'Legacy Code': 'SKU', 'Price Plan': 'DESCRIPTION', 'Product Type': 'TYPE'})
            disc = df.copy()
            disc = disc.rename(columns={'No Discount' : 'REVENUE'})
            disc = pd.DataFrame(disc,columns=['TYPE', 'SKU', 'DESCRIPTION', 'REVENUE'])
            disc['SKU'] = disc['SKU'].apply(lambda x: str(x) +"ND")
            disc_tot = disc.copy()

            # deals with all other discount SKU's (set at the beginning of the program)
            for discount in range(DISCOUNTS):
                name = f"Discount {discount + 1}"
                sku_type = f"L{discount + 1}"
                disc = df.copy()
                disc = disc.rename(columns={name : 'REVENUE'})
                disc = pd.DataFrame(disc,columns=['TYPE', 'SKU', 'DESCRIPTION', 'REVENUE'])
                disc['SKU'] = disc['SKU'].apply(lambda x: str(x) + sku_type)
                disc_tot = pd.concat([disc_tot, disc], ignore_index=True)

            # creates a commission column for business tariffs and sets commission at 5%
            disc_tot['COMMISSION'] = disc_tot['REVENUE'].apply(lambda x: x *.05)
            return disc_tot
        except:
            print("No Business Rate Card Present")
            return

def process_rate_card(filenumber, files, business_rates):
    """
    This part of the program deals with the converting and 
    cleaning the data to make it usable and then exports as
    new .CSV file
    """

    # adds excel file and static files to dataframe and ensures file is readable.
    try:
        df = pd.read_excel(os.getcwd() + input_folder + files[filenumber - 1], engine='openpyxl')
    except:
        print(os.getcwd() + input_folder)
        print("There was an error opening the file " + files[filenumber - 1] + ". Please ensure you have closed the file so it can be read.")
        retry = input("Press any key to exit: ")
        sys.exit("Read File Error.")
    try:
        static_df = pd.read_csv(os.getcwd() + static_folder + "static_rates.csv")
    except:
        print("There was an error opening the file static_rates.csv. Please ensure you have closed the file so it can be read.")
        retry = input("Press any key to exit: ")
        sys.exit("Read File Error.")

    # skips any files that don't fit a format
    if "TYPE" in df and "SOC Code" in df and "Description" in df and "Acq_Ret" in df and "MAF (Inc VAT)" in df:
        if "Webchat" in df:
            print(files[filenumber - 1] + " file present")
        elif "Leeds White Rose" in df and "Leeds 8-9 Commercial Street" in df and "Castleford" in df:
            print(files[filenumber - 1] + " file present")
        else:
            print(files[filenumber - 1] + " incorrect format")
            return
    else:
        print(files[filenumber - 1] + " incorrect format")
        return

    # rename and drop apropriate columns
    df = df.rename(columns={'SOC Code': 'SKU', 'Description': 'DESCRIPTION'})
    df['SKU'] = df.apply(lambda x: str(x['SKU'])+"N" if x['Acq_Ret'] ==
                        'Acquisition' and x['TYPE'] != 'TALK MOBILE' else x['SKU'], axis=1)
    df.dropna(subset=['TYPE'], inplace=True)

    # converts the Monthly fee to an integer that will be used to create the new SKU
    df['MAF (Inc VAT)'] = df['MAF (Inc VAT)'].apply(lambda x: x *100).map('{0:g}'.format)
    
    # we itterate through all talk mobile deals and convert to the format TM + contract length + deal price
    # eg for a 1 month contract at 7.95 = TM300795, 12 months at 11.50 = TM121150.
    for index, row in df.iterrows():
        if row['TYPE'] == 'TALK MOBILE':
            if int(row['MAF (Inc VAT)']) < 1000:
                row['MAF (Inc VAT)'] = "0" + str(row['MAF (Inc VAT)'])
            if row['Contract Length (Months)'] == 1:
                row['Contract Length (Months)'] = 30
            # applies changed to SKU
            df.loc[index, "SKU"] = "TM" + str(row['Contract Length (Months)']) + str(row['MAF (Inc VAT)'])
    # removes columns we no longer need
    df.drop(['MAF (Inc VAT)','Contract Length (Months)'], axis=1, inplace=True)
    # df = df.drop_duplicates(subset=['SKU'], keep=False)
    try:
        df.drop(['Legacy Code - EBU', 'Band'], axis=1, inplace=True)
    except:
        print("Legacy Code - EBU or Band not present")

    # establish if webchat ratecard or retail
    if 'Webchat' in df.columns:
        # renames the columns to match what the My Rate Card Database expects
        df = df.rename(columns={'Webchat': 'REVENUE'})
        # Drops columns we no longer need
        df.drop(['Acq_Ret'], axis=1, inplace=True)
        # if this exists we will drop it
        try:
            df.drop(['Line Rental (Excl. VAT)'], axis=1, inplace=True)
        except:
            print("Line Rental not present")
        
        # duplicates the revenue column and converts to the price for commission
        df['COMMISSION'] = df['REVENUE'].apply(lambda x: x *.1)
        # merges the 2 dataframes
        df = pd.concat([static_df, df], ignore_index=True)
        df = pd.concat([df, business_rates], ignore_index=True)

        # creates a .CSV for webchat
        df.to_csv(output_folder + "Webchat.csv", index = None,  header=True)
        return
    else:
        # creates the commission columns for all stores
        df['COMMISSION_WR'] = df['Leeds White Rose']
        df['COMMISSION_L8'] = df['Leeds 8-9 Commercial Street']
        df['COMMISSION_CAS'] = df['Castleford']
        df['Gigafast'] = df['Leeds 8-9 Commercial Street']
        df['COMMISSION_GIG'] = df['Gigafast']

        # we run through the rows and set the correct commission for each store
        for index, row in df.iterrows():
            if 'HBB' in row['TYPE']:
                df.loc[index, "COMMISSION_WR"] = row['Leeds White Rose'] * .1
                df.loc[index,'COMMISSION_L8'] = row['Leeds 8-9 Commercial Street'] * .1
                df.loc[index, "COMMISSION_CAS"] = row['Castleford'] * .1
                if row["Acq_Ret"] == 'Acquisition':
                    df.loc[index, "COMMISSION_GIG"] = row['Leeds 8-9 Commercial Street'] * .3
                else:
                    df.loc[index, "COMMISSION_GIG"] = row['Leeds 8-9 Commercial Street'] * .3
            else:
                df.loc[index, "COMMISSION_WR"] = row['Leeds White Rose'] * .05
                df.loc[index,'COMMISSION_L8'] = row['Leeds 8-9 Commercial Street'] * .05
                df.loc[index, "COMMISSION_CAS"] = row['Castleford'] * .05
                df.loc[index, "COMMISSION_GIG"] = 0
                df.loc[index, "Gigafast"] = 0

        # this column is no longer needed so is dropped
        df.drop(['Acq_Ret'], axis=1, inplace=True)

        # there are 3 versions of the dataframe we need so this creates them
        wr_df = df.copy()
        l8_df = df.copy()
        cas_df = df.copy()
        gig_df = df.copy()

        # cleaning up all the data so that the My Ratecard database can read it
        l8_df = df.rename(columns={'Leeds 8-9 Commercial Street': 'REVENUE', 'COMMISSION_L8': 'COMMISSION'})
        l8_df.drop(['Gigafast','Leeds White Rose','Castleford', 'COMMISSION_WR', 'COMMISSION_CAS', 'COMMISSION_GIG'], axis=1, inplace=True)
        wr_df = df.rename(columns={'Leeds White Rose': 'REVENUE', 'COMMISSION_WR': 'COMMISSION'})
        wr_df.drop(['Gigafast','Leeds 8-9 Commercial Street','Castleford', 'COMMISSION_CAS', 'COMMISSION_GIG', 'COMMISSION_L8'], axis=1, inplace=True)
        cas_df = df.rename(columns={'Castleford': 'REVENUE', 'COMMISSION_CAS': 'COMMISSION'})
        cas_df.drop(['Gigafast','Leeds White Rose','Leeds 8-9 Commercial Street', 'COMMISSION_WR', 'COMMISSION_GIG', 'COMMISSION_L8'], axis=1, inplace=True)
        gig_df = df.rename(columns={'Gigafast': 'REVENUE', 'COMMISSION_GIG': 'COMMISSION'})
        gig_df.drop(['Castleford','Leeds White Rose','Leeds 8-9 Commercial Street', 'COMMISSION_WR', 'COMMISSION_CAS', 'COMMISSION_L8'], axis=1, inplace=True)

        # brings the 2 dataframes together
        wr_df = pd.concat([static_df, wr_df], ignore_index=True)
        cas_df = pd.concat([static_df, cas_df], ignore_index=True)
        l8_df = pd.concat([static_df, l8_df], ignore_index=True)
        try:
            wr_df = pd.concat([wr_df, business_rates], ignore_index=True)
            cas_df = pd.concat([cas_df, business_rates], ignore_index=True)
            l8_df = pd.concat([l8_df, business_rates], ignore_index=True)
        except:
            pass
        # exports all dataframes to .csv files
        wr_df.to_csv(output_folder + "Whiterose.csv", index = None,  header=True)
        l8_df.to_csv(output_folder + "Leeds-8.csv", index = None,  header=True)
        cas_df.to_csv(output_folder + "Castleford.csv", index = None,  header=True)
        gig_df.to_csv(output_folder + "Gigafast.csv", index = None,  header=True)
        return

if __name__ == "__main__":
    main()