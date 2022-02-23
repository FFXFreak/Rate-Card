import csv
import sys
import xlrd

import pandas as pd
import glob, os

input_folder = "\\Rate_Cards\\"
output_folder = "\\Output_Rate_Cards\\"
static_folder = "\\Static\\"

def main():

    files = []

    #ensure files are in folder
    print(
        "Please ensure that all Rate Cards that need to be converted are in the Rate_Cards folder and you have included the static ratecard before you continue!")
    run = input("Please enter Y to continue: ")

    if run.upper() != "Y":
        sys.exit("User not confirmed.")

    #stores list of files in folder also changes current directory down a level
    os.chdir(os.getcwd() + input_folder)
    for file in glob.glob("*.xlsx"):
        files.append(file)

    #set the current directory back to the origin
    os.chdir(os.path.dirname(os.getcwd()))

    #ensures a min of 1 and max of 2 files in folder
    if not files and len(files) < 3:
        sys.exit("There are currently " + str(len(files)) + " files in the folder. Expecting Minimum of 1 and Maximum of 2")

    #we now are going to run this script until all files have been processed
    runtime = len(files)
    while runtime != 0:
        process_rate_card(runtime, files)
        runtime -= 1
    

def process_rate_card(filenumber, files):
    #adds excel file and static files to dataframe.
    df = pd.read_excel(os.getcwd() + input_folder + files[filenumber - 1])
    static_df = pd.read_csv(os.getcwd() + static_folder + "static_rates.csv")

    #rename and drop apropriate columns
    df = df.rename(columns={'SOC Code': 'SKU', 'Description': 'DESCRIPTION'})
    df['SKU'] = df.apply(lambda x: x['SKU']+"N" if x['Acq_Ret'] ==
                        'Acquisition' and x['TYPE'] != 'TALK MOBILE' else x['SKU'], axis=1)
    df.dropna(subset=['TYPE'], inplace=True)

    #testing talk mobile SKU creation tool
    #converts the 
    df['MAF (Inc VAT)'] = df['MAF (Inc VAT)'].apply(lambda x: x *100).map('{0:g}'.format)
    
    #we itterate through all talk mobile deals and convert to the format TM + contract length + deal price
    #eg for a 1 month contract at 7.95 = TM300795, 12 months at 11.50 = TM121150.
    for index, row in df.iterrows():
        if row['TYPE'] == 'TALK MOBILE':
            if int(row['MAF (Inc VAT)']) < 1000:
                row['MAF (Inc VAT)'] = "0" + str(row['MAF (Inc VAT)'])
            if row['Contract Length (Months)'] == 1:
                row['Contract Length (Months)'] = 30
            #applies changed to SKU
            df.loc[index, "SKU"] = "TM" + str(row['Contract Length (Months)']) + str(row['MAF (Inc VAT)'])
    
    df.drop(['Legacy Code - EBU', 'Band', 'MAF (Inc VAT)','Contract Length (Months)'], axis=1, inplace=True)

    #establish if webchat ratecard or retail
    if 'Webchat' in df.columns:
        df = df.rename(columns={'Webchat': 'REVENUE'})
        df.drop(['Acq_Ret','Line Rental (Excl. VAT)'], axis=1, inplace=True)
        df['COMMISSION'] = df['REVENUE'].apply(lambda x: x *.1)
        df = pd.concat([static_df, df], ignore_index=True)

        df.to_csv("Output_Rate_Cards\\Webchat.csv", index = None,  header=True)
    else:
        df['COMMISSION_WR'] = df['Leeds White Rose']
        df['COMMISSION_CAS'] = df['Castleford']
        df['Gigafast'] = df['Leeds White Rose']
        df['COMMISSION_GIG'] = df['Gigafast']

        #print(df)
        for index, row in df.iterrows():
            if 'HBB' in row['TYPE']:
                df.loc[index, "COMMISSION_WR"] = row['Leeds White Rose'] * .1
                df.loc[index, "COMMISSION_CAS"] = row['Castleford'] * .1
                if row["Acq_Ret"] == 'Acquisition':
                    df.loc[index, "COMMISSION_GIG"] = 50
                else:
                    df.loc[index, "COMMISSION_GIG"] = 20
            else:
                df.loc[index, "COMMISSION_WR"] = row['Leeds White Rose'] * .05
                df.loc[index, "COMMISSION_CAS"] = row['Castleford'] * .05
                df.loc[index, "COMMISSION_GIG"] = 0
                df.loc[index, "Gigafast"] = 0

        df.drop(['Acq_Ret'], axis=1, inplace=True)

        wr_df = df.copy()
        cas_df = df.copy()
        gig_df = df.copy()

        wr_df = df.rename(columns={'Leeds White Rose': 'REVENUE', 'COMMISSION_WR': 'COMMISSION'})
        wr_df.drop(['Gigafast','Leeds 8-9 Commercial Street','Castleford', 'COMMISSION_CAS', 'COMMISSION_GIG'], axis=1, inplace=True)
        cas_df = df.rename(columns={'Castleford': 'REVENUE', 'COMMISSION_CAS': 'COMMISSION'})
        cas_df.drop(['Gigafast','Leeds White Rose','Leeds 8-9 Commercial Street', 'COMMISSION_WR', 'COMMISSION_GIG'], axis=1, inplace=True)
        gig_df = df.rename(columns={'Gigafast': 'REVENUE', 'COMMISSION_GIG': 'COMMISSION'})
        gig_df.drop(['Castleford','Leeds White Rose','Leeds 8-9 Commercial Street', 'COMMISSION_WR', 'COMMISSION_CAS'], axis=1, inplace=True)

        wr_df = pd.concat([static_df, wr_df], ignore_index=True)
        cas_df = pd.concat([static_df, cas_df], ignore_index=True)
        gig_df = pd.concat([static_df, gig_df], ignore_index=True)

        wr_df.to_csv("Output_Rate_Cards\\Whiterose-L8.csv", index = None,  header=True)
        cas_df.to_csv("Output_Rate_Cards\\Castleford.csv", index = None,  header=True)
        gig_df.to_csv("Output_Rate_Cards\\Gigafast.csv", index = None,  header=True)

if __name__ == "__main__":
    main()