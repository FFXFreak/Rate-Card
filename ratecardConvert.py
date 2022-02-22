import csv
import sys
import xlrd

import pandas as pd

def main():

    #get filename
    print("Please provide the name of the excel file without extention and ensure it is in the current folder")
    filename = input("Filename: ")

    #adds excel file and static files to dataframe and ensures no errors.
    try:
        df = pd.read_excel(filename + ".xlsx")
    except:
        sys.exit("Excel file either doesn't in folder on has been incorrectly inputted")
    try:
        static_df = pd.read_csv("Static\\static_rates.csv")
    except:
        sys.exit("Static.csv not in expected location")

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
        df['COMMISSION_GIG'] = df['Leeds White Rose']

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

        df.drop(['Acq_Ret'], axis=1, inplace=True)

        wr_df = df.copy()
        cas_df = df.copy()
        gig_df = df.copy()

        wr_df = df.rename(columns={'Leeds White Rose': 'REVENUE', 'COMMISSION_WR': 'COMMISSION'})
        wr_df.drop(['Leeds 8-9 Commercial Street','Castleford', 'COMMISSION_CAS', 'COMMISSION_GIG'], axis=1, inplace=True)
        cas_df = df.rename(columns={'Castleford': 'REVENUE', 'COMMISSION_CAS': 'COMMISSION'})
        cas_df.drop(['Leeds White Rose','Leeds 8-9 Commercial Street', 'COMMISSION_WR', 'COMMISSION_GIG'], axis=1, inplace=True)
        gig_df = df.rename(columns={'Leeds White Rose': 'REVENUE', 'COMMISSION_GIG': 'COMMISSION'})
        gig_df.drop(['Leeds 8-9 Commercial Street', 'COMMISSION_WR', 'COMMISSION_CAS'], axis=1, inplace=True)

        wr_df = pd.concat([static_df, wr_df], ignore_index=True)
        cas_df = pd.concat([static_df, cas_df], ignore_index=True)
        gig_df = pd.concat([static_df, gig_df], ignore_index=True)

        wr_df.to_csv("Output_Rate_Cards\\Whiterose-L8.csv", index = None,  header=True)
        cas_df.to_csv("Output_Rate_Cards\\Castleford.csv", index = None,  header=True)
        gig_df.to_csv("Output_Rate_Cards\\Gigafast.csv", index = None,  header=True)

if __name__ == "__main__":
    main()