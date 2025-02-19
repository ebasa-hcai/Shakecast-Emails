#import library  
import win32com.client
import pandas as pd
import numpy as np
from datetime import datetime,timedelta


'''
Used to read and clean shakecast emails 
Takes Max (excpet for distance the min is used)
Connects to your outlook
'''

class emailReader:
    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.inbox = self.outlook.GetDefaultFolder(6)
    #fetch the emails
    def fetch_emails_between(self,start_date , end_date):
        start_date = start_date +" 00: 00"
        end_date = end_date +" 23: 59"
        
        
        #find emails between 2 dates
        restriction = f"[ReceivedTime] >= '{start_date}' AND  [ReceivedTime] <= '{end_date}'"
        #only fetch emails from hours back
        messages = self.inbox.Items.Restrict(restriction)

        return messages

    
    def clean_emails(self, messages,starts = ["UPDATE: Inspection","Inspection -"]):
        all_dates = []
        all_emails =[]
        for start in starts:
            emails =[]
            dates = []
            for message in messages:
                if message.Class == 43:
                    subject = message.Subject
                    # sender = message.SenderName
                    body = message.Body
                    time = message.ReceivedTime.date()

                
                    if subject.startswith(start):
                        output = body
                        #clean
                        output = output.split("of shaking. ")
                        output = output[1]
                        output = output.replace("\t", "\n")
                        output = output.split("ShakeCast Server")
                        output = (output[0]).replace("\r",'')

                        output = (output).replace("\n \n",'\n')
                        output = (output).replace("\n ",'\n')
                        output = (output).replace(" \n",'\n')
                        output = (output).replace("\n\n\n",'')
                        #split by new line
                        info = (output.split("\n"))
                        info.pop()

                        emails.append(info)
                        #append time as  str
                        dates.append(time.strftime("%Y-%m-%d"))
            all_dates+=dates
            all_emails+=emails
            
        emails = all_emails
        dates = all_dates
        return all_emails, all_dates

    def get_df(self, email_lis):
        emails = email_lis[0]
        dates = email_lis[1]

        data_frame=[]
        for i in range(len(emails)):
            #output
            table_cols = (emails[i][0:13])
            data = (emails[i])[13:]
            row_size = int(len(data)/len(table_cols))
            df = pd.DataFrame(np.array(data).reshape(row_size,len(table_cols)), columns = table_cols)
            df['Date'] = dates[i]
            df['Inspection Priority_date'] = df['Date']
            df['Distance (km)_date'] = df['Date']
            df['MMI_date'] = df['Date']
            df['PGA (%g)_date'] = df['Date']
            df['PGV (cm/s)_date'] = df['Date']
            df['PSA03 (%g)_date'] = df['Date']
            df['PSA10 (%g)_date'] = df['Date']
            df['PSA30 (%g)_date'] = df['Date']
            df['Shaking Value_date'] = df['Date']
            df = df.drop('Date', axis=1)
            data_frame.append(df)
            
            
        return data_frame

    def priority_mapping(self,df_list, status_col = "Inspection Priority"):
        mapping ={'Low': 0 , 'Medium': 1, 'Medium-High': 2, 'High':3}
        for df in df_list:
            df[status_col] = df[status_col].map(mapping)
        
        return df_list
        
        

    def reverse_mapping(self, df, status_col = "Inspection Priority"):
        reverse_map ={ 0:  'Low',  1:'Medium',  2:'Medium-High', 3:'High'}
        #for df in df_list:
        df[status_col] = df[status_col].map(reverse_map)
            
        return df   

    def get_max_idx(self, group, num_col,df):
        idx = df.groupby(group)[num_col].idxmax()
        #df[f'{group}_date']
        return idx
    

        
        
    def merge_max(self, df_list):
        df_list_mapped = self.priority_mapping(df_list)
        combined_df = pd.concat(df_list_mapped, ignore_index= True)
        combined_df.drop(['Distance (km)', 'Distance (km)_date'], axis = 1) 
        measurements = ['Inspection Priority','MMI','PGA (%g)','PGV (cm/s)','PSA03 (%g)','PSA10 (%g)','PSA30 (%g)','Shaking Value']
        non_num_cols = ['Facility Name', 'Facility Type', 'Short Name','Metric' ]
        max_df = combined_df.groupby(non_num_cols).max().reset_index()
        max_idx=combined_df.groupby(non_num_cols).idxmax().reset_index()
        for col in range(len(measurements)):
            col_index = (max_idx[measurements[col]]).to_list()
            for i in range(len(col_index)):
                max_df.iloc[i,max_df.columns.get_loc(f'{measurements[col]}_date')] = combined_df.loc[col_index[i],f'{measurements[col]}_date' ]
        
        max_df=max_df.sort_values(by=['Inspection Priority','Shaking Value'], ascending=False)
        reverse_map_df = self.reverse_mapping(max_df)
        
        # reorder cols 
        reverse_map_df = reverse_map_df.reindex(['Facility Name', 'Facility Type', 'Short Name', 'Metric','Inspection Priority', 'Inspection Priority_date',  'MMI',  'MMI_date', 'PGA (%g)', 'PGA (%g)_date','PGV (cm/s)','PGV (cm/s)_date','PSA03 (%g)', 'PSA03 (%g)_date', 'PSA10 (%g)', 'PSA10 (%g)_date', 'PSA30 (%g)', 'PSA30 (%g)_date', 'Shaking Value','Shaking Value_date'], axis = 1)
        
        
        
                
        return reverse_map_df 



    
    def merge_min(self,df_list):
        df_list_mapped = self.priority_mapping(df_list)
        combined_df = pd.concat(df_list_mapped, ignore_index= True) 
        combined_df.drop(['Inspection Priority','MMI','PGA (%g)','PGV (cm/s)','PSA03 (%g)','PSA10 (%g)','PSA30 (%g)','Shaking Value','Metric','Facility Name', 'Facility Type'], axis = 1)
        measurements = ['Distance (km)']
        non_num_cols = [ 'Short Name']
        min_df = combined_df.groupby(non_num_cols).min().reset_index()
        min_idx=combined_df.groupby(non_num_cols).idxmin().reset_index()
        for col in range(len(measurements)):
            col_index = (min_idx[measurements[col]]).to_list()
            for i in range(len(col_index)):
                min_df.iloc[i,min_df.columns.get_loc(f'{measurements[col]}_date')] = combined_df.loc[col_index[i],f'{measurements[col]}_date' ]
        
        min_df=min_df.sort_values(by=['Inspection Priority','Shaking Value'], ascending=False)
        reverse_map_df = self.reverse_mapping(min_df)
        
        # reorder cols 
        reverse_map_df = reverse_map_df.reindex([ 'Short Name', 'Distance (km)', 'Distance (km)_date'], axis = 1)
        
        
        
                
        return reverse_map_df

    

    def merge_df(self, min_df,max_df):
        
        return pd.merge(max_df, min_df, on='Short Name')
    
    def save_csv(self, max_df, start_date , end_date):
        start_date = start_date.replace('/','-')
        end_date = end_date.replace('/','-')
        currentDateTime = datetime.now().strftime("%m-%d-%Y %H-%M-%S %p")
        max_df.to_csv(f"ShakeCast_table Date_Range- ({start_date}- {end_date}) ---- {currentDateTime}.csv", index = False)
        