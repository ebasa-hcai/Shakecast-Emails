import tkinter as tk
from tkcalendar import DateEntry, Calendar
from email_reader import emailReader




class App:
    def __init__(self):
        
        self.root = tk.Tk()
        self.root.title("Shakecast Email App") 
        self.root.geometry("700x600")
        
        self.start_cal = Calendar(self.root, selectmode = 'day', date_pattern="mm/dd/yyyy")
        self.start_cal.grid(row = 2, column = 0, pady = 5, padx = 20)

        self.end_cal = Calendar(self.root, selectmode = 'day', date_pattern="mm/dd/yyyy")
        self.end_cal.grid(row = 2, column = 1, pady = 5, padx= 20 )
        
        self.l1 = tk.Label(self.root, text = "Start Date ")
        self.l2 = tk.Label(self.root, text = "End Date ")
        
        self.l1.grid(row = 1, column = 0,  pady = 2)
        self.l2.grid(row = 1, column = 1 , pady = 2)
        
        self.button = tk.Button(self.root, text= "Select Dates", command= self.get_dates)
        self.button.grid(row= 4, column=0, pady= 10, padx = 100)
        
        self.button2 = tk.Button(self.root, text= "Get Emails!", command= self.email_get)
        self.button2.grid(row= 4,column=1, pady= 10, padx = 100)
        
        

        
        #
        self.start_date = None
        self.end_date = None
        
        #printing status label
        self.status_lb = tk.Label(self.root, text = "Status:  Ready!!", wraplength= 300)
        self.status_lb.grid(row =0, padx=20, pady=15)
        
        
    def status_msg(self,msg):
        self.status_lb.config(text = f"Status: {msg}")
        self.root.update_idletasks()

    def get_dates(self):
        self.start_date = self.start_cal.get_date()
        self.end_date = self.end_cal.get_date()



        print(f"{self.start_date}")
    
    
        
        
    def email_get(self):
        self.status_msg("Email Function Starting...")
        print("Email function started....")
        if not hasattr(self, 'start_date') or not hasattr(self, 'end_date'):
            print("select days")
            
        try:
            
            print(f"fetching from {self.start_date}")

            email_reader =emailReader()
            messages = email_reader.fetch_emails_between(start_date=self.start_date, end_date=self.end_date)
            emails = email_reader.clean_emails(messages)

            df_lis_min = email_reader.get_df(emails)
            df_lis_max = email_reader.get_df(emails)
            
            min_df = email_reader.merge_min(df_lis_min)
            max_df = email_reader.merge_max(df_lis_max)
            
            self.status_msg("Email Function Starting...")
            final_df = email_reader.merge_df(min_df=min_df, max_df=max_df)
            print(final_df)
            self.status_msg("Saving CSV File!! ")

            email_reader.save_csv(final_df)
            
        except Exception as e:
            if str(e) == 'No objects to concatenate':
                self.status_msg("No Emails in Date Range")
                print("No Emails in Date Range")
                
            else:
                self.status_msg(f"Error: {e}")
                print({e})

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = App()
    app.run()   
