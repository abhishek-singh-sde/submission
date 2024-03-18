#I have written different functions for handling different tasks to ensure modularity
#These functions can also be wrapped in an API with and the exact process flow of function 
#calls can be tweaked as per requirement

#Importing the libraries
import pandas as pd
from datetime import date
import time    
import aspose.pdf as ap
import sqlite3
import sys


def extractor(path_to_input_pdf):

    #We can use with and chunk_size to breakdown inputfile into smaller sizes and process for biger files
    
    document = ap.Document(path_to_input_pdf)
    save_option = ap.ExcelSaveOptions()

    #This is enabled so that we don't have to do the expensive computations of 
    #manually bringing data of different pages to one page (overriding default behaviour)
    save_option.minimize_the_number_of_worksheets = True

    #The column headings across pages will remain same that's why
    #Although I have tested without this also, and it happens to work
    save_option.uniform_worksheets=True

    #We will need to clean the extracted text as there are entries/extra text added by the library and Page No breaks
    temp_path="temp/temp_excel.xlsx"
    document.save(temp_path, save_option)
    df=pd.read_excel(temp_path)
    
    #The first two rows added by aspose are removed
    df = df.iloc[2:]
    
    #Instead of checking whether whole cell is merged or not, Or checking if detected text is Page <int>
    #We can just do a simple type check which will cover both NaN and str cases in one go
    for index,row in df.iterrows():
        if type(row.iloc[0])!=int:
            df.drop(index, inplace=True)
    
    df.columns = ['App ID','Xref','Settlement Date' ,'Broker','Sub Broker','Borrower Name','Description','Total Loan Amount','Comm Rate','Upfront','Upfront Incl GST']
    return df
    
def store_extracted_text(extracted_df,path_to_input_pdf):

    #As file with same name may be uploaded multiple times, Or it might also happen that same named files have different content
    #Hence, adding the timestamp to ensure uniqueness
    today = date.today()
    d1 = today.strftime("%d-%m-%Y")
    current_date=str(d1)

    t = time.localtime()
    current_time = time.strftime("%H%M%S", t)
    current_time=str(current_time)
    
    #The InputFileHistory folder will store the cleaned and generated text
    extracted_df.to_excel("InputFileHistory/"+path_to_input_pdf + " uploaded @ "+current_date+" "+current_time+".xlsx",index=False)
    
    #We will now check for any deduplication and then insert it
    #check_deduplication_then_insert(extracted_df)
    
    return extracted_df
    
    
def check_deduplication_then_insert(extracted_df):
    
    #This shall be our reference file (which will be empty initially) and will get eventually updated with time 
    df_org=pd.read_excel("Final_Updated_Records.xlsx")


    merged_df = pd.concat([df_org, extracted_df])
    merged_df = merged_df.drop_duplicates(subset=['Xref','Total Loan Amount'], keep='first',ignore_index=True)
    #merged_df.to_excel("Final_Updated_Records.xlsx",index=False)
    
    return merged_df

def loan_during_period(start_date,end_date):

    df_org=pd.read_excel("Final_Updated_Records.xlsx")
    conn = sqlite3.connect(':memory:')
    df_org.to_sql(name='Table1', con=conn)
    cur=conn.cursor()
    
    tup=cur.execute("SELECT SUM([Total Loan Amount]) FROM Table1 WHERE [Settlement Date] BETWEEN ? AND ?", [start_date, end_date],)
    print(list(tup)[0])
    
    cur.close()

def highest_loan_by_broker(broker_name):

    df_org=pd.read_excel("Final_Updated_Records.xlsx")
    conn = sqlite3.connect(':memory:')
    df_org.to_sql(name='Table1', con=conn)
    cur=conn.cursor()
    
    tup=cur.execute("SELECT MAX([Total Loan Amount]) FROM Table1 WHERE Broker=?", [broker_name],)
    print(list(tup)[0])
    
    cur.close()
    

def report_loans_by_broker(broker_name):
    df_org=pd.read_excel("Final_Updated_Records.xlsx")
    conn = sqlite3.connect(':memory:')

    df_org.to_sql(name='Table1', con=conn)
    cur=conn.cursor()
    tup=cur.execute(''' SELECT SUM([Total Loan Amount]) AS total_loan_amount,
                        strftime('%Y', [Settlement Date]) AS year,
                        strftime('%m', [Settlement Date]) AS month,
                        strftime('%d', [Settlement Date]) AS day
                        FROM Table1
                        WHERE Broker=?
                        GROUP BY strftime('%Y', [Settlement Date]), strftime('%m', [Settlement Date]), strftime('%d', [Settlement Date])
                        ORDER BY [Total Loan Amount] DESC''', [broker_name],)
    for i in tup:
        print(i)
    
    cur.close()
    

def report_loans_by_date():
    df_org=pd.read_excel("Final_Updated_Records.xlsx")
    conn = sqlite3.connect(':memory:')
    df_org.to_sql(name='Table1', con=conn)
    cur=conn.cursor()
    
    #Have skipped order by intentionally
    tup=cur.execute("SELECT [Settlement Date],SUM([Total Loan Amount]) FROM Table1 GROUP BY [Settlement Date]")
    for i in tup:
        print(i)
    
    cur.close()


def report_loans_by_tier():
    df_org=pd.read_excel("Final_Updated_Records.xlsx")
    conn = sqlite3.connect(':memory:')
    df_org.to_sql(name='Table1', con=conn)
    cur=conn.cursor()
    
    #Have skipped order by intentionally
    tup=cur.execute(''' SELECT [Settlement Date], 
                        CASE WHEN SUM([Total Loan Amount]) > 100000 THEN 1 
                             WHEN SUM([Total Loan Amount]) > 50000 THEN 2 
                             ELSE 3 END AS TIER, 
                        COUNT(*) 
                        FROM Table1 
                        GROUP BY [Settlement Date]''')
    for i in tup:
        print(i)


if __name__=="__main__":
    
    #path_to_input_pdf = "Test PDF.pdf"
    filename=sys.argv[1]
    path_to_input_pdf=filename[7:]
    
    #This function handles the data extraction
    extracted_df=extractor(path_to_input_pdf)

    #We can run this to keep a record of the uploaded/input file's generated text 
    extracted_df=store_extracted_text(extracted_df,path_to_input_pdf)

    #Finally, running the deduplications check
    merged_df=check_deduplication_then_insert(extracted_df)
    merged_df.to_excel("Final_Updated_Records.xlsx",index=False)
    


    #Reporting Functions' Sample Calls:
    
    #Gives total loan amount b/w two dates
    loan_during_period('10/10/2023','12/10/2023')

    #Shows the highest transaction/loan amount done by a particular broker
    highest_loan_by_broker('Alexander Foldi')
    
    #Shows datewise and tierwise count
    report_loans_by_tier()