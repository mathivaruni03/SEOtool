import sqlite3
from matplotlib import pyplot as plt
import bs4
import requests
import json
import xlsxwriter
import xlrd



class Webscrap:
    final_data = []
    read_data=[]
    def data_extract(self,soup):
        data = json.loads(str(soup))
        data1 = data["personList"]["personsLists"]
        # print(data1)
        for dt in data1:
            self.data_one = [
                dt.get("rank", ''),
                dt.get("personName", ''),
                dt.get("finalWorth", ''),
                dt.get("age", ''),
dt.get("source", ''),
                dt.get("state", '')
            ]
            # print(data_one)
            self.final_data.append(self.data_one)
        # print(self.final_data)

    def excel_data(self):
        workbook=xlsxwriter.Workbook("Self_Made_Women_2021.xlsx")
        worksheet=workbook.add_worksheet()
        bold=workbook.add_format({'bold':True})
        worksheet.write('A1','RANK',bold)
        worksheet.write('B1','NAME',bold)
        worksheet.write('C1','NET_WORTH',bold)
        worksheet.write('D1','AGE',bold)
        worksheet.write('E1','SOURCE',bold)
        worksheet.write('F1','STATE',bold)
        print("***********************************************")
        print("EXCEL SHEET CREATED SUCCESSFULLY")
        row=1
        col=0
        for data1 in self.final_data:
            worksheet.write(row,col,data1[0])
            worksheet.write(row,col+1,data1[1])
            worksheet.write(row, col + 2, data1[2])
            worksheet.write(row, col + 3, data1[3])
            worksheet.write(row, col + 4, data1[4])
            worksheet.write(row, col + 5, data1[5])

            row+=1
        print("*******************************************")
        print("DATA PRINTED SUCCESSFULLY")

        chart1=workbook.add_chart({'type':'column'})
        chart1.add_series({'categories':'=Sheet1!$B$2:$B$50','values':'=Sheet1!$C$2:$C$50'})
        #chart1.add_series({'categories':'=Sheet1!$B$2:$B$50','values':'=Sheet1!$C$2:$C$50'})
        chart1.set_title({'name':'Self made Women:2021'})
        worksheet.insert_chart('K4',chart1)
        workbook.close()
        print("*********************************************")
        print("GRAPH OF THE DATA SUCCESSFULLY DRAWN ON EXCEL SHEET")

    def read_excel(self):
        wb=xlrd.open_workbook("Self_Made_Women_2021.xlsx")
worksheet= wb.sheet_by_name("Sheet1")
num_rows=worksheet.nrows
num_cols=worksheet.ncols

for current_row in range(0,num_rows,1):
            row_review=[]
            for current_col in range(0,num_cols,1):
                review=worksheet.cell_value(current_row,current_col)
                row_review.append(review)
            self.read_data.append(row_review)
        #print(self.read_data)
print("***********************************************")
print("DATA READ FROM EXCEL AND STORED SUCCESSFULLY...!!!")
def data_base(self):
        db_values=self.read_data
        con=sqlite3.connect("SELF_MADE_WOMEN_2021.db")
        print("*********************************")
        print("DATABASE CONNECTED SUCCESSFULLY..!!!")
        cur=con.cursor()
        listofTables=cur.execute("""SELECT 'SELF_MADE_WOMEN_2021' FROM sqlite_master WHERE type='table'""").fetchall()

        if listofTables==[]:
            cur.execute('''CREATE TABLE SELF_MADE_WOMEN_2021(
                Rank INTEGER NOT NULL,
                Person_name TEXT,
                Net_worth INTEGER NOT NULL,
                Age INTEGER NOT NULL,
                Source TEXT,
                State TEXT);''')

        else:
            print('Table found!!')

        #print(db_values)

        cur.executemany("INSERT INTO SELF_MADE_WOMEN_2021(Rank,Person_name,Net_worth,Age,Source,State) VALUES(?,?,?,?,?,?)",db_values)
        con.commit()
        print("*****************************")
        print("DATA STORED IN DATABASE SUCCESSFULLY!!")

def graph(self):
        first_plot=[dt[1] for dt in self.final_data]
        second_plot=[dt[2] for dt in self.final_data]
        #print(first_plot)
#print(second_plot)
        plt.bar(first_plot,second_plot,color='r')
        plt.legend(["NAMES","NET_WORTH"])
        plt.show()
        print("********************************")
        print("BAR GRAPH DRAWN SUCCESSFULLY...!!")

        plt.scatter(first_plot,second_plot,label='cases',color='r')
        plt.show()
        print("********************************")
        print("SCATTER GRAPH DRAWN SUCCESSFULLY...!!")





urllink="https://www.forbes.com/forbesapi/person/self-made-women/2021/position/true.json"
header={
'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36'
}

response=requests.get(url=urllink,headers=header)
soup=bs4.BeautifulSoup(response.content,"html.parser")
#print(soup.prettify())
w=Webscrap()
w.data_extract(soup)
w.excel_data()
w.read_excel()
w.data_base()
w.graph()
