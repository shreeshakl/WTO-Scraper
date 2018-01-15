#!/usr/bin/env python

import xlsxwriter
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from bs4 import BeautifulSoup

workbook = xlsxwriter.Workbook('data.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column(0, 0, 30)
worksheet.set_column(1, 1, 30)
worksheet.set_column(2, 2, 50)
worksheet.set_column(3, 3, 25)
worksheet.set_column(4, 7, 15)


class WTOscraper(object):
    def __init__(self):
        self.url = "http://ptadb.wto.org/ptaList.aspx"
        self.driver = webdriver.PhantomJS()
        self.driver.set_window_size(1120, 550)
        self.idNumber=0
        self.excelRow=0


    def scrape(self):
        while True:
            id="MainContent_ptaListControl1_GridView1_pta_hyperlink_";
            try:
                self.driver.get(self.url)
                soup = BeautifulSoup(self.driver.page_source)
                next_page_elem =self.driver.find_element_by_id(id+str(self.idNumber))
                print(soup.find(id=id+str(self.idNumber)).text)
                worksheet.write(self.excelRow, 0,soup.find(id=id+str(self.idNumber)).text)
                # self.excelRow+=1
                next_page_elem.click()
            except NoSuchElementException:
                print("Total Links:"+str(self.idNumber))
                break # no more pages

            def BasicInfoAndPTA(self,soup,title):
                tableLeft=soup.find_all("td", class_="td_pta_box_header_main");
                tableRight=soup.find_all("td", class_="td_pta_box_header");
                worksheet.write(self.excelRow, 1,title)
                for headings in (tableLeft,tableRight):
                    for heading in headings:
                        try:
                            tr=heading.parent.next_sibling.next_sibling.td.div.table.tbody.tr
                            print(heading.text.strip())
                            worksheet.write(self.excelRow, 2,heading.text.strip())
                            temp=self.excelRow
                            while True:
                                try:
                                    print(tr.td.a.text)
                                    worksheet.write(self.excelRow, 3,tr.td.a.text.strip())
                                    self.excelRow+=1
                                except:
                                    try:
                                        print(tr.td.text)
                                        worksheet.write(self.excelRow, 3,tr.td.text.strip())
                                        self.excelRow+=1
                                    except:
                                        if(temp==self.excelRow): #Means no data for heading, so clearing heading
                                            worksheet.write_blank(self.excelRow, 2, None)
                                        break
                                try:
                                    tr=tr.next_sibling
                                except:
                                    if(temp==self.excelRow): #Means no data for heading, so clearing heading
                                        worksheet.write_blank(self.excelRow, 2, None)
                                    break
                        except:
                            pass

            #Basic Info
            soup = BeautifulSoup(self.driver.page_source)
            BasicInfoAndPTA(self,soup,"Basic Info")

            next_page_elem =self.driver.find_element_by_link_text("Beneficiaries")
            next_page_elem.click()
            soup = BeautifulSoup(self.driver.page_source)
            table=soup.find(id="MainContent_ptaInfo_ptaBenefList_div_beneficiaries_main")
            if(table!=None):
                print("beneficiaries link")
                worksheet.write(self.excelRow, 1,"Beneficiaries")
                for link in table.findAll('a'):
                    if(link.text!=None):
                        print(link.text)
                        if(link.text=="Country / Territory"):
                            worksheet.write(self.excelRow, 2,link.text)
                        else:
                            worksheet.write(self.excelRow, 3,link.text)
                            self.excelRow+=1

            #PTA documentation
            next_page_elem =self.driver.find_element_by_link_text("PTA documentation")
            next_page_elem.click()
            soup = BeautifulSoup(self.driver.page_source)
            BasicInfoAndPTA(self,soup,"PTA documentation")


            next_page_elem =self.driver.find_element_by_link_text("Tariffs & Trade")
            next_page_elem.click()
            soup = BeautifulSoup(self.driver.page_source)
            table=soup.find(id="MainContent_ptaInfo_ptaTariffAndTrade1_TableDutyStats")
            if(table!=None):
                print("Links in Tarrifs and Trade")
                worksheet.write(self.excelRow, 1,"Tarrifs and Trade")
                for link in table.findAll('a'):
                    if(link.text!=None):
                            worksheet.merge_range(self.excelRow,2,self.excelRow+1, 2,link.text)
                            worksheet.write(self.excelRow, 3,link.parent.next_sibling.text)
                            worksheet.write(self.excelRow, 4,link.parent.next_sibling.next_sibling.text)
                            worksheet.write(self.excelRow, 5,link.parent.next_sibling.next_sibling.next_sibling.text)
                            worksheet.write(self.excelRow, 6,link.parent.next_sibling.next_sibling.next_sibling.next_sibling.text)
                            self.excelRow+=1
                            worksheet.write(self.excelRow, 3,link.parent.parent.next_sibling.td.text)
                            worksheet.write(self.excelRow, 4,link.parent.parent.next_sibling.td.next_sibling.text)
                            worksheet.write(self.excelRow, 5,link.parent.parent.next_sibling.td.next_sibling.next_sibling.text)
                            worksheet.write(self.excelRow, 6,link.parent.parent.next_sibling.td.next_sibling.next_sibling.next_sibling.text);
                            self.excelRow+=1

            try:
                table=soup.find(id="MainContent_ptaInfo_ptaTariffAndTrade1_TableImportStats")
                first_row=table.tbody.tr.next_sibling.next_sibling
                tbody=table.tbody
                tr=tbody.find_all("tr")
                worksheet.write(self.excelRow, 1,"Imports")
                for i in range(2,len(tr)):
                    td=tr[i].find_all("td")
                    if(((i-2)%3)==0):
                        column=2
                        worksheet.merge_range(self.excelRow,2,self.excelRow+2,2, 'Merged Cells')
                    else:
                        column=3
                    for j in range(0,len(td)):
                        worksheet.write(self.excelRow, column,td[j].text.strip())
                        column+=1
                    self.excelRow+=1
            except:
                pass
            self.idNumber+=1
        self.driver.quit()
        workbook.close()



if __name__ == '__main__':
    scraper = WTOscraper()
    scraper.scrape()
