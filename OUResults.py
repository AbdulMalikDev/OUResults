import bs4
import mechanize
import requests
from requests.adapters import HTTPAdapter
# from requests.packages.urllib3.util.retry import Retry
from xlwt import Workbook


class UniversityResult:

    # Enter the OU result link here ⬇️
    result_link = "https://www.osmania.ac.in/res07/20220625.jsp" 

    pre_link = result_link + "?mbstatus&htno="
    row = 1

    # ⬇️ Enter college, field and year of results you need here  
    collegeCodes = ["1604"]
    fieldCodes = ["733"]
    year = "20"

    
    def __init__(self):
        self.globalbr = mechanize.Browser()
        self.globalbr.set_handle_robots(False)
        self.wb, self.sheet = self.getNewExcelSheet()
        self.initiateFindingResult()

    def initiateFindingResult(self):

        for fieldCode in self.fieldCodes:
            for collegeCode in self.collegeCodes:


                for index in range(1, 121):
                    hall_ticket = collegeCode + self.year + \
                        fieldCode + str(index).zfill(3)
                    self.findResult(fieldCode, collegeCode,
                                    hall_ticket, int(str(index).zfill(3)))

                # Lateral Entry students
                for index in range(301, 313):
                    hall_ticket = collegeCode + self.year + \
                        fieldCode + str(index).zfill(3)
                    self.findResult(fieldCode, collegeCode,
                                    hall_ticket, int(str(index).zfill(3)))

        # ⬇️ Modify output EXCEL file name here                            
        excelFileName = "RESULT" + "_" + fieldCode + "_" + self.year + ".xls"

        self.wb.save(excelFileName)
        self.wb, self.sheet = self.getNewExcelSheet()

    def findResult(self, fieldCode, collegeCode, hall_ticket, index):

        resultLink = self.pre_link + hall_ticket
        session = requests.Session()
        # retry = Retry(connect=3, backoff_factor=0.5,total=10)
        adapter = HTTPAdapter()
        session.mount('http://', adapter)
        session.mount('https://', adapter)
        raw = session.get(resultLink)

        soup = bs4.BeautifulSoup(raw.content, "html.parser")

        # Find the Name
        table = soup.find(id="AutoNumber3")
        if table is None:
            return
        last_row = table("tr")[2]
        td_list = last_row.find_all("td")
        name = td_list[1].text

        # Find the GPA
        table = soup.find(id="AutoNumber5")
        if table is None:
            return
        last_row = table("tr")[-1]
        td_list = last_row.find_all("td")
        marks = td_list[1].text

        # Store current student result in excel sheet row 
        self.sheet.write(self.row, 0, self.row)
        self.sheet.write(self.row, 2, str(collegeCode))
        self.sheet.write(self.row, 3, str(fieldCode))
        self.sheet.write(self.row, 4, hall_ticket[-3:])
        self.sheet.write(self.row, 5, marks)
        self.sheet.write(self.row, 6, name)
        self.row += 1

        print("Row Number : " + str(self.row-1))
        print(name + " " + hall_ticket + " " + marks)

    # Initializing the sheet  
    def getNewExcelSheet(self):
        wb = Workbook()
        sheet1 = wb.add_sheet('Sheet 1')
        sheet1.write(0, 0, 'S.No')
        sheet1.write(0, 1, 'Rank')
        sheet1.write(0, 2, 'Code')
        sheet1.write(0, 3, 'Field')
        sheet1.write(0, 4, 'R.No')
        sheet1.write(0, 5, "CGPA")
        sheet1.write(0, 6, "Name")
        return [wb, sheet1]


UniversityResult()
