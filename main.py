import time, re, collections, shutil, os, sys, zipfile, xlrd, threading, csv, openpyxl, math
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

def takeFirst(elem):
    return elem[0]

class mainProcesser():
    def __init__(self):

        self.txt_name = "Data/Tax Roll.txt"
        self.xlsx_name = "Data/Sending Template Tax Roll.xlsx"

        self.db_name = "Data/worldcitiespop.csv"
        self.total_data = []

        self.outfile_name = "Result/Sending Template Tax Roll.xlsx"

        self.dest = "Result"
        if not os.path.exists(self.dest):
            os.makedirs(self.dest)

        if os.path.exists(self.outfile_name):
            self.wb1 = openpyxl.load_workbook(self.outfile_name)
        else:
            self.wb1 = openpyxl.Workbook()
        self.ws1 = self.wb1.active

    def loadingData(self):
        self.wb1 = openpyxl.load_workbook(self.xlsx_name)
        self.ws1 = self.wb1.active

        self.input_data = open(self.txt_name, "r").read().split("\n")

        indices = []
        for i, row in enumerate(self.input_data):
            if row == "":
                continue
            else:
                if len(row) == 131:
                    indices.append(i)
                    # print i

        self.paragraphs = []
        self.paragraphs_cnt = 0

        for i in range(len(indices)-1):
            # print i
            # print indices[i:i+2]
            self.paragraphs.append([self.paragraphs_cnt, self.input_data[indices[i]:indices[i+1]]])
            self.paragraphs_cnt += 1

        self.paragraphs.append([self.paragraphs_cnt, self.input_data[indices[i + 1]:]])
        self.paragraphs_cnt += 1

        self.paragraphs.reverse()

        self.csv_data = []
        self.db_file = open(self.db_name, "rb").read()
        csv_reader = self.db_file.split("\n")

        for i, row in enumerate(csv_reader):
            row = row.split(",")
            if row[0] == "us":
                self.csv_data.append(row)
            # print row.split(",")

    def totalProcessing(self):
        self.threads = []
        self.max_threads = 2

        while self.threads or self.paragraphs:
            for thread in self.threads:
                if not thread.is_alive():
                    self.threads.remove(thread)

            while len(self.threads) < self.max_threads and self.paragraphs:
                thread = threading.Thread(target=self.processOneParagraph())
                thread.setDaemon(True)
                thread.start()
                self.threads.append(thread)

    def processOneParagraph(self):
        index, paragraph = self.paragraphs.pop()
        # index, paragraph = self.paragraphs[219]
        fullname = ""
        address = ""
        city = ""
        state = ""
        zipcode = ""
        section = ""
        township = ""
        range = ""

        print "++++++++++++++++++++++++ {0} ++++++++++++++++++++++++++++++++++++++++++++++++++++".format(index)
        try:
            [fullname, address, city, state, zipcode, section, township, range] = self.scanLineByLine(paragraph)
        except:
            for line in paragraph:
                print line
            exit(1)

        # [fullname, address, city, state, zipcode, section, township, range] = self.scanLineByLine(paragraph)

        print "Full Name:\t", fullname
        print "Address:\t", address
        print "City:\t\t", city
        print "State:\t\t", state
        print "Zipcode:\t", zipcode
        print "Section:\t", section
        print "Township:\t", township
        print "Range:\t\t", range

        self.total_data.append([
            index, fullname, address, city, state, zipcode, section, township, range
        ])

        # line = [fullname, address, city, state, zipcode]
        # for j, elm in enumerate(line):
        #     self.ws1.cell(row= index+2, column=j + 1).value = elm
        #     self.wb1.save(self.outfile_name)

    def scanLineByLine(self, paragraph):
        for i, line1 in enumerate(paragraph):
            line1 = line1.strip().split("  ")[0]

            if i < 2:
                continue


            for j, line2 in enumerate(self.csv_data):
                search_wd = " ".join([line2[1].upper(), line2[3]])

                if search_wd in line1:
                    city = line2[1].upper()
                    state = line2[3]
                    zipcode = line1.strip().split("  ")[0].split(" ")[-1]
                    address = paragraph[i-1].strip().split("   ")[0]
                    fullname = paragraph[0].strip().split("   ")[0]

                    tmp_str = fullname.split(" ")
                    if tmp_str[0].isdigit():
                        fullname = " ".join(tmp_str[1:])

                    str = "\n".join(paragraph)
                    regex = r"([\d]+)-([\d]+)-([\d]+)"
                    matches = re.findall(regex, str)

                    if matches:
                        section, township, range = matches[0]
                    else:
                        section, township, range = ("", "", "")

                    return [fullname, address, city, state, zipcode, section, township, range]


        for i, line1 in enumerate(paragraph):
            if i < 2:
                continue
            for j, line2 in enumerate(self.csv_data):

                search_wd = " ".join([line2[1].upper(), line2[3]])
                line1 = line1.strip().split("  ")[0]
                if fuzz.partial_ratio(search_wd, line1) > 86:
                    # city = line2[1].upper()
                    # state = line2[3]
                    tmp = line1.split(" ")
                    if len(tmp) == 3:
                        city = tmp[0]
                        state = tmp[1]
                        zipcode = tmp[2]
                    elif len(tmp) > 3:
                        city = " ".join(tmp[0:2])
                        state = tmp[2]
                        zipcode = tmp[3]
                    else:
                        return ["", "", "", "", "", "", "", ""]

                    # state = line1.split(" ")
                    # zipcode = line1.strip().split("  ")[0].split(" ")[-1]
                    address = paragraph[i-1].strip().split("   ")[0]
                    fullname = paragraph[0].strip().split("   ")[0]

                    tmp_str = fullname.split(" ")
                    if tmp_str[0].isdigit():
                        fullname = " ".join(tmp_str[1:])

                    str = "\n".join(paragraph)
                    regex = r"([\d]+)-([\d]+)-([\d]+)"
                    matches = re.findall(regex, str)

                    if matches:
                        section, township, range = matches[0]
                    else:
                        section, township, range = ("", "", "")

                    return [fullname, address, city, state, zipcode, section, township, range]

        return ["", "", "", "", "", "", "", ""]


    def saveCSV(self):

        print "--------- Start saving XLSX ------------"
        self.dest = "Result"
        if not os.path.exists(self.dest):
            os.makedirs(self.dest)

        if os.path.exists(self.outfile_name):
            self.wb1 = openpyxl.load_workbook(self.outfile_name)
        else:
            self.wb1 = openpyxl.Workbook()
        self.ws1 = self.wb1.active

        self.total_data.sort(key=takeFirst)

        if os.path.exists(self.outfile_name):
            self.wb1 = openpyxl.load_workbook(self.outfile_name)
        else:
            self.wb1 = openpyxl.Workbook()
        self.ws1 = self.wb1.active

        for i, line in enumerate(self.total_data):
            line = line[1:]
            print line
            for j, elm in enumerate(line):
                self.ws1.cell(row=i+2, column=j+1).value = elm

        self.wb1.save(self.outfile_name)


if __name__ == '__main__':
    start_t = time.time()
    app = mainProcesser()
    app.loadingData()
    app.totalProcessing()
    app.saveCSV()
    # app.processOneParagraph()

    print time.time() - start_t