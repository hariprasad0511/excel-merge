import pandas as pd
import logging
import os
import sys
from openpyxl.workbook import Workbook
from datetime import datetime

LOG_FILE = datetime.now().strftime('LogFile_%H_%M_%S_%d_%m_%Y.log')


logging.basicConfig(filename=LOG_FILE, level=logging.INFO, format="%(asctime)s:%(levelname)s:%(message)s")

class Merger():

    def __init__(self, filePath, outFileName):
        self.filePath = filePath
        self.outFileName = outFileName
        self.writer = pd.ExcelWriter(self.outFileName)
        logging.debug("The file path is : {}".format(self.filePath))     
        logging.info("The output file name is : {}".format(self.outFileName))
        

    def merger(self):
        logging.info("Implementing merger()...")
        logging.info("Switch to file path directory {}".format(self.filePath))
        try:
            os.chdir(self.filePath)
            logging.info("Listing files in directory {}".format(self.filePath))
        except:
            logging.error("The path {} does not exist".format(self.filePath))
            logging.error("Try with a different path")
            logging.error("Exiting...")
            sys.exit()
        fileList = os.listdir(os.getcwd())
        if len(fileList) != 0:
            logging.info("The list of files in the dorectory is : {}".format(fileList))
        else:
            logging.debug("The folder in path {} does not conatin any files".format(self.filePath))
            logging.debug("Try with a different path")
            logging.debug("Exiting...")
            sys.exit()
        for fileName in fileList:
            logging.info("Working with file : {}".format(fileName))
            dataFrame = pd.DataFrame(pd.read_excel(os.getcwd()+'\\'+fileName))
            logging.info("Started writing to file {} in the sheet {}...".format(self.outFileName,fileName))
            logging.info("Completed writing to file {} in sheet {}...".format(self.outFileName,fileName))
            dataFrame.to_excel(self.writer, sheet_name=fileName, index=False)
            self.writer.save()
            logging.info("Saving the file {}".format(self.outFileName))
        logging.info("Completed writing data to file {}".format(self.outFileName))
        logging.info("Closing the file {}".format(self.outFileName))
        self.writer.close()

if __name__ == '__main__':
    logging.info("Program started to merge the multiple excel files to same file with different sheet...")
    merger1 = Merger('./files','result.xlsx')
    merger1.merger()
    logging.info("Closing the program...")









