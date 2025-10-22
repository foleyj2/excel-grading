#!/usr/bin/env python3
## By Joseph T. Foley <foley AT RU.IS>
## Created 2025-10-22
"""Prepare an excel grading template for multiple students"""
import os
import sys
import shutil
from pathlib import PurePath##https://docs.python.org/3/library/pathlib.html#module-pathlib
import argparse
import logging
import openpyxl
import copy
##sudo dnf5 install python3-openpyxl

class ExcelMerge():
    """XLSX sheet merge"""
    def __init__(self, template, data, logger):
        # first we copy the template to the output
        inpath = PurePath(template)
        outpath = inpath.stem+"-merged"+inpath.suffix
        shutil.copy2(template, outpath)
        logger.info(f"Output: {outpath}")
        self.template = template
        self.data = data
        self.logger = logger
        self.outpath = outpath
        self.workbook = openpyxl.load_workbook(filename=outpath)

    def dupsheet(self, name):
        "Assume first sheet is template, copy to the end.  Return sheet"
        workbook = self.workbook
        template_sheet = workbook.active
        ws = workbook.copy_worksheet(template_sheet)
        ws.title = name
        return ws

    def insertname(self, name, cell="B2"):
        "put name in the cell"
        # assume active sheet is the right one
        ws = self.workbook.active
        ws[cell] = name

    def merge(self):
        "do the merge!"
        with open(self.data, "r") as datafile:
            for line in datafile:
                line = line.strip()
                self.logger.info(f"Creating sheet {line}")
                self.dupsheet(line)
                self.insertname(line)
            
    
    def save(self):
        "save the output"
        self.logger.debug(f"Saving to {self.outpath}")
        self.workbook.save(self.outpath)

def main():
    """Main program loop"""
    print("prepsheet by Joseph. T. Foley<foley AT ru DOT is>")
    parser = argparse.ArgumentParser(
        description="Generate xlsx grading sheet from template")
    parser.add_argument('template',
                        help="Excel spreadsheet template .xlsx")
    parser.add_argument('data',
                        help="text file with student names")
    ## TODO:  Fix verbosity based upon another argument and logging
    parser.add_argument('--log', default="INFO",
        help='Console log level:  Number or DEBUG, INFO, WARNING, ERROR')

    args = parser.parse_args()
    ## Set up logging
    numeric_level = getattr(logging, args.log.upper(), None)
    if not isinstance(numeric_level, int):
        raise ValueError('Invalid log level: %s' % args.log)
    #print(f"Log level:  {numeric_level}")
    logger = logging.getLogger("app")
    logger.setLevel(numeric_level)
    # log everything to file
    logpath = os.path.splitext(args.template)[0]+".log"
    fh = logging.FileHandler(logpath)
    fh.setLevel(logging.DEBUG)
    # log to console
    ch = logging.StreamHandler()
    ch.setLevel(numeric_level)
    # create formatter and add to handlers
    consoleformatter = logging.Formatter('%(message)s')
    ch.setFormatter(consoleformatter)
    spamformatter = logging.Formatter('%(asctime)s %(name)s[%(levelname)s] %(message)s')
    fh.setFormatter(spamformatter)
    # add the handlers to logger
    logger.addHandler(ch)
    logger.addHandler(fh)

    logger.info("Creating prepsheet log file %s", logpath)

    # filename pre-processing for output
    SP = ExcelMerge(args.template, args.data, logger)
    SP.merge()
    SP.save()
    
if __name__ == "__main__":
  main()
