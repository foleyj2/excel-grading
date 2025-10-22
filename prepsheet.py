#!/usr/bin/env python3
## By Joseph T. Foley <foley AT RU.IS>
## Created 2025-10-22
"""Prepare an excel grading template for multiple students"""
import os
import sys
from pathlib import PurePath##https://docs.python.org/3/library/pathlib.html#module-pathlib
import argparse
import logging
import openpyxl
##sudo dnf5 install python3-openpyxl

class SheetPrepper():
    """XLSX sheet preparation"""
    def __init__(self, infd, logger):
        self.logger = logger       


def main():
    """Main program loop"""
    print("prepsheet by Joseph. T. Foley<foley AT ru DOT is>")
    parser = argparse.ArgumentParser(
        description="Generate xlsx grading sheet from template")
    parser.add_argument('template',
                        help="Excel spreadsheet template .xlsx")
    parser.add_argument('studentnames',
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
    logpath = os.path.splitext(args.filename)[0]+".log"
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

    logger.info("Creating Peerwise2TeX log file %s", logpath)

    # filename pre-processing for output
    inpath = PurePath(args.filename)
    print(f"Input: {inpath}")
    input = open(inpath, "r")
    SP = SheetPrepper(input, logger)
    outstem = inpath.stem+"-"+tag
    logger.info("Output: %s.pdf" % outstem)

    
if __name__ == "__main__":
  main()
