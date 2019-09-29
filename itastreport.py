#!/usr/bin/env python
import itast.settings
import itast.db

# Utility to parse command line arguments
import argparse
parser = argparse.ArgumentParser()
parser.add_argument("command", help="command to be executed", choices=['search','analyse','report'], nargs=1)
parser.add_argument("keyword", help="keyword to session search, or ID to analyse or report")
args = parser.parse_args()

# Utility for printing in colourful
from colorama import init
from colorama import Fore, Back, Style
init()

# Connect to the ITAST Database and hold it in db variable
db = itast.db.connect()

TX_LIMIT = 7 # Get the most recent TX_LIMIT transaction of each card/position

if args.command[0] == 'search':
  print "Searching sessions with the keyword " + Style.BRIGHT + args.keyword + Style.RESET_ALL + "..."
  print ""
  itast.db.print_sessions(db, args.keyword)

elif args.command[0] == 'analyse':
  session = itast.db.load_session_data(db, args.keyword)
  if session:
    print "Analysing test session " + Style.BRIGHT + str(session['id']) + " - " + str(session['expedient']).strip() + Style.RESET_ALL
    print ""
    cards = itast.db.load_card_results_by_session(db, args.keyword, TX_LIMIT)
    itast.db.print_xtroadmap(cards)
  else:
    print Fore.RED + "No session found by ID " + args.keyword + Fore.RESET


elif args.command[0] == 'report':
  session = itast.db.load_session_data(db, args.keyword)
  if session:
    if len(session['expedient']) > 3:
      expedient = session['expedient'].strip()
    else:
      expedient = str(session['id'])
    outFile = "Report_" + expedient + "_" + str(session['id']) + ".xlsx"
    print "Writing report of session " + str(session['id']) + " to file: " + Style.BRIGHT + outFile + Style.RESET_ALL
    
    cards = itast.db.load_card_results_by_session(db, args.keyword, TX_LIMIT)

    from openpyxl import load_workbook
    templateFile = 'Blank_ Device_Cross_Testing_Sheet_20190604.xlsx'
    wb = load_workbook(templateFile)
    wb['Cross_test_results']['B1'] = session['visa_vtf']  # add visa vtf
    itast.db.export_results_to_excel(cards, wb, TX_LIMIT)
    wb.save(outFile)
  else:
    print Fore.RED + "No session found by ID " + args.keyword + Fore.RESET

# Always close the database connection before terminating
db.close()
