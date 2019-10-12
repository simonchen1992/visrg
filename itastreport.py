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
    templateFile = session['report_template']
    #templateFile = 'Blank_ Device_Cross_Testing_Sheet_20190604.xlsx'
    wb = load_workbook('docs/VisaTemplate/' + templateFile)
    wb['Cross_test_results']['B1'] = session['visa_vtf']  # add visa vtf
    # add testing date
    import MySQLdb
    cur = db.cursor(cursorclass=MySQLdb.cursors.DictCursor)
    cur.execute("SELECT created FROM test_cases WHERE id_test_session = " + str(args.keyword) + " ORDER BY created ASC LIMIT 1")
    startdate = cur.fetchallDict()
    startdate = str(startdate[0]['created']).split(' ')[0]
    cur.execute("SELECT created FROM test_cases WHERE id_test_session = " + str(args.keyword) + " ORDER BY created DESC LIMIT 1")
    enddate = cur.fetchallDict()
    enddate = str(enddate[0]['created']).split(' ')[0]
    for cell in wb['Cross_test_results']['A']:
      if cell.value == 'Cross testing start date':
        wb['Cross_test_results']['B' + str(cell.row)] = startdate  # add visa vtf
      if cell.value == 'Cross testing end date':
        wb['Cross_test_results']['B' + str(cell.row)] = enddate  # add visa vtf
    itast.db.export_results_to_excel(cards, wb, TX_LIMIT)
    wb.save(outFile)
  else:
    print Fore.RED + "No session found by ID " + args.keyword + Fore.RESET

# Always close the database connection before terminating
db.close()
