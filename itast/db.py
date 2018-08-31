import MySQLdb
from sys import stdout

# Utility for printing in colourful
from colorama import init
from colorama import Fore, Back, Style
init()

# Return values, and either to be considered PASS or FAIL
APPROVED_OFFLINE = '5931'
APPROVED_ONLINE = '3030'
DECLINED_OFFLINE = '5A31'
INCOMPLETED = '0000'
TERMINATED_ERROR_SW = 'EF01'
TERMINATED_INCORRECT_DATA = 'EF02'
TERMINATED_TIMEOUT = 'EF03'
TERMINATED_COLLISION = 'EF04'
TERMINATED_REQ_RESET = 'EF05'
TERMINATED_SWITCH_IFACE = 'EF06'
TERMINATED_REF_MOBILE = 'EF07'

TX_PASS = (APPROVED_ONLINE, APPROVED_OFFLINE)
TX_DEC = (DECLINED_OFFLINE)
TX_FAIL = (DECLINED_OFFLINE,
           INCOMPLETED,
           TERMINATED_ERROR_SW,
           TERMINATED_INCORRECT_DATA,
           TERMINATED_TIMEOUT,
           TERMINATED_COLLISION,
           TERMINATED_REQ_RESET,
           TERMINATED_SWITCH_IFACE,
           TERMINATED_REF_MOBILE)
TX_TF = (  TERMINATED_ERROR_SW,
           TERMINATED_INCORRECT_DATA,
           TERMINATED_SWITCH_IFACE)
TX_CF = (  TERMINATED_TIMEOUT,
           TERMINATED_COLLISION,
           TERMINATED_REQ_RESET)
TX_UC = (INCOMPLETED)


def connect():
  db = MySQLdb.connect(user='itastdirector',
                       passwd='APtsd01$',
                       host='192.168.48.60',
                       db='itastdb')
  return db

def print_table(fields, cursor):
  for f in fields:
    stdout.write(Style.BRIGHT + f.ljust(fields[f]+2).capitalize()[:fields[f]+2] + Style.RESET_ALL)
  print ""
  for row in cursor.fetchallDict():
    for f in fields:
      stdout.write(str(row[f]).strip().ljust(fields[f]+2)[:fields[f]+2])
    print ""

def print_sessions(db, keyword):
  fields = {}
  fields['id'] = 8
  fields['created'] = 20
  fields['expedient'] = 32
  fields['owner'] = 16
  fields['dut1_name'] = 16
  fields['dut1_description'] = 24
  query_fileds = ','.join(fields.keys())
  cur = db.cursor(cursorclass=MySQLdb.cursors.DictCursor)
  cur.execute("SELECT " + query_fileds + " FROM test_sessions " +
              "WHERE expedient LIKE '%" + keyword + "%' " +
              "OR id LIKE '%" + keyword + "%' " +
              "OR dut1_id LIKE '%" + keyword + "%' " +
              "OR dut1_name LIKE '%" + keyword + "%' " +
              "OR dut1_description LIKE '%" + keyword + "%' " +
              "OR dut2_id LIKE '%" + keyword + "%' " +
              "OR dut2_name LIKE '%" + keyword + "%' " +
              "OR dut2_description LIKE '%" + keyword + "%'" +
              "ORDER BY created DESC LIMIT 10;")
  print_table(fields, cur)

def set_position_verdict(cardpostx):
  txPass = 0
  txFail = 0
  txCF = 0
  txTF = 0
  txUC = 0
  txDEC = 0
  txTotal = 0
  for tx in cardpostx['txs']:
    txTotal =+ txTotal + 1
    if cardpostx['txs'][tx] in TX_PASS:
      txPass = txPass + 1
    #elif cardpostx['txs'][tx] in TX_FAIL:
    #  txFail =+ txFail + 1
    elif cardpostx['txs'][tx] in TX_CF:
      txCF = txCF + 1
    elif cardpostx['txs'][tx] in TX_TF:
      txTF = txTF + 1
    elif cardpostx['txs'][tx] in TX_UC:
      txUC = txUC + 1
    else:
      print "What the hell is this result: " + cardpostx['txs'][tx]

  if txPass >= 5:
    cardpostx['verdict'] = "OK"
  elif txCF + txTF > 0 and txPass > 0:
    cardpostx['verdict'] = "DF"
  elif txCF >= 3:
    cardpostx['verdict'] = "CF"
  elif txCF < 3 and txCF + txTF >=3:
    cardpostx['verdict'] = "TF"
  elif txTotal == 0:
    cardpostx['verdict'] = "  "
  else:
    cardpostx['verdict'] = "KO"

def load_session_data(db, idsession, txlimit=7):
  session = {}
  cur = db.cursor(cursorclass=MySQLdb.cursors.DictCursor)
  # Get data from DB
  cur.execute("SELECT * FROM test_sessions WHERE id = " + idsession + ";")
  session = cur.fetchoneDict()
  return session

def load_card_results_by_session(db, idsession, txlimit=1):
  visaCards = {}
  cur = db.cursor(cursorclass=MySQLdb.cursors.DictCursor)
  # Get Visa Cards from DB
  cur.execute("SELECT * FROM visa_cards WHERE active = 1 ORDER BY id ASC;")
  visaCards = cur.fetchallDict()
  # For each card in Visa card deck, get the most recent 10 txs from DB
  for card in visaCards:
    card['txs'] = {}
    for pos in getPositions():
      card['txs'][pos] = {}
      card['txs'][pos]['txs'] = {}
      card['txs'][pos]['verdict'] = ""
      card['txs'][pos]['comment'] = ""
      cur.execute("SELECT verdict,comments FROM test_cases " +
                  "WHERE id_test_session = " + str(idsession) + " AND dut = '1' AND id_card = " + str(card['id']) + " AND pos = '" + pos + "' " +
                  "ORDER BY created DESC LIMIT 1")
      for row in cur.fetchallDict():
        card['txs'][pos]['verdict'] = row['verdict']
        card['txs'][pos]['comment'] = row['comments']
      card['txs'][pos]['count'] = len(card['txs'][pos]['txs'])
      
  return visaCards

def export_results_to_excel(cards, wb, txlimit=1):
    wsRoadmap = wb.create_sheet(title="Roadmap")
    wsRoadmap.append(['Card VTF','Verdict'] + getPositions() + ['Comment'])
    for card in cards:
      cardata = []
      cardata.append(card['vtf'].strip())
      cardata.append('XX')
      for pos in getPositions():
        cardata.append(str(card['txs'][pos]['verdict']))
      
      cardata.append(str(card['txs']['0N']['comment']))
      wsRoadmap.append(cardata)

def getPositions():
  positions = []
  for Z in ['0','1','2','3','4']: # Iterate Z
    for XY in ['N', 'S', 'E', 'W', 'C']: # Iterate cardinals
      positions.append(Z+XY)
  return positions

def print_xtroadmap(cards):
  for card in cards:
    stdout.write(card['vtf'].strip() + "\t")
    for pos in getPositions():
      stdout.write(str(card['txs'][pos]['verdict']) + " ")
    print ""










