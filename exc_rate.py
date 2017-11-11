# Reference: https://pypi.python.org/pypi/forex-python
print "Exchange rate converter"

import sys
print (sys.argv)
print "len=%s" %(len(sys.argv))
if (len(sys.argv) == 4):
  from_cur = sys.argv[1]
  to_cur = sys.argv[2]
  amount=sys.argv[3]
else:
  print "Error"
  print "Use %s from_cur to_cur amount" %(sys.argv[0])
  sys.exit()

from decimal import *
from forex_python.converter import CurrencyRates
c = CurrencyRates()

from forex_python.converter import CurrencyCodes
code = CurrencyCodes()
print "[%s %s] converted to [%s] are %s %s" %(amount,from_cur,to_cur,code.get_symbol(to_cur),c.convert(from_cur, to_cur, Decimal(amount)))
