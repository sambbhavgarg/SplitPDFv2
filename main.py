from SplitPDF import SplitPDF
import json
from pprint import pprint

'''
data.xlsx needs to be in standard format -
  1. First Column: Full Name
  2. Second Column: First Name [excel formula to get first name from full name: =LEFT(A1, FIND(" ", A1) - 1)]
  3. Third Column: Email ID (optional)
  
The output will be erroneous if above standard is not met or corresponding changes arent made in SplitPDF.py
'''
   
with open("config.json") as f:
    config = json.load(f)

pprint(config)

sPDF = SplitPDF(config)
sPDF.splitPDF()