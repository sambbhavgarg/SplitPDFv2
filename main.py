from SplitPDF import SplitPDF
import json
from pprint import pprint
   
with open("config.json") as f:
    config = json.load(f)

pprint(config)

sPDF = SplitPDF(config)
sPDF.splitPDF()
