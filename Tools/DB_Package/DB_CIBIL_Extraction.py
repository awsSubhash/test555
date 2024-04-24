import camelot
import pandas as pd
import sys
import os
pdf_path=sys.argv[1]
df=pd.DataFrame()
tables2=camelot.read_pdf(pdf_path,flavour='lattice',pages='all')
for table in tables2:
    df=pd.concat([df,table.df])
a=os.path.dirname(pdf_path)
b=os.path.splitext(os.path.basename(pdf_path))[0]+".csv"
df.to_csv(os.path.join(a,b))