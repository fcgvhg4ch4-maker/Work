import pandas as pd, io
import os
R=os.path.join(os.path.dirname(os.path.abspath(__file__)),'..','raw')+'/'
def fred(i):
    d=pd.read_csv(R+i+'.csv'); d.columns=['date',i]; d['date']=pd.to_datetime(d['date']); d[i]=pd.to_numeric(d[i],errors='coerce'); return d.set_index('date')[i].dropna()
def boj(f,name):
    rows=[l.split(',') for l in open(R+f,encoding='utf8') if l.startswith(('STRD','FXERD'))]
    s=pd.Series({pd.to_datetime(r[-2]):r[-1].strip() for r in rows}); s=pd.to_numeric(s,errors='coerce').dropna(); s.name=name; return s
def jgb():
    import re
    t=open(R+'jgbcm_all.csv','rb').read().decode('sjis').splitlines()+open(R+'jgbcm.csv','rb').read().decode('sjis').splitlines()
    rows=[]
    era={'S':1925,'H':1988,'R':2018}
    for l in t:
        m=re.match(r'([SHR])(\d+)\.(\d+)\.(\d+),(.*)',l)
        if not m: continue
        y=era[m[1]]+int(m[2]); vals=m[5].split(',')
        rows.append([pd.Timestamp(y,int(m[3]),int(m[4]))]+[pd.to_numeric(v,errors='coerce') for v in vals[:15]])
    cols=['date']+[f'{x}Y' for x in [1,2,3,4,5,6,7,8,9,10,15,20,25,30,40]]
    return pd.DataFrame(rows,columns=cols).drop_duplicates('date').set_index('date').sort_index()
def ust():
    d=pd.read_csv(R+'ust2026.csv'); d['Date']=pd.to_datetime(d['Date']); return d.set_index('Date').sort_index()
