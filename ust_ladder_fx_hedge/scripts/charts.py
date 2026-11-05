import os,sys; sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
import os; 
from load import *
import numpy as np, matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt, matplotlib.dates as mdates
from matplotlib import font_manager as fm
fm.fontManager.addfont('/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf')
B,O,A='#2a78d6','#eb6834','#1baf7a'; INK,SEC,MUT,GRID,AX,SURF='#0b0b0b','#52514e','#898781','#e1e0d9','#c3c2b7','#fcfcfb'
plt.rcParams.update({'font.family':'IPAGothic','font.size':11,'axes.edgecolor':AX,'axes.labelcolor':SEC,'xtick.color':SEC,'ytick.color':SEC,
 'axes.grid':True,'grid.color':GRID,'grid.linewidth':0.6,'axes.spines.top':False,'axes.spines.right':False,'axes.spines.left':False,
 'figure.facecolor':SURF,'axes.facecolor':SURF,'lines.linewidth':2,'legend.frameon':False,'xtick.major.size':0,'ytick.major.size':0})
OUT='/home/user/Work/ust_ladder_fx_hedge/charts/'
def fin(fig,ax,title,sub,src,name):
    fig.text(0.012,0.965,title,fontsize=15,color=INK,weight='bold',va='top')
    fig.text(0.012,0.905,sub,fontsize=10.5,color=SEC,va='top')
    fig.text(0.012,0.02,src,fontsize=8.2,color=MUT,va='bottom')
    fig.savefig(OUT+name,dpi=200); plt.close(fig)
def figax(): 
    fig,ax=plt.subplots(figsize=(11,5.6)); fig.subplots_adjust(left=0.06,right=0.86,top=0.82,bottom=0.15); return fig,ax

U=ust(); u=U.iloc[-1]; ud=U.index[-1]
d10=fred('DGS10'); d2=fred('DGS2'); fx=fred('DEXJPUS'); dff=fred('DFF'); sofr=fred('SOFR')
tona=boj('tona_all.csv','tona'); J=jgb(); ub=fred('DFEDTARU'); lb=fred('DFEDTARL')

# C1 10Y history
fig,ax=figax(); s=d10['2000':]
ax.plot(s.index,s.values,color=B,lw=1.4)
lv=s.iloc[-1]; ax.axhline(lv,color=MUT,lw=0.8,ls=(0,(4,3)))
ax.scatter([s.index[-1]],[lv],s=40,color=B,zorder=5,edgecolor=SURF,linewidth=1.5)
ax.annotate(f'{s.index[-1]:%Y/%m/%d}\n{lv:.2f}%',(s.index[-1],lv),xytext=(10,-4),textcoords='offset points',color=INK,fontsize=10,va='center')
pk=d10['2008':].idxmax(); ax.annotate(f'2008年以降の最高 {d10[pk]:.2f}%\n({pk:%Y/%m/%d})',(pk,d10[pk]),xytext=(-150,8),textcoords='offset points',color=SEC,fontsize=9.5,arrowprops=dict(arrowstyle='-',color=MUT,lw=0.8))
pct=(d10['2008':]<lv).mean()*100
ax.set_ylabel('%'); ax.set_ylim(0,7.2); ax.xaxis.set_major_locator(mdates.YearLocator(4)); ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y'))
fin(fig,ax,'米国10年国債利回りは、2008年以降の上位約1%の水準',f'米国10年国債利回り（日次、%）2000/1/3〜{s.index[-1]:%Y/%m/%d}。直近値は2008年以降の日次データの{pct:.0f}%タイル',
 '出所：米連邦準備制度理事会（FRB）H.15（FRED系列 DGS10、取得日2026/11/5）。過去の利回り水準は将来の利回りを示唆するものではありません。','c1_ust10y_history.png')

# C2 curves
fig,ax=figax()
cols=['1 Yr','2 Yr','3 Yr','5 Yr','7 Yr','10 Yr','20 Yr','30 Yr']; mats=[1,2,3,5,7,10,20,30]
U25=pd.read_csv(R+'ust2025.csv'); U25['Date']=pd.to_datetime(U25['Date']); u25=U25.set_index('Date').loc['2025-11-04']
jt=J.iloc[-1]; jd=J.index[-1]; jm=[1,2,3,5,7,10,20,30]; jv=[jt[f'{m}Y'] for m in jm]
for v,c,lab,ls in [([u[c] for c in cols],B,f'米国債 {ud:%Y/%m/%d}','-'),([u25[c] for c in cols],B,'米国債 1年前\n（2025/11/04）',(0,(3,2))),(jv,O,f'日本国債 {jd:%Y/%m/%d}','-')]:
    ax.plot(mats,v,color=c,ls=ls,marker='o',ms=5,mec=SURF,mew=1.2,alpha=1 if ls=='-' else 0.6)
    ax.annotate(lab,(30,v[-1]),xytext=(8,0),textcoords='offset points',color=INK,fontsize=9.5,va='center')
ax.axvspan(1,10,color=B,alpha=0.06,lw=0); ax.text(5.5,0.35,'ラダー想定レンジ（1〜10年）※構成は要確認',ha='center',color=SEC,fontsize=9.5)
ax.set_xscale('log'); ax.set_xticks(mats); ax.set_xticklabels([f'{m}年' for m in mats]); ax.set_ylim(0,6); ax.set_ylabel('%'); ax.set_xlim(0.9,30)
fin(fig,ax,'米国債は1〜10年が緩やかな右上がり。1年前比で1年超の全年限が上昇',
 '米国債パー・イールド（%）と日本国債金利（%）。横軸は残存年数（対数目盛）',
 '出所：米国財務省 Daily Treasury Par Yield Curve Rates、財務省（日本）国債金利情報（取得日2026/11/5）。','c2_yield_curves.png')

# C3 policy rates
fig,ax=figax(); s0='2021-01-01'
mid=((ub+lb)/2)[s0:]; ax.fill_between(ub[s0:].index,lb[s0:].reindex(ub[s0:].index).values,ub[s0:].values,color=B,alpha=0.25,lw=0,step='post')
ax.plot(mid.index,mid.values,color=B,lw=1.6,drawstyle='steps-post')
tm=tona[s0:].resample('W').mean().dropna(); ax.plot(tm.index,tm.values,color=O,lw=1.8)
sep={2026:4.125,2027:4.125,2028:3.875}  # plotted as reported medians below
med={'2026年末':4.1,'2027年末':4.1,'2028年末':3.9}
xs=[pd.Timestamp('2026-12-31'),pd.Timestamp('2027-12-31'),pd.Timestamp('2028-12-31')]
ax.scatter(xs,list(med.values()),marker='D',s=46,color=B,edgecolor=SURF,linewidth=1.5,zorder=5)
for x,(k,v) in zip(xs,med.items()): ax.annotate(f'{v:.1f}',(x,v),xytext=(0,9),textcoords='offset points',ha='center',fontsize=9.5,color=INK)
ax.text(pd.Timestamp('2023-08-01'),5.75,'米FF金利誘導目標（レンジ・中心）',color=INK,fontsize=10)
ax.annotate('FOMC参加者見通し中央値\n（2026年9月SEP）',(xs[1],4.1),xytext=(-30,-55),textcoords='offset points',color=SEC,fontsize=9.5)
ax.annotate(f'日銀：無担保コール翌日物（週平均）\n直近 {tona.dropna().iloc[-1]:.3f}%（政策金利1.25%程度）',(pd.Timestamp('2024-01-15'),1.05),color=INK,fontsize=10)
ax.set_ylim(-0.5,6); ax.set_ylabel('%'); ax.set_xlim(pd.Timestamp(s0),pd.Timestamp('2029-03-31'))
ax.xaxis.set_major_locator(mdates.YearLocator()); ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y'))
fin(fig,ax,'日米とも利上げ局面。ただし政策金利差は約2.6%と依然大きい',
 '米FF金利誘導目標レンジ（%）、日本の無担保コール翌日物（%）、FOMC参加者のFF金利見通し中央値（年末、%）',
 '出所：FRB（FRED系列 DFEDTARU/DFEDTARL）、FRB Summary of Economic Projections（2026/9/16）、日本銀行 時系列統計（STRDCLUCON）。見通しは予測であり実現を保証しません。','c3_policy_rates.png')

# C4 hedge cost proxy history
fig,ax=figax()
hc=(dff.resample('ME').mean()-tona.resample('ME').mean()).dropna()['2005':]
ax.fill_between(hc.index,0,hc.values,color=B,alpha=0.18,lw=0); ax.plot(hc.index,hc.values,color=B,lw=1.6)
ax.annotate(f'{hc.index[-1]:%Y/%m} {hc.iloc[-1]:.2f}%',(hc.index[-1],hc.iloc[-1]),xytext=(8,0),textcoords='offset points',color=INK,fontsize=10,va='center')
pk=hc.idxmax(); ax.annotate(f'ピーク {hc[pk]:.2f}%（{pk:%Y/%m}）',(pk,hc[pk]),xytext=(-60,10),textcoords='offset points',color=SEC,fontsize=9.5)
ax.axhline(0,color=AX,lw=0.8); ax.set_ylabel('%（年率）'); ax.set_ylim(-0.3,6)
ax.xaxis.set_major_locator(mdates.YearLocator(3)); ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y'))
fin(fig,ax,'為替ヘッジコストの目安（日米短期金利差）は2023年のピークから低下も、なお2%台後半',
 '米実効FF金利 − 日本 無担保コール翌日物（月平均、%）。実際のヘッジコストはこれにベーシス等が加わる（参考指標）',
 '出所：FRB（FRED系列 DFF）、日本銀行（STRDCLUCON）より当社作成。為替フォワードから算出される実際のヘッジコストとは一致しません。','c4_hedge_cost_proxy.png')

# C5 yield bars
lad=np.interp(range(1,11),[1,2,3,5,7,10],[u['1 Yr'],u['2 Yr'],u['3 Yr'],u['5 Yr'],u['7 Yr'],u['10 Yr']]).mean()
jl=np.mean([jt[f'{k}Y'] for k in range(1,11)])
hcost=(sofr-tona).dropna(); hcv=hcost.iloc[-1]; hcd=hcost.index[-1]
fig,ax=plt.subplots(figsize=(11,5.6)); fig.subplots_adjust(left=0.06,right=0.97,top=0.82,bottom=0.18)
labs=['米国債ラダー\n（1〜10年・均等）','ヘッジコスト目安\n（SOFR−TONA）','フルヘッジ後\n参考利回り','日本国債ラダー\n（1〜10年・均等）']
vals=[lad,-hcv,lad-hcv,jl]; bottoms=[0,lad,0,0]; heights=[lad,-hcv,lad-hcv,jl]
cols_=[B,MUT,B,O]
for i,(b,h,c) in enumerate(zip([0,lad-hcv,0,0],[lad,hcv,lad-hcv,jl],cols_)):
    ax.bar(i,h,bottom=b,color=c,width=0.55,alpha=0.55 if i==2 else 1,edgecolor=SURF,linewidth=2)
    txt=f'{vals[i]:+.2f}%' if i==1 else f'{vals[i]:.2f}%'
    ax.text(i,b+h+0.08,txt,ha='center',color=INK,fontsize=12,weight='bold')
ax.set_xticks(range(4)); ax.set_xticklabels(labs,color=INK); ax.set_ylim(0,5.6); ax.set_ylabel('%（年率）'); ax.grid(axis='x',visible=False)
fin(fig,ax,'フルヘッジ後の米国債ラダー参考利回りは、同年限の日本国債を下回る',
 f'米国債・日本国債は{ud:%Y/%m/%d}時点のパー利回り／複利利回り（年限補間・単純平均）、ヘッジコスト目安は{hcd:%Y/%m/%d}時点の翌日物金利差',
 '出所：米国財務省、財務省（日本）、FRB（FRED系列 SOFR）、日本銀行より当社作成。ラダー構成は仮定（要確認）。ヘッジコストは参考値で、実際の為替フォワード・ベーシス・取引コスト・信託報酬は未反映。','c5_yield_comparison.png')

# C6 hedged 10y vs JGB10
fig,ax=figax()
m10=d10.resample('ME').mean(); h10=(m10-hc).dropna()['2010':]; j10=J['10Y'].resample('ME').mean()['2010':]
ax.plot(h10.index,h10.values,color=B); ax.plot(j10.index,j10.values,color=O); ax.axhline(0,color=AX,lw=0.8)
ax.annotate(f'米国10年債\n（フルヘッジ後目安）\n{h10.iloc[-1]:.2f}%',(h10.index[-1],h10.iloc[-1]),xytext=(8,-6),textcoords='offset points',color=INK,fontsize=10,va='center')
ax.annotate(f'日本10年国債 {j10.iloc[-1]:.2f}%',(j10.index[-1],j10.iloc[-1]),xytext=(8,8),textcoords='offset points',color=INK,fontsize=10,va='center')
ax.set_ylabel('%'); ax.xaxis.set_major_locator(mdates.YearLocator(2)); ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y'))
fin(fig,ax,'2022年11月以降、フルヘッジ米国債は日本国債を下回り続ける（差は縮小傾向）',
 '米国10年国債利回り −（米実効FF金利 − 無担保コール翌日物） と 日本10年国債利回り（月平均、%）',
 '出所：FRB（FRED系列 DGS10, DFF）、日本銀行（STRDCLUCON）、財務省（日本）国債金利情報より当社作成。ヘッジコストは短期金利差による近似。','c6_hedged_vs_jgb.png')

# C7 USDJPY with interventions
fig,ax=figax(); s=fx['2021-07-01':]
ax.plot(s.index,s.values,color=B,lw=1.4)
iv=[('2022-09-22',28382),('2022-10-21',56202),('2022-10-24',7296),('2024-04-29',59185),('2024-05-01',38700),('2024-07-11',31678),('2024-07-12',23670),('2026-04-30',62787),('2026-05-04',7802),('2026-05-06',46759)]
for d,a in iv:
    d=pd.Timestamp(d); y=s[:d].iloc[-1] if d>=s.index[0] else None
    ax.axvline(d,color=O,lw=0.9,alpha=0.7)
ax.axvspan(pd.Timestamp('2026-07-30'),pd.Timestamp('2026-08-26'),color=O,alpha=0.12,lw=0)
ax.text(pd.Timestamp('2026-08-12'),112,'2026/7/30〜8/26\n介入総額15.4兆円\n（日次は11/9公表予定）',ha='center',fontsize=8.6,color=SEC)
for d,t in [('2022-10-01','2022/9〜10月\n計9.2兆円'),('2024-06-01','2024/4〜7月\n計15.3兆円'),('2026-05-03','2026/4〜5月\n計11.7兆円')]:
    ax.text(pd.Timestamp(d),127,t,ha='center',fontsize=8.6,color=SEC)
ax.annotate(f'{s.index[-1]:%Y/%m/%d} {s.iloc[-1]:.2f}',(s.index[-1],s.iloc[-1]),xytext=(8,0),textcoords='offset points',fontsize=10,color=INK,va='center')
mx=s.idxmax(); ax.annotate(f'{s[mx]:.2f}（{mx:%Y/%m/%d}）\n1986年以来の円安水準',(mx,s[mx]),xytext=(-170,4),textcoords='offset points',fontsize=9.5,color=SEC)
ax.plot([],[],color=O,lw=1,label='円買い介入実施日（財務省）'); ax.plot([],[],color=B,label='USD/JPY'); ax.legend(loc='upper left',fontsize=9.5)
ax.set_ylim(105,170); ax.set_ylabel('円/ドル'); ax.xaxis.set_major_locator(mdates.YearLocator()); ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y'))
fin(fig,ax,'円買い介入は2022年以降4局面。介入後の値動きは局面ごとに異なる',
 'USD/JPY（NY正午買値、日次）と財務省公表の円買い介入実施日（縦線）',
 '出所：FRB H.10（FRED系列 DEXJPUS）、財務省「外国為替平衡操作の実施状況」（日次：〜2026年6月期、月次：〜2026/10/28）。','c7_usdjpy_interventions.png')
print('done', lad, jl, hcv, hcd)
