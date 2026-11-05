import os,sys; sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)),'charts.py')).read().split('# C1 10Y history')[0])
import numpy as np
lad=np.interp(range(1,11),[1,2,3,5,7,10],[u['1 Yr'],u['2 Yr'],u['3 Yr'],u['5 Yr'],u['7 Yr'],u['10 Yr']]).mean()
hcv=(sofr-tona).dropna().iloc[-1]; jl=np.mean([J.iloc[-1][f'{k}Y'] for k in range(1,11)]); spot=fx.iloc[-1]
x=np.linspace(-15,8,93)
fig,ax=figax()
for h,c,lab,ls in [(0,O,'ヘッジなし（0%）','-'),(0.5,MUT,'50%固定ヘッジ',(0,(4,2))),(1,B,'フルヘッジ（100%）','-')]:
    y=lad-h*hcv+(1-h)*x; ax.plot(x,y,color=c,ls=ls)
    ax.annotate(f'{lab}',(x[-1],y[-1]),xytext=(8,-9 if h==1 else 0),textcoords='offset points',color=INK,fontsize=10,va='center')
ax.axhline(jl,color=A,lw=1.6); ax.annotate(f'日本国債ラダー {jl:.2f}%',(x[-1],jl),xytext=(8,9),textcoords='offset points',color=INK,fontsize=10,va='center')
ax.axhline(0,color=AX,lw=0.8); ax.axvline(0,color=AX,lw=0.8)
be=-hcv; ax.scatter([be],[lad-hcv],s=40,color=INK,zorder=5)
ax.annotate(f'全戦略の損益分岐：円高 {hcv:.1f}%\n（USD/JPY {spot:.1f}→約{spot*(1-hcv/100):.0f}円）',(be,lad-hcv),xytext=(40,-90),textcoords='offset points',fontsize=9.5,color=SEC,arrowprops=dict(arrowstyle='-',color=MUT,lw=0.8))
ax.set_xlabel('1年間のUSD/JPY変化率（%、マイナス＝円高）'); ax.set_ylabel('円建てリターン（%、1年・概算）'); ax.set_ylim(-12,14)
fin(fig,ax,'ヘッジ比率の選択は「利回り」と「円高耐性」のトレードオフ',
 f'米国債ラダー参考利回り{lad:.2f}%、ヘッジコスト目安{hcv:.2f}%で一定、債券価格変動・費用なしとした1年間の簡易試算',
 '出所：米国財務省、財務省（日本）、FRB（FRED SOFR, DEXJPUS）、日本銀行より当社作成（2026/11/4前後のデータ）。説明用の概算であり、将来の運用成果を示すものではありません。','c8_hedge_ratio_payoff.png')
print(lad,hcv,jl,spot,spot*(1-hcv/100))
