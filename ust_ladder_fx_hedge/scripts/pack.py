import os,sys
sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
from load import *
import numpy as np
fx=fred('DEXJPUS'); d10=fred('DGS10'); d2=fred('DGS2'); dff=fred('DFF'); sofr=fred('SOFR'); tona=boj('tona_all.csv','tona'); J=jgb(); U=ust()
ub=fred('DFEDTARU'); lb=fred('DFEDTARL')
iv=[('2022-09-22',28382),('2022-10-21',56202),('2022-10-24',7296),('2024-04-29',59185),('2024-05-01',38700),('2024-07-11',31678),('2024-07-12',23670),('2026-04-30',62787),('2026-05-04',7802),('2026-05-06',46759)]
ivd=pd.DataFrame(iv,columns=['実施日','金額（億円）']); ivd['実施日']=pd.to_datetime(ivd['実施日']); ivd['内容']='米ドル売り・日本円買い'
ivd['出所']='財務省「外国為替平衡操作の実施状況」過去の介入実績全データ（平成3年4月〜令和8年6月期）'
ivm=pd.DataFrame([['2026-04-28〜2026-05-27',117349],['2026-05-28〜2026-06-26',0],['2026-06-29〜2026-07-29',0],['2026-07-30〜2026-08-26',153993],['2026-08-27〜2026-09-28',0],['2026-09-29〜2026-10-28',0]],columns=['対象期間','介入額（億円）'])
ivm['備考']=['日次内訳は上表（4/30, 5/4, 5/6）','','','日次内訳（実施日・金額）は2026/11/9 8:50公表予定 → 要確認','','']
ivm['出所']='財務省「外国為替平衡操作の実施状況（月次ベース）」'
ev=[('A','2007/6〜2008/3','円キャリー取引巻き戻し（サブプライム問題）','2007-06-01','2008-03-31','円高（介入なし）'),
('B','2008/8〜2008/12','リーマン・ショック','2008-08-01','2008-12-31','円高（介入なし）'),
('C','2016/1〜2016/6','日銀マイナス金利導入〜英EU離脱投票','2016-01-01','2016-06-30','円高（介入なし）'),
('D','2022/9〜2022/10','円買い介入（2022年9/22・10/21・10/24）','2022-09-01','2022-10-31','介入局面：効果が一時的'),
('E','2022/10〜2023/1','米インフレ鈍化・日銀YCC修正（2022/12）','2022-10-20','2023-01-31','介入後の円高トレンド'),
('F','2024/4〜2024/5','円買い介入（2024年4/29・5/1）','2024-04-15','2024-05-31','介入局面：効果が一時的'),
('G','2024/7〜2024/9','円買い介入（7/11・12）・日銀利上げ（7/31）・キャリー巻き戻し','2024-07-01','2024-09-30','介入＋急激な円高'),
('H','2026/4〜2026/6','円買い介入（2026年4/30・5/4・5/6）','2026-04-15','2026-06-30','介入局面：その後円安再開'),
('I','2026/7〜2026/8','円買い介入（7/30〜8/26の期間中、日次は未公表）','2026-07-15','2026-08-31','介入局面（日付要確認）'),
('J','2026/9','急激な円高（当該月次期間の介入実績は0円）','2026-08-27','2026-09-30','介入なしの急変')]
rows=[]
for k,lab,bg,s,e,typ in ev:
    w=fx[s:e]; dd=w/w.cummax()-1; tr=dd.idxmin(); pk=w[:tr].idxmax()
    ivs=ivd[(ivd['実施日']>=s)&(ivd['実施日']<=e)]
    rows.append([k,lab,bg,typ,w.index[0].date(),w.iloc[0],w.index[-1].date(),w.iloc[-1],round((w.iloc[-1]/w.iloc[0]-1)*100,1),pk.date(),w[pk],tr.date(),w[tr],round((w[tr]/w[pk]-1)*100,1),
      round(d10[s:e].iloc[-1]-d10[s:e].iloc[0],2), ', '.join(ivs['実施日'].dt.strftime('%Y/%m/%d')) or ('期間内に実施（日次未公表）' if k=='I' else 'なし'), int(ivs['金額（億円）'].sum()) if len(ivs) else (153993 if k=='I' else 0)])
evd=pd.DataFrame(rows,columns=['ID','期間','市場背景','分類','開始日','開始USDJPY','終了日','終了USDJPY','期間騰落率(%)','最大下落の起点日','起点値','最大下落の到達日','到達値','期間内最大円高率(%)','米10年債利回り変化(%pt)','円買い介入実施日（財務省）','介入額合計（億円）'])
# curves
uc=U.iloc[[-1]].T; uc.columns=['2026-11-04']
U25=pd.read_csv(R+'ust2025.csv'); U25['Date']=pd.to_datetime(U25['Date']); uc['2025-11-04']=U25.set_index('Date').loc['2025-11-04']
uc['2026-10-05（10年債の2008年以降最高日）']=U.loc['2026-10-05']
jc=J.iloc[[-1]].T; jc.columns=[f'{J.index[-1].date()}']; jc[str(J[:'2025-11-04'].index[-1].date())]=J[:'2025-11-04'].iloc[-1]
# ladder
u=U.iloc[-1]; ten=[1,2,3,5,7,10]; uy=[u['1 Yr'],u['2 Yr'],u['3 Yr'],u['5 Yr'],u['7 Yr'],u['10 Yr']]
lad=pd.DataFrame({'残存(年)':range(1,11)}); lad['米国債パー利回り(%)']=np.interp(lad['残存(年)'],ten,uy).round(3)
lad['補間有無']=['実測' if y in ten else '線形補間' for y in lad['残存(年)']]
lad['日本国債(%)']=[J.iloc[-1][f'{y}Y'] for y in range(1,11)]
hcv=(sofr-tona).dropna(); 
kpi=[
['米国10年国債利回り',u['10 Yr'],'%','2026/11/04','米国財務省 Daily Par Yield Curve'],
['米国2年国債利回り',u['2 Yr'],'%','2026/11/04','同上'],
['米国30年国債利回り',u['30 Yr'],'%','2026/11/04','同上'],
['米2年−10年スプレッド',round(u['10 Yr']-u['2 Yr'],2),'%pt','2026/11/04','同上より算出'],
['米10年債 2008年以降最高値',5.31,'%','2026/10/05','FRB H.15（FRED DGS10）'],
['米10年債 2008年以降のパーセンタイル（FRED 11/3値4.95%）',round((d10['2008':]<d10.iloc[-1]).mean()*100,1),'%','2026/11/03','FRED DGS10より算出'],
['米国債ラダー（1〜10年・均等）参考YTM',round(lad['米国債パー利回り(%)'].mean(),2),'%','2026/11/04','米国財務省より算出（構成は仮定・要確認）'],
['日本国債ラダー（1〜10年・均等）',round(lad['日本国債(%)'].mean(),2),'%','2026/11/04','財務省 国債金利情報より算出'],
['日本10年国債利回り',J.iloc[-1]['10Y'],'%','2026/11/04','財務省 国債金利情報'],
['日米10年金利差',round(u['10 Yr']-J.iloc[-1]['10Y'],2),'%pt','2026/11/04','算出'],
['FF金利誘導目標レンジ','3.75〜4.00','%','2026/10/28決定','FOMC声明（10対2、2名は0.25%利上げを主張）'],
['FOMC見通し中央値 2026/2027/2028/長期','4.1 / 4.1 / 3.9 / 3.2','%','2026/9/16','FRB SEP'],
['FOMC見通し PCEインフレ 2026','3.7','%','2026/9/16','FRB SEP'],
['日銀 政策金利（無担保コール翌日物誘導目標）','1.25程度','%','2026/9/18変更、10/30据え置き','日本銀行 公表文'],
['SOFR',sofr.iloc[-1],'%',str(sofr.index[-1].date()),'NY連銀（FRED SOFR）'],
['無担保コール翌日物（TONA）',tona.dropna().iloc[-1],'%',str(tona.dropna().index[-1].date()),'日本銀行 STRDCLUCON'],
['ヘッジコスト目安（SOFR−TONA）',round(hcv.iloc[-1],2),'%',str(hcv.index[-1].date()),'算出（ベーシス・フォワード実勢は未反映）'],
['フルヘッジ後 米国債ラダー参考利回り',round(lad['米国債パー利回り(%)'].mean()-hcv.iloc[-1],2),'%','2026/11/04・11/02','算出（参考値）'],
['USD/JPY（NY正午）',fx.iloc[-1],'円/ドル',str(fx.index[-1].date()),'FRB H.10（FRED DEXJPUS）'],
['USD/JPY（東京17時）',157.75,'円/ドル','2026/11/02','日本銀行 FXERD04'],
['USD/JPY 2026年高値（1986年以来）',163.86,'円/ドル','2026/07/29','FRB H.10'],
]
kp=pd.DataFrame(kpi,columns=['指標','値','単位','基準日','出所'])
sep=pd.DataFrame([['FF金利（中央値）',4.1,4.1,3.9,3.6,3.2],['FF金利（6月時点中央値）',3.8,3.6,3.4,3.1,None],['実質GDP',2.3,2.4,2.2,2.1,2.0],['失業率',4.1,4.1,4.1,4.1,4.2],['PCEインフレ',3.7,2.3,2.1,2.0,2.0],['コアPCE',3.4,2.5,2.2,2.0,None]],columns=['項目','2026','2027','2028','2029','長期'])
pol=pd.DataFrame({'FF上限':ub,'FF下限':lb}).loc['2020':]; pol=pol[pol['FF上限'].diff()!=0]
hcm=pd.DataFrame({'米実効FF金利(月平均)':dff.resample('ME').mean(),'無担保コール翌日物(月平均)':tona.resample('ME').mean()}).dropna()
hcm['短期金利差(%)']=hcm.iloc[:,0]-hcm.iloc[:,1]; hcm['米10年債(月平均)']=d10.resample('ME').mean(); hcm['米10年債フルヘッジ目安']=hcm['米10年債(月平均)']-hcm['短期金利差(%)']; hcm['日本10年国債(月平均)']=J['10Y'].resample('ME').mean(); hcm=hcm['2000':].round(3)
daily=pd.DataFrame({'USDJPY':fx,'US10Y':d10,'US2Y':d2,'SOFR':sofr,'TONA':tona,'EFFR':dff}).loc['2005':].dropna(how='all')
src=pd.DataFrame([
['米国債利回り（イールドカーブ）','米国財務省 Daily Treasury Par Yield Curve Rates','https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve','〜2026/11/04'],
['米国債利回り 長期時系列','FRB H.15（FRED: DGS1〜DGS30）','https://fred.stlouisfed.org/series/DGS10','〜2026/11/03'],
['FF金利誘導目標','FRB（FRED: DFEDTARU, DFEDTARL）／FOMC声明','https://www.federalreserve.gov/newsevents/pressreleases/monetary20261028a.htm','2026/10/28'],
['FOMC経済見通し','FRB Summary of Economic Projections','https://www.federalreserve.gov/monetarypolicy/fomcprojtabl20260916.htm','2026/9/16'],
['SOFR / 実効FF金利','NY連銀（FRED: SOFR, DFF）','https://fred.stlouisfed.org/series/SOFR','〜2026/11/03'],
['日本国債金利','財務省 国債金利情報','https://www.mof.go.jp/jgbs/reference/interest_rate/index.htm','〜2026/11/04'],
['無担保コール翌日物・USDJPY（東京）','日本銀行 時系列統計データ（STRDCLUCON, FXERD04）','https://www.stat-search.boj.or.jp/','〜2026/11/02'],
['日銀金融政策','日本銀行 当面の金融政策運営について（2026/10/30）他','https://www.boj.or.jp/mopo/mpmdeci/state_2026/index.htm','2026/10/30'],
['USD/JPY（NY）','FRB H.10（FRED: DEXJPUS）','https://fred.stlouisfed.org/series/DEXJPUS','〜2026/10/30'],
['為替介入（日次）','財務省 外国為替平衡操作の実施状況（過去の介入実績全データ）','https://www.mof.go.jp/policy/international_policy/reference/feio/data/index.html','〜2026年6月期'],
['為替介入（月次）','財務省 外国為替平衡操作の実施状況（月次ベース）','https://www.mof.go.jp/policy/international_policy/reference/feio/data/monthly/index.html','〜2026/10/28'],
],columns=['データ','出所','URL','データ基準日'])
src['取得日']='2026/11/05'
out='/home/user/Work/ust_ladder_fx_hedge/market_data_pack.xlsx'
with pd.ExcelWriter(out,engine='openpyxl') as w:
    src.to_excel(w,sheet_name='出所一覧',index=False); kp.to_excel(w,sheet_name='主要指標',index=False)
    uc.to_excel(w,sheet_name='米国債カーブ'); jc.to_excel(w,sheet_name='日本国債カーブ'); lad.to_excel(w,sheet_name='ラダー参考利回り',index=False)
    sep.to_excel(w,sheet_name='FOMC見通し(2026年9月)',index=False); pol.to_excel(w,sheet_name='FF金利変更履歴')
    hcm.to_excel(w,sheet_name='ヘッジコスト目安(月次)'); ivd.to_excel(w,sheet_name='為替介入(日次)',index=False); ivm.to_excel(w,sheet_name='為替介入(月次2026)',index=False)
    evd.to_excel(w,sheet_name='バックテスト候補局面',index=False); daily.to_excel(w,sheet_name='日次データ')
from openpyxl import load_workbook
wb=load_workbook(out)
for ws in wb:
    for col in ws.columns:
        L=max(len(str(c.value)) if c.value is not None else 0 for c in col[:200]); ws.column_dimensions[col[0].column_letter].width=min(max(10,L*1.3),60)
    ws.freeze_panes='B2'
wb.save(out)
pd.set_option('display.width',250); print(evd.iloc[:,[0,1,4,5,6,7,8,9,10,11,12,13,14,15]].to_string())
