# Гипотеза Арсагеры: выкупают акции компании с низким P/BV (ниже рынка, ниже 1).
import pandas as pd, numpy as np, json
from scipy import stats
from pathlib import Path
B=str(Path(__file__).absolute().parent.parent)+'/'
p=pd.read_csv(B+'pre_cancellation/panel.csv')
d=pd.read_csv(B+'buyback_precursors_1pct/dataset.csv')
# 1. P/BV рынка по годам (компании с положительным капиталом)
q=p[(p.equity>0)&p.cap.notna()&(p.year>=2016)]
mkt=q.groupby('year').apply(lambda g: pd.Series({'n':len(g),'pbv_aggr':g.cap.sum()/g.equity.sum(),
   'pbv_median':g.pbv.median(),'share_below1':(g.pbv<1).mean()})).reset_index()
# 2. Выкупавшие vs невыкупавшие, P/BV на конец t-1
s=d.dropna(subset=['f_pbv','bb']).copy(); s=s[s.f_pbv>0]
s['mkt_med']=s.year.map(dict(zip(mkt.year+1,mkt.pbv_median)))
s['mkt_aggr']=s.year.map(dict(zip(mkt.year+1,mkt.pbv_aggr)))
s['rel']=s.f_pbv/s.mkt_med
s=s.dropna(subset=['rel'])  # годы, для которых известен P/BV рынка на t-1
def cmp(a,b,name):
    r={'group':name,'n_a':len(a),'n_b':len(b)}
    for col in ['f_pbv','rel']:
        r[col+'_med_a']=a[col].median(); r[col+'_med_b']=b[col].median()
        r[col+'_p_mw']=stats.mannwhitneyu(a[col],b[col]).pvalue
    for lab,f in [('below1',lambda x:x.f_pbv<1),('above_mkt_med',lambda x:x.rel>1),('above_mkt_aggr',lambda x:x.f_pbv>x.mkt_aggr)]:
        ka,kb=f(a).sum(),f(b).sum()
        r[lab+'_a']=ka/len(a); r[lab+'_b']=kb/len(b)
        r[lab+'_p_fisher']=stats.fisher_exact([[ka,len(a)-ka],[kb,len(b)-kb]])[1]
    return r
res=[cmp(s[s.bb==1],s[s.bb==0],'выкуп vs нет выкупа'),
     cmp(s[s.cancel==1],s[s.bb==0],'погашение vs нет выкупа')]
s['ret_q']=s.f_roe
# 3. Сочетание: высокий ROE и P/BV — критерий E/P=ROE/(P/BV)
s2=s.dropna(subset=['f_roe']).copy(); s2['ep_ars']=s2.f_roe/s2.f_pbv
r3={'ep_ars_med_buy':s2[s2.bb==1].ep_ars.median(),'ep_ars_med_no':s2[s2.bb==0].ep_ars.median(),
    'p':stats.mannwhitneyu(s2[s2.bb==1].ep_ars,s2[s2.bb==0].ep_ars).pvalue,
    'roe_med_buy':s2[s2.bb==1].f_roe.median(),'roe_med_no':s2[s2.bb==0].f_roe.median()}
# 4. Логит по годам? простой: bb ~ log(rel) с кластерами
import statsmodels.formula.api as smf
s['lrel']=np.log(s.rel); s=s.dropna(subset=['lrel'])
m=smf.logit('bb ~ lrel',data=s).fit(disp=0,cov_type='cluster',cov_kwds={'groups':pd.factorize(s.company)[0]})
r4={'coef_lrel':m.params['lrel'],'p':m.pvalues['lrel'],'or_per_doubling':float(np.exp(m.params['lrel']*np.log(2)))}
by_year=s.groupby(['year','bb']).f_pbv.median().unstack()
by_year.columns=['P/BV невыкупавших','P/BV выкупавших']
by_year=by_year.join(mkt.assign(year=mkt.year+1).set_index('year')[['pbv_aggr','pbv_median']].rename(columns={'pbv_aggr':'P/BV рынка (агрегат, t-1)','pbv_median':'P/BV рынка (медиана, t-1)'}))
with pd.ExcelWriter(B+'pbv_hypothesis_1pct/pbv_hypothesis_results.xlsx') as w:
    mkt.to_excel(w,sheet_name='P_BV рынка',index=False); pd.DataFrame(res).to_excel(w,sheet_name='Сравнение',index=False)
    by_year.to_excel(w,sheet_name='По годам'); pd.DataFrame([r3]).to_excel(w,sheet_name='Критерий Арсагеры',index=False)
    pd.DataFrame([r4]).to_excel(w,sheet_name='Логит',index=False)
print(mkt.round(2).to_string()); print(pd.DataFrame(res).T.to_string()); print(r3); print(r4); print(by_year.round(2).to_string())
json.dump({'market':mkt.to_dict('records'),'cmp':res,'arsagera':r3,'logit':r4},open(B+'pbv_hypothesis_1pct/summary.json','w'),ensure_ascii=False,indent=1,default=float)
# 5. «Идеальный» по Арсагере кандидат: P/BV<1 и ROE не ниже медианы года (t-1), годы 2018+
z=d.dropna(subset=['f_pbv','f_roe','bb']); z=z[z.year>=2018].copy()
z['ideal']=(z.f_pbv<1)&(z.f_roe>=z.groupby('year').f_roe.transform('median'))
ct=pd.crosstab(z.bb,z.ideal)
r5={'buy_rate_ideal':z[z.ideal].bb.mean(),'n_ideal':int(z.ideal.sum()),'buy_rate_other':z[~z.ideal].bb.mean(),
    'n_other':int((~z.ideal).sum()),'p_fisher':stats.fisher_exact(ct.values)[1]}
print(r5)
with pd.ExcelWriter(B+'pbv_hypothesis_1pct/pbv_hypothesis_results.xlsx',mode='a',engine='openpyxl') as w:
    pd.DataFrame([r5]).to_excel(w,sheet_name='Идеал Арсагеры',index=False)
