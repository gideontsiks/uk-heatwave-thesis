import json, csv, collections, os

# Output folder is created next to wherever this script is run from.
OUT = 'looker'
os.makedirs(OUT, exist_ok=True)

# Region centroids (approx lat/lon) for Looker bubble maps.
CENTROID = {
 'North East': (54.90,-1.80), 'North West': (54.00,-2.70),
 'Yorkshire and The Humber': (53.80,-1.30), 'East Midlands': (52.90,-0.80),
 'West Midlands': (52.50,-2.00), 'East of England': (52.20,0.50),
 'East': (52.20,0.50), 'London': (51.50,-0.12),
 'South East': (51.20,-0.50), 'South West': (50.80,-3.60), 'Wales': (52.30,-3.80),
}

g = json.load(open('gis_results.json'))
a = json.load(open('analysis_results.json'))

lisa = {d['region']: d for d in g['lisa']}
getis = {d['region']: d for d in g['getis_ord']}

# 1) Region summary (one row per region)
rows = list(csv.DictReader(open('spatial_analysis_vars.csv')))
with open(f'{OUT}/looker_region_summary.csv','w',newline='') as f:
    w = csv.writer(f)
    w.writerow(['Region','Latitude','Longitude','Population','ExcessDeaths',
                'ExcessPer100k','GreenspacePerCapita_m2','AvgDistParkM','PctBuiltUp',
                'DwellingsPerKm2','PopDensityPerKm2','LISA_Cluster','LISA_p',
                'LISA_Significant','GetisOrd_Z','GetisOrd_Level'])
    for r in rows:
        nm = r['Region.name']
        lat,lon = CENTROID.get(nm,('',''))
        l = lisa.get(nm,{}); ge = getis.get(nm,{})
        w.writerow([nm,lat,lon,round(float(r['pop'])),round(float(r['excess']),1),
                    round(float(r['excess_per_100k']),2),
                    round(float(r['greenspace_m2_per_capita_fit']),1),
                    round(float(r['avg_dist_park_m']),1),
                    round(float(r['pct_builtup']),2),
                    round(float(r['dwellings_per_km2']),1),
                    round(float(r['pop_density_per_km2']),1),
                    l.get('quadrant',''), l.get('p',''),
                    'Yes' if l.get('significant') else 'No',
                    round(ge.get('z_gstar',0),2), ge.get('level','')])

# 2) Monthly mortality (region x year x month, sexes summed)
agg = collections.defaultdict(lambda: {'deaths':0,'pop':0,'temp':[],'heat':0,'n':0})
for r in csv.DictReader(open('panel_harmonized.csv')):
    key = (r['Region.name'], r['year'], r['month'])
    d = agg[key]
    d['deaths'] += float(r['deaths'])
    d['pop'] = max(d['pop'], float(r['population']))  # per-sex pop; use region total below
    d['temp'].append(float(r['temp']))
    d['heat'] += 1 if r['is_heat_day']=='True' else 0
    d['n'] += 1
# region total population per year (sum of sexes) from panel: recompute properly
poptot = collections.defaultdict(float)
seen = set()
for r in csv.DictReader(open('panel_harmonized.csv')):
    k=(r['Region.name'],r['year'],r['Sex'])
    if k in seen: continue
    seen.add(k); poptot[(r['Region.name'],r['year'])] += float(r['population'])
with open(f'{OUT}/looker_monthly_mortality.csv','w',newline='') as f:
    w=csv.writer(f)
    w.writerow(['Region','Year','Month','MonthStart','Deaths','Population',
                'DeathsPer100k','MeanTempC','HeatDays'])
    for (nm,y,m),d in sorted(agg.items()):
        pop = poptot[(nm,y)]
        dpc = d['deaths']/pop*100000 if pop else ''
        mt = sum(d['temp'])/len(d['temp'])
        hd = d['heat']//2  # counted per sex (2), collapse to days
        w.writerow([nm,y,int(m),f'{y}-{int(m):02d}-01',int(d['deaths']),round(pop),
                    round(dpc,2) if dpc!='' else '', round(mt,2), hd])

# 3) Annual global Moran's I
with open(f'{OUT}/looker_annual_moran.csv','w',newline='') as f:
    w=csv.writer(f); w.writerow(['Year','MoranI','p_value','Significant'])
    for y,d in sorted(g['moran_annual'].items()):
        w.writerow([y, round(d['I'],3), round(d['p_sim'],3),
                    'Yes' if d['p_sim']<0.05 else 'No'])

# 4) DLNM exposure-response curve
with open(f'{OUT}/looker_dlnm_curve.csv','w',newline='') as f:
    w=csv.writer(f); w.writerow(['TempC','RelativeRisk'])
    for pt in a['model4_curve']:
        w.writerow([round(pt['temp'],2), round(pt['rr'],4)])

print('Looker CSVs written to', OUT)
for fn in sorted(os.listdir(OUT)):
    import os
    n=sum(1 for _ in open(f'{OUT}/{fn}'))-1
    print(f'  {fn}: {n} rows')
