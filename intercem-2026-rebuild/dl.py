import json,sys,time,urllib.request,urllib.parse,os
UA={'User-Agent':'IntercemDeckBuilder/1.0 (presentation research; python urllib)'}
titles=sys.argv[1:]
url='https://commons.wikimedia.org/w/api.php?'+urllib.parse.urlencode({'action':'query','titles':'|'.join(titles),'prop':'imageinfo','iiprop':'url|size|extmetadata','iiurlwidth':2600,'format':'json'})
for i in range(6):
    try: d=json.load(urllib.request.urlopen(urllib.request.Request(url,headers=UA),timeout=30)); break
    except Exception: time.sleep(4*(i+1))
meta=json.load(open('meta.json')) if os.path.exists('meta.json') else {}
for p in d['query']['pages'].values():
    ii=p['imageinfo'][0]; m=ii['extmetadata']
    fn='w_'+''.join(c if c.isalnum() else '_' for c in p['title'][5:])[:60]+'.jpg'
    u=ii.get('thumburl') or ii['url']
    for k in range(6):
        try:
            open(fn,'wb').write(urllib.request.urlopen(urllib.request.Request(u,headers=UA),timeout=60).read()); break
        except Exception as e: time.sleep(4*(k+1))
    import re
    meta[fn]={'title':p['title'],'license':m.get('LicenseShortName',{}).get('value'),'artist':re.sub('<[^>]+>','',m.get('Artist',{}).get('value','')).strip()}
    print(fn,meta[fn]); time.sleep(1)
json.dump(meta,open('meta.json','w'),indent=1)
