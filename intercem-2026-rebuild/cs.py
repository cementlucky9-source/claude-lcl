import json,sys,time,urllib.request,urllib.parse
UA={'User-Agent':'IntercemDeckBuilder/1.0 (presentation research; python urllib)'}
def api(params):
    url='https://commons.wikimedia.org/w/api.php?'+urllib.parse.urlencode({**params,'format':'json'})
    for i in range(6):
        try: return json.load(urllib.request.urlopen(urllib.request.Request(url,headers=UA),timeout=30))
        except Exception as e: time.sleep(3*(i+1))
    raise SystemExit('fail '+url)
for q in sys.argv[1:]:
    d=api({'action':'query','generator':'search','gsrsearch':q,'gsrnamespace':6,'gsrlimit':20,'prop':'imageinfo','iiprop':'url|size|extmetadata','iiurlwidth':2400})
    print('##',q)
    for p in sorted(d.get('query',{}).get('pages',{}).values(),key=lambda p:p.get('index',0)):
        ii=p['imageinfo'][0]; m=ii.get('extmetadata',{})
        if not p['title'].lower().endswith(('.jpg','.jpeg','.png')): continue
        print(f"{ii['width']}x{ii['height']} | {m.get('LicenseShortName',{}).get('value','?')} | {p['title']} | {ii.get('thumburl')}")
    time.sleep(2)
