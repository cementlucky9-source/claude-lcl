import math, os, sys, io, urllib.request, concurrent.futures as cf
from PIL import Image
CACHE='/tmp/claude-0/build/tilecache'; os.makedirs(CACHE,exist_ok=True)
SRC={'esri':'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
     'carto':'https://basemaps.cartocdn.com/light_nolabels/{z}/{x}/{y}.png',
     'cartolab':'https://basemaps.cartocdn.com/light_only_labels/{z}/{x}/{y}.png'}
def deg2px(lat,lon,z):
    n=256*2**z; x=(lon+180)/360*n
    y=(1-math.log(math.tan(math.radians(lat))+1/math.cos(math.radians(lat)))/math.pi)/2*n
    return x,y
def get(src,z,x,y):
    p=f'{CACHE}/{src}_{z}_{x}_{y}'
    if not os.path.exists(p):
        for i in range(4):
            try:
                r=urllib.request.urlopen(urllib.request.Request(SRC[src].format(z=z,x=x,y=y),headers={'User-Agent':'Mozilla/5.0'}),timeout=30).read()
                open(p,'wb').write(r); break
            except Exception as e:
                if i==3: raise
    return Image.open(p).convert('RGBA')
def mosaic(src,lat0,lon0,lat1,lon1,z):
    x0,y0=deg2px(lat1,lon0,z); x1,y1=deg2px(lat0,lon1,z)
    tx0,ty0,tx1,ty1=int(x0//256),int(y0//256),int(x1//256),int(y1//256)
    W=(tx1-tx0+1)*256;H=(ty1-ty0+1)*256
    img=Image.new('RGBA',(W,H))
    jobs={}
    with cf.ThreadPoolExecutor(16) as ex:
        for tx in range(tx0,tx1+1):
            for ty in range(ty0,ty1+1):
                jobs[(tx,ty)]=ex.submit(get,src,z,tx,ty)
    for (tx,ty),f in jobs.items(): img.paste(f.result(),((tx-tx0)*256,(ty-ty0)*256))
    off=(x0-tx0*256,y0-ty0*256)
    img=img.crop((int(off[0]),int(off[1]),int(off[0]+x1-x0),int(off[1]+y1-y0)))
    return img,(x0,y0)
