import sys; sys.path.insert(0,'.')
from tiles import *
from lightmap import lighten
import math
lat0,lon0,lat1,lon1,z=21.6,54.6,28.4,68.8,8
im,_=mosaic('esri',lat0,lon0,lat1,lon1,z); im.convert('RGB').save('region_raw.jpg',quality=92)
lighten('region_raw.jpg','region_light.jpg')
print(im.size)
def hav(a,b):
    R=6371;p1,p2=math.radians(a[0]),math.radians(b[0]);dp=p2-p1;dl=math.radians(b[1]-a[1])
    return 2*R*math.asin(math.sqrt(math.sin(dp/2)**2+math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2))
H=(26.57,56.25);G=(25.115,62.33);K=(24.835,66.98);Q=(24.775,67.33);D=(25.01,55.06);M=(23.6,58.55)
print('G-H',hav(G,H),'K-H',hav(K,H),'G-K',hav(G,K),'K-JebelAli',hav(K,D),'G-Muscat',hav(G,M))
