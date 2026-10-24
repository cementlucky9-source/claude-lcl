import sys; sys.path.insert(0,'.')
from tiles import *
import numpy as np
from PIL import ImageEnhance, ImageFilter
lat0,lon0,lat1,lon1,z=23.15,61.2,26.95,68.4,9
sat,_=mosaic('esri',lat0,lon0,lat1,lon1,z)
car,_=mosaic('carto',lat0,lon0,lat1,lon1,z)
sat.convert('RGB').save('coast_sat_raw.jpg'); car.convert('RGB').save('coast_carto_raw.png')
a=np.asarray(car.convert('RGB')).astype(int)
print(sat.size, 'carto colors', np.unique(a.reshape(-1,3),axis=0,return_counts=True)[0][:5])
from collections import Counter
c=Counter(map(tuple,a.reshape(-1,3)[::50])); print(c.most_common(6))
