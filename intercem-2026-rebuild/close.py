import sys; sys.path.insert(0,'.')
from tiles import *
for name,b,z in [('kpt',(24.80,66.94,24.86,67.02),15),('pq',(24.70,67.25,24.82,67.42),14),('gw',(25.09,62.30,25.14,62.37),15)]:
    im,_=mosaic('esri',b[0],b[1],b[2],b[3],z); im=im.convert('RGB'); im.save(f'{name}_raw.jpg',quality=92)
    t=im.copy(); t.thumbnail((1000,1000)); t.save(f'prev_{name}.jpg'); print(name,im.size)
