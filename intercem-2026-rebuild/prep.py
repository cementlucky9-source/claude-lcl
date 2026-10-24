from PIL import Image, ImageEnhance
import os
SP='/tmp/claude-0/-home-user-claude-lcl/0df4295e-eccc-5ee1-9fd9-e8dd1ffb1afe/scratchpad'
O=SP+'/orig/ppt/media/'; V=SP+'/v3x/ppt/media/'; P=SP+'/pdfimg/'; B='/tmp/claude-0/build/'
def crop(src,out,w,h,fx=.5,fy=.5,box=None,maxw=2400,enh=None):
    im=Image.open(src).convert('RGB')
    if box: im=im.crop(box)
    W,H=im.size; r=w/h
    if W/H>r: nw=int(H*r); x=int((W-nw)*fx); im=im.crop((x,0,x+nw,H))
    else: nh=int(W/r); y=int((H-nh)*fy); im=im.crop((0,y,W,y+nh))
    if im.size[0]>maxw: im=im.resize((maxw,int(maxw/r)),Image.LANCZOS)
    if enh: im=ImageEnhance.Brightness(im).enhance(enh)
    im.save(B+'assets/'+out,quality=90); print(out,im.size, f'{im.size[0]/w:.0f}dpi')
crop(V+'image14.jpg','title.jpg',7.13,7.5,fx=.62)
crop(B+'kpt_raw.jpg','index.jpg',5.0,7.5,fx=.55,box=(560,0,1864,1541))
crop(B+'coast_light.jpg','coast.jpg',13.333,7.5,fy=.25)
crop(V+'image-11-2.jpeg','kpt_sat.jpg',12.33,4.3,fy=.5)
crop(V+'image-12-1.jpeg','dredge.jpg',5.6,5.0)
crop(P+'kg-198.jpg','kgtml_plan.jpg',6.6,3.15)
crop(P+'pibt-004.jpg','pibt_aerial.jpg',7.4,4.6)
crop(P+'pibt-006.jpg','pibt_dusk.jpg',6.0,7.5,fx=.55)
crop(B+'gw_raw.jpg','gwadar.jpg',6.4,7.5,box=(600,200,1631,1286),fx=.35)
crop(V+'image13.jpg','plant.jpg',4.3,2.4,fx=.4)
crop(V+'image15.jpg','renew.jpg',6.6,7.5,fx=.55)
crop(O+'image81.jpeg','lineh.jpg',3.85,2.3)
crop(V+'image-17-4.jpeg','ifrs.jpg',3.85,2.3,enh=1.05)
crop(V+'image12.jpg','clinker.jpg',3.0,2.0)
crop(V+'image10.jpg','silos.jpg',3.0,2.0)
crop(V+'image-1-1.jpg','closing.jpg',6.67,7.5,fx=.5)
crop(O+'image43.jpg','bulkship.jpg',3.0,2.0)
import shutil
shutil.copy(O+'image2.png',B+'assets/logo.png')
shutil.copy(V+'image-12-2.png',B+'assets/kgtl_logo.png')
shutil.copy(V+'image9.png',B+'assets/pibt_logo.png')
shutil.copy(O+'image91.png',B+'assets/utis.png')
