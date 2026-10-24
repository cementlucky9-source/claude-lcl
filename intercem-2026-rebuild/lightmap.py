import numpy as np
from PIL import Image, ImageFilter, ImageEnhance
def lighten(src, out, sea=(222,231,236), keep=0.10, land_bright=1.08, land_sat=0.92):
    im=Image.open(src).convert('RGB')
    a=np.asarray(im).astype(float)
    r,g,b=a[...,0],a[...,1],a[...,2]
    lum=0.3*r+0.59*g+0.11*b
    water=((b>r+8)&(lum<150))|((b>r+25)&(g>r+10))
    m=Image.fromarray((water*255).astype('uint8')).filter(ImageFilter.MedianFilter(7)).filter(ImageFilter.GaussianBlur(2))
    m=np.asarray(m).astype(float)[...,None]/255
    land=Image.fromarray(a.astype('uint8'))
    land=ImageEnhance.Color(land).enhance(land_sat); land=ImageEnhance.Brightness(land).enhance(land_bright)
    L=np.asarray(land).astype(float)
    # sea: pale tone with a hint of the original texture (bathymetry)
    tex=(lum-lum.mean())[...,None]*keep
    S=np.array(sea,float)[None,None,:]+tex*1.5
    o=L*(1-m)+S*m
    Image.fromarray(np.clip(o,0,255).astype('uint8')).save(out,quality=92)
if __name__=='__main__':
    import sys; lighten(sys.argv[1],sys.argv[2])
