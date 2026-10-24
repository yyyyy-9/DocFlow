"""Create project-owned synthetic diagrams/source; no external document input."""
from pathlib import Path
import argparse,hashlib,json,os,shutil,sys
ROOT=Path(__file__).resolve().parents[1]

def emf(path,color):
    import ctypes
    from ctypes import wintypes as W
    g=ctypes.WinDLL('gdi32',use_last_error=True)
    g.CreateEnhMetaFileW.argtypes=[W.HDC,W.LPCWSTR,ctypes.c_void_p,W.LPCWSTR];g.CreateEnhMetaFileW.restype=W.HDC
    g.CreateSolidBrush.argtypes=[W.DWORD];g.CreateSolidBrush.restype=W.HANDLE
    g.SelectObject.argtypes=[W.HDC,W.HANDLE];g.SelectObject.restype=W.HANDLE
    g.Rectangle.argtypes=[W.HDC,ctypes.c_int,ctypes.c_int,ctypes.c_int,ctypes.c_int]
    g.CloseEnhMetaFile.argtypes=[W.HDC];g.CloseEnhMetaFile.restype=W.HANDLE
    g.DeleteEnhMetaFile.argtypes=[W.HANDLE];g.DeleteObject.argtypes=[W.HANDLE]
    dc=g.CreateEnhMetaFileW(None,str(path),None,'DocFlow\0Synthetic diagram\0\0')
    if not dc:raise ctypes.WinError(ctypes.get_last_error())
    red,green,blue=color;brush=g.CreateSolidBrush(red|(green<<8)|(blue<<16));old=g.SelectObject(dc,brush)
    g.Rectangle(dc,0,0,240,160);g.SelectObject(dc,old);g.DeleteObject(brush)
    handle=g.CloseEnhMetaFile(dc);assert handle;g.DeleteEnhMetaFile(handle)

def source_files():
    from PIL import Image,ImageDraw
    home=ROOT/'HRS_Bridge/fixtures/public_source';assets=home/'assets';assets.mkdir(parents=True,exist_ok=True)
    colors=[(36+i*17,62+i*11,100+i*8) for i in range(9)]
    for i,color in enumerate(colors,1):
        image=Image.new('RGB',(240,160),'white');draw=ImageDraw.Draw(image);draw.rectangle((2,2,237,157),fill=color,outline='black',width=2);draw.line((10,140,225,15),fill='white',width=3)
        if not (assets/(f'{i:016x}'+'.png')).exists():image.save(assets/(f'{i:016x}'+'.png'))
    for i in range(7,10):
        path=assets/(f'{i:016x}'+'.emf')
        if path.is_file():continue
        if os.name=='nt':emf(path,colors[i-1])
        elif not path.is_file():raise RuntimeError('Checked-in synthetic EMF assets are required outside Windows')
    for i in range(3):
        for suffix in ('.png','.emf'):shutil.copyfile(assets/(f'{7+i:016x}'+suffix),assets/(f'{100+i:016x}'+suffix))
    def block(kind,data):return '```'+kind+'\n'+json.dumps(data,ensure_ascii=False,indent=2)+'\n```\n\n'
    raw='# Public integration sample\n\n## Native objects\n\nSynthetic data only. Inline mathematics: $x+1$, $y^2$, $a/b$, $\\alpha+1$ and $\\boldsymbol{k}$.\n\n'
    for i in range(1,10):
        data=dict(schema='report-figure-v1',id='fig-c6394cd828f7' if i==1 else 'fig-public-'+str(i),caption='Synthetic diagram '+str(i),kind='image',src='assets/'+f'{i:016x}'+'.png',alt='Project-owned geometric illustration')
        if i>=7:data.update(sourceAsset='assets/'+f'{i:016x}'+'.emf',rendition=dict(format='png',dpi=144,page=1,width=240,height=160,engine='synthetic'))
        raw+=block('report-figure',data)
    raw+='| Parameter | Value |\n|---|---|\n| Length | 1.5 mm |\n\n'
    for ident in ('tbl-public-first','tbl-deebf0cf6ded','tbl-public-last'):
        raw+=block('report-table',dict(schema='report-table-v1',id=ident,caption='Synthetic parameters',rows=[[dict(text=x,rowspan=1,colspan=1) for x in row] for row in [('Parameter','Value'),('Length','1.5 mm')]]))
    raw+='Reference @fig:fig-c6394cd828f7 and @tbl:tbl-deebf0cf6ded.\n\n'
    for i in range(5):
        ident='eq-c7babca5e6cd' if i==4 else 'eq-public-'+str(i)
        raw+='<!-- hrs:eq id='+ident+' -->\n$$\nx_'+str(i)+'=\\frac{1}{2}\n$$\n\n'
    raw+='::: {#fig-4c59dfec6c89 layout-ncol=3 layout-nrow=2 subfig-labels=on}\n'
    for i in range(6):
        key=100+i if i<3 else i-2
        raw+='![Public child '+str(i+1)+'](assets/'+f'{key:016x}'+'.png)'
        if i<3:raw+=' {source=assets/'+f'{key:016x}'+'.emf format=png dpi=144 width=240 height=160 engine=synthetic}'
        raw+='\n'
    raw+='Synthetic grouped diagrams\n:::\n'
    (home/'report.md').write_bytes(raw.encode('utf8'))
    (home/'meta.json').write_text(json.dumps(dict(render_profile='engineering')),encoding='utf8')
    return home

def main():
    p=argparse.ArgumentParser();p.add_argument('--archive',action='store_true');p.add_argument('--browser');a=p.parse_args()
    home=source_files();digest=hashlib.sha256((home/'report.md').read_bytes()).hexdigest()
    if a.archive:
        sys.path.insert(0,str(ROOT/'HTML_Report_Studio/app'))
        import html_cli
        output=ROOT/'HRS_Bridge/fixtures/hrs_sample.html'
        if output.exists():raise ValueError('Use a new archive output; remove only the previous generated fixture explicitly')
        result=html_cli.export_html(home,output,browser=a.browser)
        print(json.dumps(result,ensure_ascii=False))
    print('Public source SHA256:',digest)

if __name__=='__main__':main()
