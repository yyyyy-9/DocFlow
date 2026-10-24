"""Shared DOCX formatting operations; no source-specific preview or CLI."""
from copy import deepcopy
import hashlib
import json
from pathlib import Path
import posixpath
import re
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree as E
from .resolve_styles import resolve_style_names

NS = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'm':'http://schemas.openxmlformats.org/officeDocument/2006/math',
      'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
      'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
REL='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
def q(name):
    prefix, local = name.split(':'); return '{'+NS[prefix]+'}'+local
def sha(path): return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def xml(data): return E.fromstring(data)
def pack(root): return E.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)
def elt(name, **attrs):
    node=E.Element(q(name))
    for key,value in attrs.items():node.set(q('w:'+key),str(value))
    return node
def ordinary_text(p):return ''.join(p.xpath('.//w:t[not(ancestor::m:oMath)]/text()',namespaces=NS))


def style_paragraph(p, style_id, preserve_emphasis=True, compact=False, source_styles=None):
    previous=p.find(q('w:pPr'))
    if previous is not None:p.remove(previous)
    props=elt('w:pPr');props.append(elt('w:pStyle',val=style_id));p.insert(0,props)
    # Preserve explicit content emphasis; remove source font, size and paragraph layout.
    allowed={'b','bCs','i','iCs','u','strike','dstrike','vertAlign'} if preserve_emphasis else {'vertAlign'}
    for run in p.xpath('.//w:r[not(ancestor::m:oMath)]',namespaces=NS):
        rp=run.find(q('w:rPr'))
        if rp is None:continue
        if preserve_emphasis and source_styles:
            ref=rp.find(q('w:rStyle'));sid=ref.get(q('w:val')) if ref is not None else None
            chain=[];seen=set()
            while sid in source_styles and sid not in seen:
                seen.add(sid);s=source_styles[sid];chain.append(s)
                base=s.find(q('w:basedOn'));sid=base.get(q('w:val')) if base is not None else None
            effective={}
            toggles={'b','bCs','i','iCs','strike','dstrike'}
            for s in reversed(chain):
                for prop in s.findall('w:rPr/*',NS):
                    key=E.QName(prop).localname
                    if key not in allowed:continue
                    val=prop.get(q('w:val'),'1')
                    if key in toggles:
                        if val not in ('0','false','off'):effective[key]=not effective.get(key,False)
                    else:effective[key]=deepcopy(prop)
            for key,val in effective.items():
                if rp.find(q('w:'+key)) is None:
                    rp.append(elt('w:'+key,val=int(val)) if key in toggles else val)
        kept=[deepcopy(c) for c in rp if E.QName(c).localname in allowed]
        run.remove(rp)
        if kept:
            new=elt('w:rPr');new.extend(kept);run.insert(0,new)
    if compact:
        props.append(elt('w:spacing',before=0,after=0,line=240,lineRule='auto'))
        props.append(elt('w:ind',left=0,right=0,firstLine=0,firstLineChars=0,leftChars=0,rightChars=0,hanging=0,hangingChars=0))
        props.append(elt('w:jc',val='center'))
    return props


def picture_row_table(paragraph, columns):
    """Partition picture-only paragraph events into cells, preserving their order."""
    table=elt('w:tbl');table.append(elt('w:tblPr'));grid=elt('w:tblGrid');table.append(grid)
    for _ in range(columns):grid.append(elt('w:gridCol',w=1))
    row=elt('w:tr');table.append(row);paragraphs=[]
    for _ in range(columns):
        cell=elt('w:tc');cell.append(elt('w:tcPr'));p=elt('w:p');cell.append(p);row.append(cell);paragraphs.append(p)
    index=0;seen=0
    for child in paragraph:
        if child.tag==q('w:pPr'):continue
        if child.tag==q('w:r'):
            rp=child.find(q('w:rPr'))
            for token in child:
                if token.tag==q('w:rPr'):continue
                if token.tag==q('w:drawing'):
                    if seen:index+=1
                    seen+=1
                if index>=columns:raise ValueError('Picture row has more drawings than declared columns')
                run=elt('w:r')
                if rp is not None:run.append(deepcopy(rp))
                run.append(deepcopy(token));paragraphs[index].append(run)
        else:paragraphs[index].append(deepcopy(child))
    if seen!=columns:raise ValueError('Picture row count changed before assembly')
    return table


def style_figure_table(table, styles, width, source_styles=None):
    """Image layout container, deliberately independent of the data table style."""
    old=table.find(q('w:tblPr'))
    if old is not None:table.remove(old)
    pr=elt('w:tblPr');table.insert(0,pr)
    pr.extend([elt('w:tblW',w=width,type='dxa'),elt('w:jc',val='center'),elt('w:tblInd',w=0,type='dxa'),elt('w:tblLayout',type='fixed')])
    borders=elt('w:tblBorders')
    for side in ('top','left','bottom','right','insideH','insideV'):borders.append(elt('w:'+side,val='nil'))
    pr.append(borders);margins=elt('w:tblCellMar')
    for side,value in (('top',65),('bottom',65),('left',85),('right',85)):margins.append(elt('w:'+side,w=value,type='dxa'))
    pr.append(margins)
    grid=table.find(q('w:tblGrid'));columns=list(grid) if grid is not None else []
    widths=[max(1,int(n.get(q('w:w'),'1'))) for n in columns];total=sum(widths) or 1
    widths=[round(width*v/total) for v in widths]
    if widths:widths[-1]+=width-sum(widths)
    for col,value in zip(columns,widths):col.set(q('w:w'),str(value))
    for row in table.findall(q('w:tr')):
        rp=row.find(q('w:trPr'))
        if rp is None:rp=elt('w:trPr');row.insert(0,rp)
        for n in list(rp):
            if n.tag in (q('w:trHeight'),q('w:tblHeader'),q('w:cantSplit')):rp.remove(n)
        rp.append(elt('w:cantSplit',val=0));cursor=0
        for cell in row.findall(q('w:tc')):
            cp=cell.find(q('w:tcPr'))
            if cp is None:cp=elt('w:tcPr');cell.insert(0,cp)
            span=cp.find(q('w:gridSpan'));count=int(span.get(q('w:val'))) if span is not None else 1
            cell_width=sum(widths[cursor:cursor+count]) if widths else width;cursor+=count
            for n in list(cp):
                if E.QName(n).localname in ('tcW','vAlign','tcMar','tcBorders','shd','noWrap','tcFitText'):cp.remove(n)
            cp.insert(0,elt('w:tcW',w=cell_width,type='dxa'));cp.append(elt('w:vAlign',val='center'))
            paragraphs=cell.findall(q('w:p'))
            for i,p in enumerate(paragraphs):
                image=bool(p.findall('.//w:drawing',NS))
                pp=style_paragraph(p,styles['图形段' if image else '图题'],source_styles=source_styles)
                pp.append(elt('w:jc',val='center'))
                pp.append(elt('w:ind',left=0,right=0,firstLine=0,hanging=0,leftChars=0,rightChars=0,firstLineChars=0,hangingChars=0))
                pp.append(elt('w:spacing',before=0,after=0,line=240,lineRule='auto'))
                for name,value in (('pageBreakBefore',0),('keepNext',int(image and i+1<len(paragraphs))),('keepLines',1)):
                    pp.append(elt('w:'+name,val=value))
            for nested in cell.findall(q('w:tbl')):style_figure_table(nested,styles,max(200,cell_width-170),source_styles)


def style_table(table, styles, width, source_styles=None):
    old=table.find(q('w:tblPr'))
    if old is not None:table.remove(old)
    pr=elt('w:tblPr');table.insert(0,pr)
    pr.extend([elt('w:tblStyle',val=styles['项目表格']),elt('w:tblW',w=width,type='dxa'),elt('w:jc',val='center'),elt('w:tblLayout',type='fixed')])
    borders=elt('w:tblBorders')
    for edge in ('top','left','bottom','right','insideH','insideV'):
        borders.append(elt('w:'+edge,val='single',sz=4,color='B7C1CC'))
    pr.append(borders)
    margins=elt('w:tblCellMar')
    for side,val in [('top',65),('bottom',65),('left',85),('right',85)]:margins.append(elt('w:'+side,w=val,type='dxa'))
    pr.append(margins)
    grid=table.find(q('w:tblGrid')); columns=list(grid) if grid is not None else []
    before=[int(c.get(q('w:w'),'1')) for c in columns]; total=sum(before) or 1
    widths=[round(width*x/total) for x in before]
    if widths:widths[-1]+=width-sum(widths)
    for node,val in zip(columns,widths):node.set(q('w:w'),str(val))
    rows=table.findall(q('w:tr'))
    for row_index,row in enumerate(rows):
        rpr=row.find(q('w:trPr'))
        if rpr is None:rpr=elt('w:trPr');row.insert(0,rpr)
        for node in list(rpr):
            if node.tag in (q('w:trHeight'),q('w:tblHeader'),q('w:cantSplit')):rpr.remove(node)
        # Long text rows may split; row count does not measure table height.
        rpr.append(elt('w:cantSplit',val=0))
        if row_index==0:rpr.append(elt('w:tblHeader'))
        cursor=0
        for cell in row.findall(q('w:tc')):
            cp=cell.find(q('w:tcPr'))
            if cp is None:cp=elt('w:tcPr');cell.insert(0,cp)
            span=cp.find(q('w:gridSpan'));n=int(span.get(q('w:val'))) if span is not None else 1
            cw=sum(widths[cursor:cursor+n]) if widths else width;cursor+=n
            for node in list(cp):
                if E.QName(node).localname in ('tcW','vAlign','tcMar','tcBorders','shd','noWrap','tcFitText'):cp.remove(node)
            cp.insert(0,elt('w:tcW',w=cw,type='dxa'));cp.append(elt('w:vAlign',val='center'))
            cp.append(elt('w:shd',val='clear',fill='EEF1F5' if row_index==0 else 'FFFFFF'))
            for p in cell.findall(q('w:p')):
                pp=style_paragraph(p,styles['表内正文'],compact=True,source_styles=source_styles)
                for name in ('keepNext','keepLines','pageBreakBefore'):
                    pp.append(elt('w:'+name,val=0))
            for nested in cell.findall(q('w:tbl')):style_table(nested,styles,max(200,cw-170),source_styles)
