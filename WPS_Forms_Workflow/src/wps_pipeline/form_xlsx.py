"""Fixed, inert XLSX forms. No spreadsheet engine or model service required.

Accepts inline/shared strings, including WPS/Excel resaves. Binding is checked
against the on-disk manifest by the caller, never trusted from hidden cells.
"""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import json, re, posixpath
from lxml import etree as E

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
N = {'s': S, 'r': R}


def col(n):
    out = ''
    while n:
        n, k = divmod(n-1, 26); out = chr(65+k)+out
    return out


def xml(node):
    return E.tostring(node, xml_declaration=True, encoding='UTF-8', standalone=True)


def write_form(path, kind, binding, rows, headers, editable, title, instructions,
               choices=None, widths=None, global_label='确认', global_value='未确认'):
    """Rows are explicit cell values; row 8 is the immutable header."""
    path = Path(path)
    if path.exists(): raise FileExistsError('表单不得覆盖: '+str(path))
    sheet = E.Element('{'+S+'}worksheet', nsmap={None:S,'r':R})
    views=E.SubElement(sheet,'{'+S+'}sheetViews'); view=E.SubElement(views,'{'+S+'}sheetView',workbookViewId='0')
    E.SubElement(view,'{'+S+'}pane',ySplit='8',topLeftCell='A9',activePane='bottomLeft',state='frozen')
    cols=E.SubElement(sheet,'{'+S+'}cols')
    for i,w in enumerate(widths or [20]*len(headers),1):
        E.SubElement(cols,'{'+S+'}col',min=str(i),max=str(i),width=str(w),customWidth='1')
    data=E.SubElement(sheet,'{'+S+'}sheetData')
    def row(number, values, styles=None):
        r=E.SubElement(data,'{'+S+'}row',r=str(number),ht='32' if number>=9 else '25',customHeight='1')
        for i,v in enumerate(values,1):
            if v is None: v=''
            c=E.SubElement(r,'{'+S+'}c',r=col(i)+str(number),s=str((styles or {}).get(i,3 if i in editable and number>=9 else 2)))
            if type(v) in (int,float): E.SubElement(c,'{'+S+'}v').text=str(v)
            else:
                c.set('t','inlineStr'); t=E.SubElement(E.SubElement(c,'{'+S+'}is'),'{'+S+'}t')
                t.set('{http://www.w3.org/XML/1998/namespace}space','preserve'); t.text=str(v)
    row(1,[title],{1:1}); row(3,[instructions],{1:2}); row(5,['表单版本',binding['schema']],{1:2,2:2})
    row(6,[global_label,global_value],{1:2,2:3}); row(8,headers,{i:1 for i in range(1,len(headers)+1)})
    for n,values in enumerate(rows,9): row(n,values)
    merges=E.SubElement(sheet,'{'+S+'}mergeCells',count='2')
    for ref in ('A1:'+col(len(headers))+'1','A3:'+col(len(headers))+'3'): E.SubElement(merges,'{'+S+'}mergeCell',ref=ref)
    if rows:
        auto=E.Element('{'+S+'}autoFilter',ref='A8:'+col(len(headers))+str(8+len(rows)))
        sheet.insert(list(sheet).index(merges),auto)
    if choices:
        dvs=E.SubElement(sheet,'{'+S+'}dataValidations',count=str(len(choices)))
        for area,values in choices.items():
            dv=E.SubElement(dvs,'{'+S+'}dataValidation',type='list',allowBlank='1',showErrorMessage='1',errorTitle='请选择约定值',error='请使用下拉选项',sqref=area)
            E.SubElement(dv,'{'+S+'}formula1').text='"'+','.join(values)+'"'
    # Locked cells are a UI hint; import validation is the actual boundary.
    protection=E.Element('{'+S+'}sheetProtection',sheet='1',objects='1',scenarios='1',autoFilter='0',selectLockedCells='0',selectUnlockedCells='0')
    sheet.insert(list(sheet).index(data)+1,protection)
    hidden=E.Element('{'+S+'}worksheet',nsmap={None:S}); hd=E.SubElement(hidden,'{'+S+'}sheetData')
    payload=json.dumps(dict(binding,kind=kind),ensure_ascii=False,sort_keys=True)
    for i,offset in enumerate(range(0,len(payload),24000),1):
        hr=E.SubElement(hd,'{'+S+'}row',r=str(i)); hc=E.SubElement(hr,'{'+S+'}c',r='A'+str(i),t='inlineStr')
        E.SubElement(E.SubElement(hc,'{'+S+'}is'),'{'+S+'}t').text=payload[offset:offset+24000]
    wb=E.Element('{'+S+'}workbook',nsmap={None:S,'r':R}); sheets=E.SubElement(wb,'{'+S+'}sheets')
    for i,name in enumerate((kind,'_绑定'),1):
        x=E.SubElement(sheets,'{'+S+'}sheet',name=name,sheetId=str(i));x.set('{'+R+'}id','rId'+str(i))
        if i==2:x.set('state','veryHidden')
    rels=E.Element('{'+P+'}Relationships',nsmap={None:P})
    for i,target in ((1,'worksheets/sheet1.xml'),(2,'worksheets/sheet2.xml'),(3,'styles.xml')):
        E.SubElement(rels,'{'+P+'}Relationship',Id='rId'+str(i),Type=R+('/styles' if i==3 else '/worksheet'),Target=target)
    styles=b'''<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Microsoft YaHei"/></font><font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Microsoft YaHei"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF173D59"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="3" borderId="0" applyAlignment="1" applyProtection="1"><alignment vertical="center" wrapText="1"/><protection locked="0"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'''
    ct=b'''<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'''
    rootrels=E.Element('{'+P+'}Relationships',nsmap={None:P}); E.SubElement(rootrels,'{'+P+'}Relationship',Id='rId1',Type=R+'/officeDocument',Target='xl/workbook.xml')
    path.parent.mkdir(parents=True,exist_ok=True)
    with ZipFile(path,'x',ZIP_DEFLATED) as z:
        for name,value in {'[Content_Types].xml':ct,'_rels/.rels':xml(rootrels),'xl/workbook.xml':xml(wb),'xl/_rels/workbook.xml.rels':xml(rels),'xl/styles.xml':styles,'xl/worksheets/sheet1.xml':xml(sheet),'xl/worksheets/sheet2.xml':xml(hidden)}.items():z.writestr(name,value)
    return str(path)


def read_form(path,kind):
    path=Path(path)
    if path.suffix.lower()!='.xlsx': raise ValueError('只接受无宏 .xlsx 表单')
    with ZipFile(path) as z:
        if sum(i.file_size for i in z.infolist())>30_000_000:raise ValueError('表单过大')
        if any('vbaProject' in n or n.startswith('xl/externalLinks/') for n in z.namelist()):raise ValueError('表单包含宏或外部工作簿关系')
        parse=lambda p:E.fromstring(z.read(p),E.XMLParser(resolve_entities=False,no_network=True))
        strings=[]
        if 'xl/sharedStrings.xml' in z.namelist():strings=[''.join(x.itertext()) for x in parse('xl/sharedStrings.xml')]
        rel={x.get('Id'):x.get('Target') for x in parse('xl/_rels/workbook.xml.rels')}
        sheets={s.get('name'):s for s in parse('xl/workbook.xml').findall('s:sheets/s:sheet',N)}
        def cells(name):
            if name not in sheets:raise ValueError('缺少固定工作表: '+name)
            target=rel[sheets[name].get('{'+R+'}id')]
            part=target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/'+target)
            if not part.startswith('xl/'):raise ValueError('工作表关系越界')
            result={}
            for c in parse(part).findall('.//s:sheetData/s:row/s:c',N):
                if c.find('s:f',N) is not None:raise ValueError('输入表单不接受公式: '+c.get('r',''))
                if c.get('t')=='inlineStr':value=''.join(c.xpath('s:is//s:t/text()',namespaces=N))
                else:
                    value=c.findtext('s:v','',N)
                    if c.get('t')=='s':value=strings[int(value)]
                result[c.get('r')]=value
            return result
        data=cells(kind); hidden=cells('_绑定')
        binding=json.loads(''.join(hidden[k] for k in sorted(hidden,key=lambda x:int(re.search(r'\d+',x)[0])) if k.startswith('A')) or '{}')
        if binding.get('kind')!=kind:raise ValueError('表单种类绑定不符')
    return data,binding
