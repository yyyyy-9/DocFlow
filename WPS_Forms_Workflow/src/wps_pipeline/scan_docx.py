"""Experimental v2 fact scanner: ordered segments + physical containers.

Reads only the supplied DOCX. No truth, fixture roles, or mutation metadata.
Native payloads stay on disk; display_text is a reading aid, not transport data.
"""
import argparse
import hashlib
import json
import posixpath
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
from .numbering_facts import NumberingFacts
from .object_transport import is_legacy_equation,pictures,resources

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M='http://schemas.openxmlformats.org/officeDocument/2006/math'
WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
N={'w':W,'m':M,'wp':WP,'a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
VERSION='2.0.0-experimental.8'


def sha(data):return hashlib.sha256(data).hexdigest()


def canonical_node(node):
    children=[canonical_node(c) for c in node if isinstance(c.tag,str)]
    return [node.tag,sorted(node.attrib.items()),node.text if not children else None,children]


def structural_hash(node):
    data=json.dumps(canonical_node(node),ensure_ascii=False,sort_keys=True,separators=(',',':')).encode('utf8')
    return sha(data)


class Scanner:
    def __init__(self,path,output,page_facts=None):
        self.page_facts=page_facts
        self.path=Path(path);self.output=Path(output);self.output.mkdir(parents=True,exist_ok=False)
        (self.output/'payloads').mkdir()
        with ZipFile(path) as z:self.parts={n:z.read(n) for n in z.namelist()}
        self.root=E.fromstring(self.parts['word/document.xml']);self.body=self.root.find('w:body',N)
        styles=E.fromstring(self.parts['word/styles.xml'])
        self.styles={s.get('{'+W+'}styleId'):s for s in styles.findall('w:style',N)}
        (self.output/'payloads/styles.xml').write_bytes(self.parts['word/styles.xml'])
        self.default=next((k for k,s in self.styles.items() if s.get('{'+W+'}default')=='1' and s.get('{'+W+'}type')=='paragraph'),'Normal')
        self.rels={x.get('Id'):dict(x.attrib) for x in E.fromstring(self.parts['word/_rels/document.xml.rels'])}
        self.blocks=[];self.objects={};self.unsupported=[]
        self.numbering_facts=NumberingFacts(self.parts)

    def attr(self,element,path):
        if element is None:return None
        node=element.find(path,N)
        return node.get('{'+W+'}val') if node is not None else None

    def style_attr(self,sid,path,seen=None):
        seen=set() if seen is None else seen
        if sid in seen:return None
        seen.add(sid);s=self.styles.get(sid)
        direct=self.attr(s,path)
        if direct is not None:return direct
        based=self.attr(s,'w:basedOn')
        return self.style_attr(based,path,seen) if based else None

    def obj(self,node,kind):
        oid=f'o{len(self.objects)+1:06d}';raw=E.tostring(node,encoding='utf8',with_tail=False)
        rel='payloads/'+oid+'.xml';(self.output/rel).write_bytes(raw)
        data={'object_id':oid,'kind':kind,'payload_ref':rel,'xml_sha256':sha(raw),'structural_sha256':structural_hash(node),
              'source_path':self.root.getroottree().getpath(node)}
        self.objects[oid]=data
        return data

    def paragraph(self,node,container,body_index):
        block=self.new(node,'paragraph',container,body_index)
        sid=self.attr(node,'w:pPr/w:pStyle') or self.default
        outline=self.attr(node,'w:pPr/w:outlineLvl') or self.style_attr(sid,'w:pPr/w:outlineLvl')
        block.update({'style_id':sid,'style_name':self.attr(self.styles.get(sid),'w:name'),
                      'outline_level_raw':int(outline) if outline is not None else None,
                      'numbering':{'num_id':self.attr(node,'w:pPr/w:numPr/w:numId'),'ilvl':self.attr(node,'w:pPr/w:numPr/w:ilvl')},
                      'numbering_effective':self.numbering_facts.inspect(node),
                      'paragraph_properties_raw':canonical_node(node.find('w:pPr',N)) if node.find('w:pPr',N) is not None else None,
                      'segments':[],'display_text':'','plain_text':'','hyperlinks':[],
                      'paragraph_format_effective':self.paragraph_layout(node,sid)})
        segments=block['segments'];visible=[];ordinary=[]

        def visit(e,run=None):
            local=E.QName(e).localname;space=E.QName(e).namespace
            if space==W and local in ('pPr','rPr'):return
            if space==W and local in ('proofErr','lastRenderedPageBreak'):
                # Spelling flags and cached pagination do not carry content.
                # Preserve their original XML in transport, inventory separately.
                o=self.obj(e,'layout_marker');segments.append({'kind':'layout_marker','object_id':o['object_id']});return
            if space==W and local=='hyperlink':
                rid=e.get('{'+N['r']+'}id')
                span={'link_index':len(block['hyperlinks']),'start':len(segments),'end':None}
                block.setdefault('hyperlink_ranges',[]).append(span)
                block['hyperlinks'].append({'attributes':dict(e.attrib),'relationship':self.rels.get(rid) if rid else None})
                for child in e:visit(child,run)
                span['end']=len(segments)
                return
            if space==W and local=='pict':
                # Narrow, observed legacy VML horizontal rule. Other VML remains
                # unsupported until separately inventoried and transported.
                if len(e)==1 and e[0].tag=='{urn:schemas-microsoft-com:vml}rect' and e[0].get('{urn:schemas-microsoft-com:office:office}hr')=='t':
                    o=self.obj(e,'legacy_separator');segments.append({'kind':'legacy_separator','object_id':o['object_id']});return
                kind='legacy_equation' if is_legacy_equation(e) else 'legacy_picture'
                o=self.obj(e,kind);o['resources']=resources(e,self.parts,self.rels)
                if kind=='legacy_picture':
                    o.update(layout='vml',media=[],extent={});segments.append({'kind':'legacy_picture','object_id':o['object_id']});visible.append('[图片:'+o['object_id']+']')
                else:segments.append({'kind':kind,'object_id':o['object_id']});visible.append('[公式对象:'+o['object_id']+']')
                return
            if space==W and local=='object':
                kind='legacy_equation' if is_legacy_equation(e) else 'embedded_object'
                o=self.obj(e,kind);o['resources']=resources(e,self.parts,self.rels);o['prog_ids']=[n.get('ProgID') for n in e.iter() if n.get('ProgID')]
                segments.append({'kind':kind,'object_id':o['object_id']});visible.append('[公式对象:'+o['object_id']+']' if kind=='legacy_equation' else '[嵌入对象:'+o['object_id']+']');return
            if space==M and local in ('oMath','oMathPara'):
                kind='math_display' if local=='oMathPara' else 'math_inline';o=self.obj(e,kind)
                simple=kind=='math_inline' and all(E.QName(x).localname in ('oMath','oMathPr','r','t','rPr','sty','nor','scr','rFonts','sz','szCs') for x in e.iter())
                letters=''.join(e.xpath('.//m:t/text()',namespaces=N))
                preview=letters if simple else '[数学:'+o['object_id']+']'
                o['preview']=preview;o['preview_status']='exact_simple_symbol' if simple else 'opaque_structure_preserved'
                segments.append({'kind':kind,'object_id':o['object_id']});visible.append(preview);return
            if space==WP and local in ('inline','anchor'):
                o=self.obj(e,'drawing');o['layout']=local;o['media']=[]
                extent=e.find('wp:extent',N);o['extent']=dict(extent.attrib) if extent is not None else {}
                for rid in e.xpath('.//a:blip/@r:embed',namespaces=N):
                    if rid not in self.rels:o.setdefault('missing_media',[]).append(dict(relationship_id=rid,reason='MISSING_RELATIONSHIP'));continue
                    rel=self.rels[rid]
                    if rel.get('TargetMode')=='External':o.setdefault('missing_media',[]).append(dict(relationship_id=rid,reason='EXTERNAL_ONLY'));continue
                    name=rel['Target'];name=name.lstrip('/') if name.startswith('/') else posixpath.normpath(posixpath.join('word',name))
                    if name not in self.parts:o.setdefault('missing_media',[]).append(dict(relationship_id=rid,reason='MISSING_PART',part=name));continue
                    media=self.parts[name];digest=sha(media);dest='payloads/'+digest+Path(name).suffix
                    (self.output/dest).write_bytes(media)
                    o['media'].append({'relationship_id':rid,'part':name,'sha256':digest,'payload_ref':dest})
                o['resources']=resources(e,self.parts,self.rels)
                if not e.xpath('.//a:blip[@r:embed or @r:link]',namespaces=N):
                    segments.append({'kind':'raw_fragment','object_id':o['object_id']});o['kind']='empty_drawing';return
                segments.append({'kind':'drawing','object_id':o['object_id']});visible.append('[图片:'+o['object_id']+']');return
            if space==W and local=='t':
                text=e.text or '';size=self.attr(run,'w:rPr/w:sz') or self.style_attr(sid,'w:rPr/w:sz')
                bold=self.attr(run,'w:rPr/w:b')
                if run is not None and run.find('w:rPr/w:b',N) is not None and bold is None:bold='1'
                if bold is None:bold=self.style_attr(sid,'w:rPr/w:b')
                rp=run.find('w:rPr',N) if run is not None else None
                segments.append({'kind':'text','text':text,'font_size_pt':float(size)/2 if size else None,'bold':None if bold is None else bold not in ('0','false','off'),
                                 'run_properties_raw':canonical_node(rp) if rp is not None else None})
                visible.append(text);ordinary.append(text);return
            if space==W and local in ('bookmarkStart','bookmarkEnd'):
                s={'kind':local,'bookmark_id':e.get('{'+W+'}id')}
                if local=='bookmarkStart':s['name']=e.get('{'+W+'}name')
                segments.append(s);return
            if space==W and local in ('tab','br','cr','noBreakHyphen','softHyphen'):
                segments.append({'kind':local,'attributes':dict(e.attrib)});visible.append('\t' if local=='tab' else '\n' if local in ('br','cr') else '-');return
            if space==W and local in ('fldChar','instrText','fldSimple'):
                o=self.obj(e,'field_evidence');segments.append({'kind':'field_evidence','object_id':o['object_id']})
                if local=='fldSimple':visible.append(''.join(e.xpath('.//w:t/text()',namespaces=N)))
                return
            if space==W and local in ('footnoteReference','endnoteReference','object','pict','txbxContent','sdt','ins','del','sym','hyperlink'):
                o=self.obj(e,'unsupported');segments.append({'kind':'unsupported','object_id':o['object_id']});self.unsupported.append({'block_id':block['block_id'],'reason':local});return
            if space==W and local not in ('r','drawing','smartTag'):
                o=self.obj(e,'unsupported');segments.append({'kind':'unsupported','object_id':o['object_id']});self.unsupported.append({'block_id':block['block_id'],'reason':local});return
            if space==W and local=='r':run=e
            for child in e:visit(child,run)

        for c in node:visit(c)
        block['display_text']=''.join(visible);block['plain_text']=''.join(ordinary)
        first=next((s for s in segments if s['kind']=='text' and s['text'].strip()),{})
        block['format_group']=[sid,first.get('font_size_pt'),first.get('bold')]
        return block['block_id']

    def paragraph_layout(self,node,sid):
        def value(path,attribute):
            direct=node.find('w:pPr/'+path,N) if node is not None else None
            if direct is not None and direct.get('{'+W+'}'+attribute) is not None:return direct.get('{'+W+'}'+attribute)
            seen=set();current=sid
            while current in self.styles and current not in seen:
                seen.add(current);style=self.styles[current];prop=style.find('w:pPr/'+path,N)
                if prop is not None and prop.get('{'+W+'}'+attribute) is not None:return prop.get('{'+W+'}'+attribute)
                current=self.attr(style,'w:basedOn')
            return None
        result={'alignment':value('w:jc','val') or 'left'}
        result.update({key:value('w:ind',key) or '0' for key in ('firstLine','firstLineChars','hanging','hangingChars','left','leftChars','right','rightChars')})
        return result

    def new(self,node,kind,container,index):
        b={'block_id':f'b{len(self.blocks)+1:06d}','kind':kind,'container_id':container,
           'source_locator':{'part':'word/document.xml','path':self.root.getroottree().getpath(node),'body_child_index':index}}
        self.blocks.append(b);return b

    def table(self,node,container,index):
        b=self.new(node,'table',container,index);b.update({'rows':len(node.findall('w:tr',N)),'columns':len(node.findall('w:tblGrid/w:gridCol',N)),'cells':[]})
        for ri,tr in enumerate(node.findall('w:tr',N)):
            ci=0
            for tc in tr.findall('w:tc',N):
                cell=self.new(tc,'cell',b['block_id'],index);span=int(self.attr(tc,'w:tcPr/w:gridSpan') or 1)
                vm=tc.find('w:tcPr/w:vMerge',N)
                cell.update({'row':ri,'col':ci,'col_span':span,'v_merge_raw':None if vm is None else vm.get('{'+W+'}val','continue'),'children':[]})
                b['cells'].append(cell['block_id']);ci+=span
                for ch in tc:
                    if ch.tag=='{'+W+'}p':cell['children'].append(self.paragraph(ch,cell['block_id'],index))
                    elif ch.tag=='{'+W+'}tbl':cell['children'].append(self.table(ch,cell['block_id'],index))
        return b['block_id']

    def document_layout_facts(self,top):
        """Raw body layout/text facts, including opaque cover/TOC containers.

        No role or exclusion decision is made here. Header/footer parts are
        deliberately outside this title search.
        """
        facts=[];spans=[];active=[];bookmarks=[];byid={b['block_id']:b for b in self.blocks}
        for bid in top:
            block=byid[bid]
            index=block['source_locator']['body_child_index'];node=self.body[index];paragraphs=[]
            for p in node.iter('{'+W+'}p'):
                sid=self.attr(p,'w:pPr/w:pStyle') or self.default
                text=[]
                for event in p.iter():
                    if event.tag=='{'+W+'}t':text.append(event.text or '')
                    elif event.tag=='{'+W+'}tab':text.append('\t')
                    elif event.tag in ('{'+W+'}br','{'+W+'}cr'):text.append('\n')
                paragraphs.append(dict(text=''.join(text),plain_text=''.join(p.xpath('.//w:t/text()',namespaces=N)),path=self.root.getroottree().getpath(p),
                    style_name=self.attr(self.styles.get(sid),'w:name'),
                    outline=self.attr(p,'w:pPr/w:outlineLvl') or self.style_attr(sid,'w:pPr/w:outlineLvl'),
                    alignment=self.paragraph_layout(p,sid)['alignment'],
                    hyperlinks=p.xpath('.//w:hyperlink/@w:anchor',namespaces=N)))
            instructions=node.xpath('.//w:instrText/text()|.//w:fldSimple/@w:instr',namespaces=N)
            facts.append(dict(block_id=bid,body_child_index=index,tag=E.QName(node).localname,
                xml_sha256=sha(E.tostring(node,encoding='utf8',with_tail=False)),paragraphs=paragraphs,
                field_instructions=instructions,galleries=node.xpath('.//w:docPartGallery/@w:val',namespaces=N),
                drawing_count=len(pictures(node)),math_count=len(node.findall('.//m:oMath',N))+sum(is_legacy_equation(n) for n in node.iter() if n.tag in ('{'+W+'}object','{'+W+'}pict')),
                page_break=bool(node.xpath('.//w:br[@w:type="page"]|.//w:pageBreakBefore[not(@w:val="0")]|.//w:sectPr',namespaces=N))))
            for event in node.iter():
                if event.tag in ('{'+W+'}bookmarkStart','{'+W+'}bookmarkEnd'):
                    bookmarks.append(dict(kind=E.QName(event).localname,bookmark_id=event.get('{'+W+'}id'),
                        name=event.get('{'+W+'}name'),body_child_index=index,path=self.root.getroottree().getpath(event)))
                if event.tag=='{'+W+'}fldSimple':
                    spans.append(dict(instruction=event.get('{'+W+'}instr',''),start=index,end=index,complete=True))
                elif event.tag=='{'+W+'}fldChar':
                    kind=event.get('{'+W+'}fldCharType')
                    if kind=='begin':active.append(dict(instruction='',start=index,end=None,complete=False,separated=False))
                    elif kind=='separate' and active:active[-1]['separated']=True
                    elif kind=='end' and active:
                        span=active.pop();span.update(end=index,complete=True);spans.append(span)
                elif event.tag=='{'+W+'}instrText' and active and not active[-1]['separated']:
                    active[-1]['instruction']+=event.text or ''
        return facts,spans+active,bookmarks

    def run(self):
        top=[]
        for i,node in enumerate(self.body):
            if node.tag=='{'+W+'}p':top.append(self.paragraph(node,None,i))
            elif node.tag=='{'+W+'}tbl':top.append(self.table(node,None,i))
            elif node.tag!='{'+W+'}sectPr':
                b=self.new(node,'boundary',None,i);o=self.obj(node,'boundary');b['object_id']=o['object_id'];top.append(b['block_id'])
                b['opaque_preserved']=True
        section=self.body.find('.//w:sectPr',N)
        page=section.find('w:pgSz',N) if section is not None else None
        margins=section.find('w:pgMar',N) if section is not None else None
        width=(int(page.get('{'+W+'}w'))-int(margins.get('{'+W+'}left'))-int(margins.get('{'+W+'}right')))*635 if page is not None and margins is not None else None
        normal=self.paragraph_layout(None,self.default)
        layout,fields,bookmarks=self.document_layout_facts(top)
        result={'schema_version':'2.0-experimental','scan_version':VERSION,'source_sha256':sha(self.path.read_bytes()),
                'body_layout_facts':layout,'document_field_spans':fields,'document_bookmarks':bookmarks,
                'body_width_emu':width,'normal_paragraph_layout':[normal['alignment']]+[normal[k] for k in ('firstLine','firstLineChars','hanging','hangingChars','left','leftChars','right','rightChars')],
                'source_name':self.path.name,'styles_ref':'payloads/styles.xml','top_level_ids':top,'blocks':self.blocks,'objects':self.objects,'unsupported':self.unsupported,
                'scope':'Body and cell content; raw merged-cell facts; no role inference or page layout verification.'}
        if self.page_facts and self.page_facts.get('source_sha256')==result['source_sha256'] and self.page_facts.get('status')=='PASS' and self.page_facts.get('source_unchanged'):
            result['first_page_facts']=self.page_facts
        with (self.output/'scan.json').open('x',encoding='utf8') as f:json.dump(result,f,ensure_ascii=False,indent=2)
        return result
