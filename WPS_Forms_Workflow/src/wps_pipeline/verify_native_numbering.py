"""Independent readback of effective output numbering and managed fields.

Does not import the emitter, recognizer, authorizer or NumberingFacts. The
restart oracle uses occurrence clocks rather than the production counter reset.
"""
import re
from collections import defaultdict
from lxml import etree as E

N={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
Q=lambda name:'{'+N['w']+'}'+name


def val(node,path,default=None):
    child=node.find(path,N) if node is not None else None
    return child.get(Q('val'),default) if child is not None else default


def effective(parts):
    styles=E.fromstring(parts['word/styles.xml']);definitions={s.get(Q('styleId')):s for s in styles.findall('w:style',N)}
    default=next((k for k,v in definitions.items() if v.get(Q('default'))=='1' and v.get(Q('type'))=='paragraph'),'Normal')
    def prop(p,path):
        result=val(p,'w:pPr/'+path)
        if result is not None:return result
        key=val(p,'w:pPr/w:pStyle',default);visited=set()
        while key in definitions and key not in visited:
            visited.add(key);s=definitions[key];result=val(s,'w:pPr/'+path)
            if result is not None:return result
            key=val(s,'w:basedOn')
        return None
    return prop


def render(value,fmt):
    if fmt in ('decimal','decimalZero'):return str(value).zfill(2) if fmt=='decimalZero' else str(value)
    if fmt in ('upperLetter','lowerLetter') and value>0:
        v=value;s=''
        while v:v,x=divmod(v-1,26);s=chr(65+x)+s
        return s.lower() if fmt=='lowerLetter' else s
    if fmt=='chineseCounting' and 0<value<100:
        digits='零一二三四五六七八九'
        if value<10:return digits[value]
        return (digits[value//10] if value>=20 else '')+'十'+(digits[value%10] if value%10 else '')
    if fmt in ('lowerRoman','upperRoman') and value>0:
        text=''
        for n,s in ((1000,'M'),(900,'CM'),(500,'D'),(400,'CD'),(100,'C'),(90,'XC'),(50,'L'),(40,'XL'),(10,'X'),(9,'IX'),(5,'V'),(4,'IV'),(1,'I')):
            count,value=divmod(value,n);text+=s*count
        return text.lower() if fmt=='lowerRoman' else text
    return None


def labels(package):
    prop=effective(package.parts);root=E.fromstring(package.parts.get('word/numbering.xml',b'<w:numbering xmlns:w="'+N['w'].encode()+b'"/>'))
    nums={v.get(Q('numId')):v for v in root.findall('w:num',N)};abstracts={v.get(Q('abstractNumId')):v for v in root.findall('w:abstractNum',N)}
    counters=defaultdict(dict);used=defaultdict(dict);ticks=defaultdict(dict);answer={}
    for clock,p in enumerate(package.body.iter(Q('p')),1):
        nid=prop(p,'w:numPr/w:numId')
        try:level=int(prop(p,'w:numPr/w:ilvl') or 0)
        except (TypeError,ValueError):level=None
        if nid in (None,'0'):answer[p]=dict(num_id=nid,ilvl=level,status='none',label=None,values=None,definition=None);continue
        if level is None or not 0<=level<=8:answer[p]=dict(num_id=nid,ilvl=level,status='unresolved',label=None,values=None,definition=None);continue
        number=nums.get(nid);aid=val(number,'w:abstractNumId');abstract=abstracts.get(aid)
        if abstract is None:answer[p]=dict(num_id=nid,status='unresolved',label=None);continue
        try:
            levels={int(x.get(Q('ilvl'))):x for x in abstract.findall('w:lvl',N)};starts={}
            for override in number.findall('w:lvlOverride',N):
                k=int(override.get(Q('ilvl')));replacement=override.find('w:lvl',N)
                if replacement is not None:levels[k]=replacement
                if val(override,'w:startOverride') is not None:starts[k]=int(val(override,'w:startOverride'))
            for k,lvl in levels.items():starts.setdefault(k,int(val(lvl,'w:start','1')))
        except (TypeError,ValueError):answer[p]=dict(num_id=nid,ilvl=level,status='unresolved',label=None,values=None,definition=None);continue
        def reset_clock(k,seen=()):
            if k==0 or k in seen:return 0
            restart=val(levels.get(k),'w:lvlRestart');parent=k-1 if restart is None else int(restart)-1
            if restart=='0' or parent<0 or parent>=k:return 0
            return max(ticks[nid].get(parent,0),reset_clock(parent,seen+(k,)))
        if level not in levels:answer[p]=dict(num_id=nid,status='unresolved',label=None);continue
        if any(val(v,'w:lvlRestart') is not None and not val(v,'w:lvlRestart').isdigit() for v in levels.values()):answer[p]=dict(num_id=nid,status='unresolved',label=None);continue
        reset=reset_clock(level)
        if level not in used[nid] or reset>used[nid][level]:counters[nid][level]=starts[level]
        else:counters[nid][level]+=1
        used[nid][level]=clock;ticks[nid][level]=clock
        text=val(levels[level],'w:lvlText','');fmt=val(levels[level],'w:numFmt')
        def substitute(match):
            k=int(match[1])-1;f=val(levels.get(k),'w:numFmt');v=counters[nid].get(k,starts.get(k,1));result=render(v,f)
            return result if result is not None else '?'
        label=re.sub(r'%(\d)',substitute,text)
        if fmt=='bullet':label=text
        definition=dict(format=fmt,text=text,start=starts[level],suffix=val(levels[level],'w:suff','tab'),restart=val(levels[level],'w:lvlRestart'))
        valid_restarts=all(val(lvl,'w:lvlRestart') is None or (val(lvl,'w:lvlRestart').isdigit() and int(val(lvl,'w:lvlRestart'))<=k) for k,lvl in levels.items())
        supported=abstract.find('w:numStyleLink',N) is None and abstract.find('w:styleLink',N) is None and valid_restarts and '?' not in label
        values=[counters[nid].get(k,starts[k]) for k in range(level+1)] if all(k in starts for k in range(level+1)) else [counters[nid][level]]
        answer[p]=dict(num_id=nid,abstract_id=aid,ilvl=level,status='resolved' if supported else 'unresolved',label=label if supported else None,values=values if supported else None,definition=definition)
    return answer


def verify(report,out):
    from .verify_report_build import Package,checked_fields,ordinary,bookmark_texts
    errors=[];actual=labels(out);counts=defaultdict(lambda:[0]*5);chapter_ids=set();list_ids=set();chapter_abstracts=set();list_abstracts=set();groups={};seq_counts=defaultdict(int)
    def check(ok,detail):
        if not ok:errors.append('native:'+detail)
    native=bool(report.get('native_chapter_num_id'))
    for row in sorted(report['mapping'],key=lambda x:x['output_index']):
        p=out.body[row['output_index']];f=actual.get(p,{});role=row['role'];key=row['source_id']+':'+row['block_id'];changes=row['changes']
        if role=='heading':
            level=row['level'];counter=counts[row['source_id']];counter[level-1]+=1
            for k in range(level,5):counter[k]=0
            expected='.'.join(map(str,counter[:level]))
            source=next(s for s in report['sources'] if s['source_id']==row['source_id'])
            if report['master']:expected='.'.join(map(str,source.get('source_prefix') or [source['source_number']]))+'.'+expected
            if native or f.get('status')=='resolved':
                check(f.get('status')=='resolved' and f.get('label')==expected,key+':chapter_sequence')
                check(not any(c['operation']=='heading_number' and c['new'] for c in changes),key+':literal_and_native_duplicate')
                check(f.get('ilvl')==level-1+(len(source.get('source_prefix') or [source['source_number']]) if report['master'] else 0),key+':chapter_level')
                chapter_ids.add(f.get('num_id'));chapter_abstracts.add(f.get('abstract_id'))
        elif role=='document_title' and report['master']:
            # A source title is the approved H1 in a master. Verify it as such,
            # rather than treating subsequent source titles as body on outline.
            source=next(s for s in report['sources'] if s['source_id']==row['source_id'])
            if native or f.get('status')=='resolved':
                prefix=source.get('source_prefix') or [source['source_number']]
                check(f.get('status')=='resolved' and f.get('label')=='.'.join(map(str,prefix)),key+':source_group_sequence')
                check(f.get('ilvl')==len(prefix)-1,key+':source_group_level')
                check(not any(c['operation']=='source_group_number' and c['new'] for c in changes),key+':literal_and_native_group_duplicate')
                chapter_ids.add(f.get('num_id'));chapter_abstracts.add(f.get('abstract_id'))
        elif role=='list_item':
            manual=next((c for c in changes if c['operation']=='manual_list_prefix'),None)
            retained=any(c['operation']=='preserve_automatic_list_numbering' for c in changes)
            opaque=any(c['operation']=='preserve_automatic_list_numbering' and c.get('label_resolved') is False for c in changes)
            if manual or retained:
                check(f.get('status')=='resolved' or opaque,key+':list_number_resolves')
                list_ids.add(f.get('num_id'));list_abstracts.add(f.get('abstract_id'))
            else:check(any(c['operation']=='unresolved_numbering_notice' for c in changes) or any(x.get('block_id')==row['block_id'] and 'retained' in x.get('kind','') for x in report['exceptions']),key+':literal_list_not_declared')
            if manual:
                auth=manual['authorization'];check(auth['block_id']==row['block_id'] and auth['source_sha256']==next(s['sha256'] for s in report['sources'] if s['source_id']==row['source_id']),key+':list_permission_binding')
                marker=re.findall(r'\d+',f.get('label') or '')
                if marker:check(int(marker[-1])==auth['numbers'][-1],key+':list_ordinal_matches_source')
                else:check((f.get('label') or '').startswith(render(auth['numbers'][-1],'chineseCounting') or '?'),key+':chinese_list_ordinal')
                g=(row['source_id'],auth['group_anchor_id']);groups.setdefault(g,f.get('num_id'));check(groups[g]==f.get('num_id'),key+':list_group_instance')
            if retained and not opaque:
                source=next(s for s in report['sources'] if s['source_id']==row['source_id']);sp=Package(source['path']);before=labels(sp).get(sp.body[row['source_index']],{})
                check(f.get('label')==before.get('label'),key+':source_native_list_label')
        elif f.get('num_id') in chapter_ids and f.get('status')=='resolved':check(False,key+':nonheading_uses_outline')
        managed=any(c['operation'] in ('caption_number','materialize_caption_seq','native_caption_seq','promote_caption_number') for c in changes)
        if managed and report.get('object_numbering_contract')!='physical-objects-v1':
            label='图' if role=='figure_caption' else '表';seq_counts[label]+=1
            fields=[v for v in checked_fields(p) if re.match(r'\s*SEQ\b',v['instruction'],re.I)]
            check(len(fields)==1 and re.fullmatch(r'\s*SEQ\s+'+label+r'\s+\\\*\s+ARABIC\s*',fields[0]['instruction'],re.I) is not None,key+':caption_native_sequence')
            if fields:check(fields[0]['cache']==str(seq_counts[label]),key+':caption_cache_output_order')
    for group in report['generated_blocks']:
        if group.get('kind') in ('source_group_title','source_gap_heading','organization_group') and native:
            prefix=group.get('prefix') or group.get('source_prefix') or [group['source_number']]
            f=actual.get(out.body[group['output_index']],{});check(f.get('label')=='.'.join(map(str,prefix)),'generated_group:'+group.get('source_id',group.get('label','')));chapter_ids.add(f.get('num_id'));chapter_abstracts.add(f.get('abstract_id'))
    check(not chapter_ids&list_ids,'chapter_list_instance_isolation');check(not chapter_abstracts&list_abstracts,'chapter_list_definition_isolation')
    bookmarks=bookmark_texts(out.body)
    for row in report['mapping']:
        p=out.body[row['output_index']]
        for c in row['changes']:
            if c['operation'] in ('native_caption_ref','promote_caption_reference'):
                fields=[f for f in checked_fields(p) if f['instruction']==c.get('new_instruction',c.get('instruction'))]
                match=re.match(r'\s*REF\s+(\w+)',c.get('new_instruction',c.get('instruction','')),re.I)
                check(bool(match and fields and bookmarks.get(match[1])==c['new'] and any(f['cache']==c['new'] for f in fields)),'reference_bookmark_cache:'+row['block_id'])
    return errors
