"""Independent checks of selections against original frozen package content."""
from pathlib import Path
from zipfile import ZipFile
from collections import Counter
from copy import deepcopy
import hashlib,re
from lxml import etree as E
from .job_support import sha,read,contained
N={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main','m':'http://schemas.openxmlformats.org/officeDocument/2006/math'}
Q=lambda name:'{'+N['w']+'}'+name


def parts(path):
    with ZipFile(path) as z:return {n:z.read(n) for n in z.namelist()}


def text(n):return ''.join(n.xpath('.//w:t[not(ancestor::m:oMath)]/text()',namespaces=N))


def verify_preparation(run,inventory,plan):
    run=Path(run);errors=[];checks=0
    def ck(value,message):
        nonlocal checks
        checks+=1
        if not value:errors.append(message)
    sources={s['id']:s for s in inventory['sources']};entries=plan['entries'];excluded={x['id'] for x in plan['excluded']}
    ck(len({e['id'] for e in entries})==len(entries),'unique_selected_units')
    for s in inventory['sources']:
        selected=[e for e in entries if e['source_id']==s['id']]
        whole=[e for e in selected if e['kind']=='whole']
        ck(not whole or len(selected)==1,s['id']+':whole_and_split_mutually_exclusive')
        if not whole and s['id'] not in excluded:
            ids={u['id'] for u in s['units']}
            ck({e['id'] for e in selected}|(excluded&ids)==ids,s['id']+':all_suggested_units_accounted')
            intervals=[(u['start'],u['end']) for u in s['units']]
            sp=parts(run/s['frozen']);sb=E.fromstring(sp['word/document.xml']).find('w:body',N)
            cover=[i for a,b in intervals for i in range(a,b)]
            ck(cover==list(range(len(sb)-1)),s['id']+':partition_covers_all_body_blocks')
    for d in plan['derivations']:
        source=sources[d['source_id']];sp=parts(run/source['frozen']);up=parts(contained(run,d['file']))
        ck(sha(contained(run,d['file']))==d['sha256'],d['id']+':derived_hash')
        ck({n:v for n,v in sp.items() if n!='word/document.xml'}=={n:v for n,v in up.items() if n!='word/document.xml'},d['id']+':resources_numbering_styles_unchanged')
        sb=E.fromstring(sp['word/document.xml']).find('w:body',N);ub=E.fromstring(up['word/document.xml']).find('w:body',N)
        ck(E.tostring(sb[-1])==E.tostring(ub[-1]),d['id']+':section_preserved')
        unit=next((u for u in source['units'] if u['id']==d['id']),None)
        entry=next((e for e in entries if e['id']==d['id']),None)
        ck(unit is not None and d['title']==unit['title'] and text(ub[0])==unit['title'],d['id']+':intrinsic_source_unit_title')
        ck(entry is not None and d.get('display_title')==entry['title'],d['id']+':approved_display_title')
        maps=d['mapping'];ck([m['source_index'] for m in maps]==list(range(*d['range'])),d['id']+':exact_source_interval')
        used=[];candidate=read(run/'source_scan/sources'/source['id']/'candidate.json');byidx={r['source_locator']['body_child_index']:r for r in candidate['blocks']}
        for m in maps:
            a=sb[m['source_index']];b=ub[m['unit_index']];op=m['operation']
            if op in ('root_to_title','source_title_to_title'):
                ck(m['unit_index']==0,d['id']+':title_index')
                if op=='root_to_title':ck(m['authorization']==byidx[m['source_index']].get('number_edit') and m['authorization'].get('allowed'),d['id']+':root_number_permission')
                continue
            used.append(m['unit_index'])
            if op=='copy':ck(E.tostring(a)==E.tostring(b),d['id']+':exact_block:'+str(m['source_index']));continue
            ck(op=='rebase_heading',d['id']+':known_transform')
            edit=m['edit'];auth=edit['authorization'];lo,hi=auth['span']
            ck(auth==byidx[m['source_index']].get('number_edit') and auth.get('allowed'),d['id']+':original_authorization')
            from .verify_number_edits import heading_edit_error
            ck(heading_edit_error(text(a),dict(operation='heading_number',start=lo,end=hi,old=auth['old'],new=edit['new_prefix'],authorization=auth)) is None,d['id']+':independent_quantity_guard')
            ck(text(b)==text(a)[:lo]+edit['new_prefix']+text(a)[hi:],d['id']+':only_authorized_prefix_changed')
            # Compare all non-format/non-text structure, including drawings,
            # equations, fields, and their attributes; then compare math text.
            def skeleton(n):
                n=deepcopy(n)
                for x in list(n.findall('w:pPr',N)):n.remove(x)
                for x in n.xpath('.//w:t[not(ancestor::m:oMath)]',namespaces=N):x.text=''
                # A zero-length insertion may add an ordinary text run.
                for x in list(n.findall('w:r',N)):
                    if all(c.tag==Q('t') for c in x) and not any(c.text for c in x):n.remove(x)
                return E.tostring(n)
            ck(skeleton(a)==skeleton(b),d['id']+':objects_math_fields_preserved')
        ck(sorted(used)==list(range(1,len(ub)-1)),d['id']+':no_duplicate_or_unmapped_derived_blocks')
    return dict(status='PASS' if not errors else 'FAIL',checks=checks,errors=errors,
                scope='source-to-unit content, exact intervals, approved prefix edits and selection coverage; not semantic truth')
