"""Independent lexical invariant for declared heading edits.

No imports from the recognizer or its authorization code. This catches a
bounded family of unsafe plans as well as bad execution; it is not a general
semantic oracle for every scientific notation or heading language.
"""
import re
import hashlib,json


def list_edit_error(text,change):
    a=change.get('authorization',{});body=dict(a);binding=body.pop('binding',None)
    expected=hashlib.sha256(json.dumps(body,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode()).hexdigest()
    if not a.get('allowed') or binding!=expected:return 'list_authorization_integrity'
    if a.get('text_sha256')!=hashlib.sha256(text.encode()).hexdigest() or a.get('span')!=[change['start'],change['end']] or a.get('old')!=change['old']:return 'list_authorization_source_binding'
    if change['new'] or change['start']!=0 or text[:change['end']]!=change['old']:return 'list_prefix_span'
    if not re.fullmatch(r'\s*(?:\d+(?:\.\d+)*[、．.)）]?|[（(][一二三四五六七八九十百\d]+[）)]|[一二三四五六七八九十百]+[、.．)）]?)\s*',change['old']):return 'list_prefix_not_a_marker'
    guard=dict(change,operation='heading_number')
    return heading_edit_error(text,guard)


def heading_edit_error(text,change):
    if change['operation'] not in ('heading_number','source_group_number'):
        return None
    if change['start']!=0:return 'heading_edit_not_at_paragraph_start'
    # Check raw authorization independently; nonempty strings and 1 are not
    # explicit JSON true, even if upstream accepted or rehashed the record.
    resolution=change.get('authorization',{}).get('ambiguity_resolution',{})
    if not isinstance(resolution,dict):return 'invalid_numeric_ambiguity_resolution'
    if 'resolve_numeric_ambiguity' in resolution and type(resolution['resolve_numeric_ambiguity']) is not bool:
        return 'numeric_ambiguity_flag_not_boolean'
    # Tokenize the complete scalar independently of the production prefix parser.
    value=re.match(r'\s*[+\-−]?(?:\d+\.\d*|\.\d+|\d+)(?:[eE][+\-−]?\d+)?',text)
    if not value:return None
    token=value.group().strip();rest=text[value.end():]
    if re.search(r'[eE][+\-−]?\d+$',token):return 'scientific_value_as_heading'
    if re.match(r'^\s*\d{4}[-./年]\d{1,2}[-./月]\d{1,2}',text):return 'date_as_heading'
    if re.match(r'\s*[×x·]\s*10\s*(?:\^\s*)?[+\-−]?\d',rest):return 'power_notation_as_heading'
    if re.match(r'\s*(?:±|\+/-|[~～–—-]|至|到)\s*[+\-−]?\d',rest):return 'range_or_uncertainty_as_heading'
    # Independent unit-token check; no recognizer or edit-authorizer imports.
    chinese_units=set('米 毫米 厘米 微米 纳米 千米 秒 毫秒 微秒 纳秒 分钟 小时 摄氏度 开尔文 伏特 毫伏 伏 安培 毫安 安 帕斯卡 千帕 兆帕 帕 赫兹 千赫兹 兆赫兹 吉赫兹 瓦特 千瓦 瓦 焦耳 牛顿 千克 毫克 克 吨 升 毫升 分贝 弧度 度 平方米 平方厘米 平方毫米 立方米 立方厘米 立方毫米 倍 个 次 组 片 套 台 路 阶 级 像素 字节 比特 周期'.split())
    suffix=rest.lstrip()
    unit=next((u for u in sorted(chinese_units,key=len,reverse=True) if suffix.startswith(u)),None)
    if unit:
        suffix=suffix[len(unit):].lstrip()
        if suffix.startswith(('/', '每')):
            divisor=suffix[1:].lstrip()
            den=next((u for u in sorted(chinese_units,key=len,reverse=True) if divisor.startswith(u)),None)
            if den:suffix=divisor[len(den):].strip()
        complete=not suffix or suffix[0] in '，。；,:：!?！？' or suffix.startswith(('为','是','约为','左右','以上','以下','以内','以外'))
        resolution=change.get('authorization',{}).get('ambiguity_resolution',{})
        if complete:return 'chinese_quantity_as_heading'
        if not (resolution.get('resolve_numeric_ambiguity') is True and resolution.get('origin') in ('model','user') and resolution.get('reason') and resolution.get('evidence_ids')):
            return 'ambiguous_chinese_unit_title_without_resolution'
    if re.match(r'\s*[A-Za-zμµΩ°℃%]',rest):
        resolution=change.get('authorization',{}).get('ambiguity_resolution',{})
        known=re.match(r'\s*(?:[munpμµkMGT]?[VAW]|mm|cm|nm|μm|m|s|ms|Hz|kHz|MHz|GHz|MPa|Pa|K|kg|g|dB|dBm|dBc|°C|℃|%)(?=$|\s|[，。；,;])',rest,re.I)
        if known or not (resolution.get('resolve_numeric_ambiguity') is True and resolution.get('origin') in ('model','user') and resolution.get('reason') and resolution.get('evidence_ids')):
            return 'numeric_latin_or_unit_prefix_as_heading'
    if not rest.strip():return 'scalar_without_title_as_heading'
    if 0<change['end']<value.end():return 'heading_prefix_splits_numeric_token'
    return None
