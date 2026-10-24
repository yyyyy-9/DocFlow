"""Opaque legacy carriers and local missing-media representation.

Never starts an OLE server, parses equation binary data or downloads a link.
The independent verifier has its own inventory and repair preconditions.
"""
import hashlib,posixpath,struct,zlib,re
from .docx_formatting import NS,q

V='urn:schemas-microsoft-com:vml'
O='urn:schemas-microsoft-com:office:office'
def is_legacy_equation(node):
    return any(n.get('ProgID','').lower().startswith(('equation.','mathtype.')) or n.get('equationxml') is not None for n in node.iter())
def pictures(node):
    result=[]
    for n in node.iter():
        if n.tag==q('w:drawing') and n.xpath('.//a:blip[@r:embed or @r:link]',namespaces=NS):result.append(n)
        elif n.tag==q('w:pict') and not is_legacy_equation(n) and not any(a.tag==q('w:object') for a in n.iterancestors()) and n.findall('.//{'+V+'}imagedata'):result.append(n)
    return result
def image_attribute(node,key):
    return (node.tag==q('a:blip') and key in (q('r:embed'),q('r:link'))) or (node.tag=='{'+V+'}imagedata' and key==q('r:id'))
def part_name(target):
    return target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('word',target))
def resources(node,parts,rels):
    answer=[]
    for n in node.iter():
        for k,rid in n.attrib.items():
            if not k.startswith('{'+NS['r']+'}'):continue
            rel=rels.get(rid);rel=dict(rel.attrib) if hasattr(rel,'attrib') else rel
            row=dict(relationship_id=rid,attribute=k,kind='image' if image_attribute(n,k) else 'related')
            if rel is None:row.update(status='MISSING_RELATIONSHIP',target=None)
            elif rel.get('TargetMode')=='External':row.update(status='EXTERNAL',target=rel.get('Target'),relationship_type=rel.get('Type'))
            else:
                target=part_name(rel.get('Target',''));row.update(status='PRESENT' if target in parts else 'MISSING_PART',target=target,relationship_type=rel.get('Type'))
                if target in parts:row['sha256']=hashlib.sha256(parts[target]).hexdigest()
            answer.append(row)
    return answer
def missing_png():
    # A fixed red cross is a program marker, never a fabricated technical image.
    width,height=120,80;data=bytearray()
    for y in range(height):
        data.append(0)
        for x in range(width):
            red=abs(x*height-y*width)<420 or abs((width-x)*height-y*width)<420 or x<2 or y<2 or x>=width-2 or y>=height-2
            data.extend((210,25,25) if red else (255,245,245))
    def chunk(name,data):return struct.pack('>I',len(data))+name+data+struct.pack('>I',zlib.crc32(name+data)&0xffffffff)
    return b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',width,height,8,2,0,0,0))+chunk(b'IDAT',zlib.compress(bytes(data)))+chunk(b'IEND',b'')
def equation_references(text):
    """Explicit formula references, excluding scientific values and word suffixes."""
    found=[]
    patterns=[r'(?:公式|(?<![格方模样形制程公])式)\s*[（(]\s*(\d+)\s*[）)]',
              r'(?:参见|见|由|根据|代入|利用|按|如)\s*(?:公式|式)\s*(\d+)(?![\d.eE×xX*·])',
              r'(?:^|[\s，。；;：:])式\s*(\d+)(?![\d.eE×xX*·])']
    for pattern in patterns:
        for m in re.finditer(pattern,text):
            tail=text[m.end():]
            if re.match(r'\s*(?:[×xX*·]\s*10|[eE][+\-]?\d)',tail):continue
            found.append(m.group(1))
    return found
