"""Build genuine presentation Office Math from LaTeX without a picture fallback."""
from functools import lru_cache
from copy import deepcopy
import re
from lxml import etree as ET
from latex2mathml.converter import convert as latex_to_mathml
from mathml2omml import convert as mathml_to_omml

M='http://schemas.openxmlformats.org/officeDocument/2006/math'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
A14='http://schemas.microsoft.com/office/drawing/2010/main'
def node(ns,tag,**attrs):return ET.Element('{'+ns+'}'+tag,attrs)
def tag(ns,t):return '{'+ns+'}'+t

@lru_cache(maxsize=256)
def expression(tex):
    try:
        xml=mathml_to_omml(latex_to_mathml(tex))
    except Exception as exc:
        raise ValueError(f'Cannot convert formula {tex!r} to editable Office Math: {exc}') from exc
    math=ET.fromstring(xml.replace('<m:oMath>',f'<m:oMath xmlns:m="{M}">',1))
    # mathml2omml emits boxes around every mrow. Unwrap plain grouping and keep the mathematical substructure intact.
    for box in reversed(list(math.iter(tag(M,'box')))):
        e=box.find(tag(M,'e'))
        if e is None or len(box)!=1:continue
        parent=box.getparent();index=parent.index(box)
        for child in list(e):
            parent.insert(index,child);index+=1
        parent.remove(box)
    # Accent hats have a dedicated native math element.
    for lim in list(math.iter(tag(M,'limUpp'))):
        sign=lim.find(tag(M,'lim')); base=lim.find(tag(M,'e'))
        if base is not None and sign is not None and ''.join(sign.itertext()) in ('^','ˆ','̂'):
            acc=node(M,'acc');pr=node(M,'accPr');ch=node(M,'chr');ch.set(tag(M,'val'),'̂');pr.append(ch);acc.extend([pr,base]);lim.getparent().replace(lim,acc)
    # Keep words together. Unmerged math letters introduce artificial inter-letter spacing in Impress.
    for parent in list(math.iter()):
        previous=None
        for child in list(parent):
            if child.tag!=tag(M,'r'):
                previous=None
                continue
            t=child.find(tag(M,'t'))
            if t is None or not t.text:
                parent.remove(child)
                continue
            if t.text.isascii() and t.text.isalpha():
                props=child.find(tag(M,'rPr'))
                style=ET.tostring(props) if props is not None else None
                if previous and previous[0]==style:
                    previous[1].text+=t.text
                    parent.remove(child)
                else:previous=(style,t)
            else:previous=None
    for r in math.iter(tag(M,'r')):
        t=r.find(tag(M,'t'))
        pr=r.find(tag(M,'rPr'))
        if t is None or t.text is None:continue
        style=pr.find(tag(M,'sty')) if pr is not None else None
        if t.text in ('*','>','<') or (t.text.isalpha() and len(t.text)>1 and
                style is not None and style.get(tag(M,'val'))=='p'):
            if pr is None:pr=node(M,'rPr');r.insert(0,pr)
            pr.insert(0,node(M,'nor' if t.text.isalpha() else 'lit'))
    if any('\\' in (q.text or '') for q in math.iter(tag(M,'t'))):
        raise ValueError(f'Unconverted command in equation {tex!r}: '+str(''.join(math.itertext())))
    return math

def math_run(tex,size_pt,color='151A22'):
    wrapper=ET.Element(tag(A14,'m'),nsmap={'a14':A14,'m':M})
    para=node(M,'oMathPara')
    pr=node(M,'oMathParaPr');jc=node(M,'jc');jc.set(tag(M,'val'),'centerGroup');pr.append(jc);para.append(pr)
    math=deepcopy(expression(tex))
    for r in math.iter(tag(M,'r')):
        rp=node(A,'rPr',sz=str(round(size_pt*100)),lang='en-US')
        fill=node(A,'solidFill');fill.append(node(A,'srgbClr',val=color));rp.append(fill)
        rp.append(node(A,'latin',typeface='Cambria Math'))
        r.insert(1 if r.find(tag(M,'rPr')) is not None else 0,rp)
    para.append(math);wrapper.append(para)
    return wrapper
