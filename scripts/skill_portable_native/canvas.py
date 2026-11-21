"""PowerPoint drawing primitives in reference-image coordinates, with native text and Office Math."""
from contextlib import contextmanager
from functools import lru_cache
from pathlib import Path
import json, math, re, warnings
from lxml import etree as ET
from PIL import ImageFont
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN, MSO_AUTO_SIZE
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn
from pptx.util import Pt

from .equations import math_run,expression
EMU_PER_MM=36000

TONES={'blue':('#DCE9F8','#5275A7'),'green':('#EFF6E6','#76A351'),
       'gray':('#F2F2F0','#92969A'),'purple':('#EAE2F6','#8E76AA')}
PAPER={'blue':('#EFF4FC','#A7BCDA'),'green':('#F1F6EB','#93AD78'),
       'gray':('#F6F6F4','#85898D'),'purple':('#F5F1FA','#B5A8CC')}
INK={'ink':'#17191D','arrow':'#5279AE','blue':'#5279AE','green':'#76A351','gray':'#898E95','purple':'#765B98',
     'black':'#101316','white':'#FFFFFF','none':None}
def em(v,scale):return round(v*scale)
def col(k):return INK.get(k,k)
def rgb(k):return RGBColor.from_string(col(k).lstrip('#'))
def elem(s,**attrs):
    n=OxmlElement(s)
    for k,v in attrs.items():n.set(k,str(v))
    return n
@lru_cache(maxsize=64)
def resolve_font(font,bold,italic):
    """Return the installed font file and its actual family name for measurement and PPTX."""
    from matplotlib.font_manager import FontProperties, findfont
    requested=FontProperties(family=font, weight='bold' if bold else 'normal',
                             style='italic' if italic else 'normal')
    path=Path(findfont(requested,fallback_to_default=True))
    actual=FontProperties(fname=str(path)).get_name()
    if actual.casefold()!=font.casefold():
        warnings.warn(f'Font {font!r} unavailable; using installed {actual!r}',RuntimeWarning,stacklevel=3)
    return path,actual

def metric_font(font,bold,italic):return resolve_font(font,bold,italic)[0]
@lru_cache(maxsize=32)
def metrics(path):return ImageFont.truetype(str(path),160)
def text_width(t,fs,font,bold,italic):return metrics(metric_font(font,bold,italic)).getlength(t)*fs/160

def load_formula_metrics(source):
    """Read an explicit JSON mapping: TeX -> [width_em, height_em]."""
    data=json.loads(Path(source).read_text()) if isinstance(source,(str,Path)) else dict(source)
    for tex, dims in data.items():
        if not isinstance(tex,str) or not isinstance(dims,(list,tuple)) or len(dims)!=2:
            raise ValueError(f'Invalid formula metric for {tex!r}; expected [positive width_em, positive height_em]')
        values=[float(v) for v in dims]
        if any(not math.isfinite(v) or v<=0 for v in values):
            raise ValueError(f'Invalid formula metric for {tex!r}; expected finite positive dimensions')
    return {tex:(float(dims[0]),float(dims[1])) for tex,dims in data.items()}

@lru_cache(maxsize=300)
def math_size(tex):
    # MathText only measures the mathematical ink box. The delivered object is OMML.
    from matplotlib import rc_context
    from matplotlib.font_manager import FontProperties
    from matplotlib.mathtext import MathTextParser
    expr=tex
    prop=FontProperties(size=100,math_fontfamily='stix')
    try:
        with rc_context({'mathtext.fontset':'stix'}):
            p=MathTextParser('path').parse('$'+expr+'$',dpi=72,prop=prop)
        return max(.28,p.width/100),max(.52,p.height/100)
    except ValueError:
        # Grouped and upright operators are still measured from their parsed math tree.
        base=''.join(expression(tex).itertext())
        return max(.3,metrics(metric_font('Times New Roman',False,True)).getlength(base)/160*.93),1.12

class Canvas:
    def __init__(self,slide,width_px=1300,height_px=743,width_mm=183,formula_metrics=None):
        self.slide=slide;self.shapes=slide.shapes
        if width_px<=0 or height_px<=0 or width_mm<=0:raise ValueError('Canvas dimensions must be positive')
        self.width_px=float(width_px);self.height_px=float(height_px);self.width_mm=float(width_mm)
        self.scale=self.width_mm*EMU_PER_MM/self.width_px
        self.pt_per_px=self.scale/12700
        if formula_metrics is not None:
            self.metrics=load_formula_metrics(formula_metrics)
        else:self.metrics={}
        self.shapes.turbo_add_enabled=True
        self.labels=[];self.formulas=[]
    @contextmanager
    def group(self,name):
        tree=self.shapes._spTree
        start=len(tree)
        yield
        children=list(tree)[start:]
        if not children:return
        wrappers=[self.shapes._shape_factory(child) for child in children]
        grp=self.shapes.add_group_shape(wrappers)
        # python-pptx allocates the group through spTree and bypasses the turbo counter.
        # Advance that counter exactly once so that the next object cannot reuse the group ID.
        grp._element.nvGrpSpPr.cNvPr.id=self.shapes._next_shape_id
        grp.name=name
    def _style(self,s,fill,stroke,width):
        style=s._element.find(qn('p:style'))
        if style is not None:s._element.remove(style)
        if fill and col(fill):
            s.fill.solid();s.fill.fore_color.rgb=rgb(fill)
        else:s.fill.background()
        if stroke and col(stroke) and width>0:
            s.line.color.rgb=rgb(stroke);s.line.width=Pt(width*self.pt_per_px)
        else:s.line.fill.background()
        return s
    def rect(self,x,y,w,h,fill,stroke=None,width=1.4,radius=0,dash=False,name='Rectangle'):
        pr=MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
        s=self.shapes.add_shape(pr,em(x,self.scale),em(y,self.scale),max(1,em(w,self.scale)),max(1,em(h,self.scale)));s.name=name or 'Rectangle'
        if radius:s.adjustments[0]=min(.49,float(radius)/max(1,min(w,h)))
        self._style(s,fill,stroke,width)
        if dash:s.line.dash_style=MSO_LINE_DASH_STYLE.DASH
        return s
    def ellipse(self,x,y,w,h,fill,stroke=None,width=1.4,name='Ellipse'):
        s=self.shapes.add_shape(MSO_SHAPE.OVAL,em(x,self.scale),em(y,self.scale),max(1,em(w,self.scale)),max(1,em(h,self.scale)));s.name=name or 'Ellipse'
        return self._style(s,fill,stroke,width)
    def panel(self,x,y,w,h,tone='blue',name='Panel',dash=True):
        fill,outline=PAPER[tone]
        return self.rect(x,y,w,h,fill,outline,2.1 if min(w,h)>140 else 1.35,min(27,min(w,h)*.10),dash,name)
    def box(self,x,y,w,h,tone='blue',name='',radius=True,dash=False):
        fill,outline=TONES[tone]
        return self.rect(x,y,w,h,fill,outline,1.65,min(11,h*.20,w*.20) if radius else 0,dash,name or 'Module')
    def _freeform(self,commands,fill,stroke,width,name,as_connector=False,head=False,dash=False):
        points=[p for cmd,pts in commands for p in pts]
        if not points:return
        xs,ys=zip(*points);l,t=min(xs),min(ys);w=max(.001,max(xs)-l);h=max(.001,max(ys)-t)
        if as_connector:
            s=self.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,em(l,self.scale),em(t,self.scale),em(l+w,self.scale),em(t+h,self.scale))
        else:s=self.shapes.add_shape(MSO_SHAPE.RECTANGLE,em(l,self.scale),em(t,self.scale),max(1,em(w,self.scale)),max(1,em(h,self.scale)))
        s.name=name or 'Editable curve';pr=s._element.spPr
        for child in list(pr):
            if ET.QName(child).localname in ('prstGeom','custGeom'):pr.remove(child)
        cg=elem('a:custGeom')
        for k in ('avLst','gdLst','ahLst','cxnLst'):cg.append(elem('a:'+k))
        cg.append(elem('a:rect',l='l',t='t',r='r',b='b'));lst=elem('a:pathLst')
        path=elem('a:path',w=max(1,round(w*1000)),h=max(1,round(h*1000)),fill='norm' if fill and col(fill) else 'none',stroke='1' if stroke and col(stroke) else '0')
        for cmd,pts in commands:
            entry=elem('a:'+cmd)
            for x,y in pts:entry.append(elem('a:pt',x=round((x-l)*1000),y=round((y-t)*1000)))
            path.append(entry)
        lst.append(path);cg.append(lst);pr.insert(1,cg)
        if as_connector:
            if s._element.find(qn('p:style')) is not None:s._element.remove(s._element.find(qn('p:style')))
            s.line.color.rgb=rgb(stroke);s.line.width=Pt(width*self.pt_per_px)
        else:self._style(s,fill,stroke,width)
        if stroke and col(stroke):
            ln=pr.get_or_add_ln();ln.set('cap','rnd')
            ln.append(elem('a:round'))
            if head:ln.append(elem('a:tailEnd',type='triangle',w='med',len='med'))
            if dash:s.line.dash_style=MSO_LINE_DASH_STYLE.DASH
        return s
    def poly(self,pts,fill,stroke=None,width=1.4,name='Polygon'):
        return self._freeform([('moveTo',[pts[0]])]+[('lnTo',[pt]) for pt in pts[1:]]+[('close',[])],fill,stroke,width,name)
    def arrow(self,points,tone='arrow',width=2.8,head=True,dash=False,name='Connection',radius=0):
        # Trim each adjacent straight and join their tangents with a native cubic quarter-circle.
        if radius and len(points)>2:
            pts=[points[0]]
            pts.extend(p for i,p in enumerate(points[1:]) if p!=points[i])
            commands=[('moveTo',[pts[0]])]
            for a,b,c in zip(pts,pts[1:],pts[2:]):
                ab,bc=math.dist(a,b),math.dist(b,c)
                dx,dy=(b[0]-a[0])/ab,(b[1]-a[1])/ab
                ex,ey=(c[0]-b[0])/bc,(c[1]-b[1])/bc
                if abs(dx*ey-dy*ex)<0.001:
                    commands.append(('lnTo',[b]))
                    continue
                r=min(radius,ab/2,bc/2)
                start=(b[0]-dx*r,b[1]-dy*r)
                end=(b[0]+ex*r,b[1]+ey*r)
                k=.55228475*r
                commands.extend([('lnTo',[start]),('cubicBezTo',[(start[0]+dx*k,start[1]+dy*k),
                                                               (end[0]-ex*k,end[1]-ey*k),end])])
            commands.append(('lnTo',[pts[-1]]))
            if not head:return self._freeform(commands,None,tone,width,name,dash=dash)
            # A short terminal stroke after a fillet can make an Office marker rotate
            # towards the preceding curve. Anchor the tip on the final segment tangent.
            a,tip=pts[-2:]
            length=math.dist(a,tip)
            ux,uy=(tip[0]-a[0])/length,(tip[1]-a[1])/length
            base=(tip[0]-ux*width*4,tip[1]-uy*width*4)
            half=width*2.05
            commands[-1]=('lnTo',[(tip[0]-ux*width,tip[1]-uy*width)])
            with self.group(name):
                self._freeform(commands,None,tone,width,name+' • rounded route',dash=dash)
                self.poly([tip,(base[0]+uy*half,base[1]-ux*half),(base[0]-uy*half,base[1]+ux*half)],tone,name=name+' • arrowhead')
            return
        if len(points)==2:
            (x,y),(xx,yy)=points
            s=self.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,em(x,self.scale),em(y,self.scale),em(xx,self.scale),em(yy,self.scale));s.name=name or 'Connection'
            style=s._element.find(qn('p:style'))
            if style is not None:s._element.remove(style)
            s.line.color.rgb=rgb(tone);s.line.width=Pt(width*self.pt_per_px)
            if dash:s.line.dash_style=MSO_LINE_DASH_STYLE.DASH
            ln=s._element.spPr.get_or_add_ln();ln.set('cap','rnd')
            if head:ln.append(elem('a:tailEnd',type='triangle',w='med',len='med'))
            return s
        with self.group(name or 'Routed connection'):
            segments=[(a,b) for a,b in zip(points[:-1],points[1:]) if a!=b]
            for i,(a,b) in enumerate(segments):
                self.arrow([a,b],tone,width,head and i==len(segments)-1,dash,f'{name} • segment {i+1}')
    def curve(self,points,tone='arrow',width=2.5,head=False,name='Curve'):
        assert (len(points)-1)%3==0,points
        return self._freeform([('moveTo',[points[0]])]+[('cubicBezTo',points[i:i+3]) for i in range(1,len(points),3)],None,tone,width,name,False,head)
    def brace(self,x,y,w,h,tone='ink',name='Brace'):
        y0=y;y1=y+h*.43;y2=y+h
        p=[(x,y0),(x+w*.006,y1),(x+w*.022,y1),(x+w*.055,y1),
           (x+w*.24,y1),(x+w*.40,y1),(x+w*.445,y1),
           (x+w*.48,y1),(x+w*.49,y1),(x+w*.5,y2),
           (x+w*.511,y1),(x+w*.52,y1),(x+w*.555,y1),
           (x+w*.65,y1),(x+w*.88,y1),(x+w*.945,y1),
           (x+w*.978,y1),(x+w*.992,y1),(x+w,y0)]
        return self.curve(p,tone,1.4,False,name)
    def bars(self,x,y,w,h,n=4,tone='green',name='Feature map',outer=True):
        with self.group(name):
            if outer:self.box(x,y,w,h,tone,'Module boundary')
            pad=min(w*.084,h*.1) if outer else 0
            gap=min(w*.07,h*.12);bw=(w-2*pad-(n-1)*gap)/n
            for j in range(n):
                fc={'green':'#E0EECC','blue':'#D3E4F7','purple':'#E7DCF5','gray':'#E8EBEB'}[tone]
                self.rect(x+pad+j*(bw+gap),y+pad,bw,h-2*pad,fc,TONES[tone][1],1.3,min(4,bw*.22),name=f'Feature channel {j+1}')
    def unet(self,x,y,w,h,tone='green',name='U-Net',snow=False,outer=True):
        with self.group(name):
            if outer:self.box(x,y,w,h,tone,'U-Net boundary')
            margin=w*.05;usable=w-2*margin;step=usable/7;bw=step*.75
            for i,height in enumerate([.80,.62,.41,.24,.41,.62,.80]):
                cx=x+margin+step*i
                a=h*height/2;b=a-h*.027 if i<3 else a+h*.027 if i>3 else a
                self.poly([(cx,y+h*.5-a),(cx+bw,y+h*.5-b),(cx+bw,y+h*.5+b),(cx,y+h*.5+a)],'#E4F0D1',TONES[tone][1],1.5,f'U-Net feature block {i+1}')
            if snow:self.snow(x+w*.1,y+h*.04,min(25,h*.2))
    def snow(self,x,y,size,name='Frozen network'):
        with self.group(name):
            for k in range(6):
                angle=k*math.pi/3;dx,dy=math.sin(angle),math.cos(angle);cx=x+size/2;cy=y+size/2;r=size*.47
                self.arrow([(cx,cy),(cx+dx*r,cy+dy*r)],'#278CB6',1.8,False,name='Ice crystal axis')
                for side in [-1,1]:
                    bx,by=cx+dx*r*.57,cy+dy*r*.57
                    self.arrow([(bx,by),(bx+dx*r*.17+dy*r*.27*side,by+dy*r*.17-dx*r*.27*side)],'#278CB6',1.45,False,name='Ice crystal branch')
    def _textbox(self,x,y,w,h,name):
        s=self.shapes.add_textbox(em(x,self.scale),em(y,self.scale),max(1,em(w,self.scale)),max(1,em(h,self.scale)));s.name=name[:230]
        tf=s.text_frame;tf.word_wrap=False;tf.auto_size=MSO_AUTO_SIZE.NONE;tf.vertical_anchor=MSO_ANCHOR.MIDDLE
        tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=0
        return s
    def _plain(self,x,y,w,h,lines,fs,bold,italic,align,color,font,name):
        s=self._textbox(x,y,w,h,name);tf=s.text_frame
        for i,line in enumerate(lines):
            p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
            p.space_before=p.space_after=Pt(0);p.line_spacing=Pt(fs*1.09*self.pt_per_px)
            p.alignment={'left':PP_ALIGN.LEFT,'center':PP_ALIGN.CENTER,'right':PP_ALIGN.RIGHT}[align]
            run=p.add_run();run.text=line;f=run.font;f.name=resolve_font(font,bold,italic)[1];f.size=Pt(fs*self.pt_per_px);f.bold=bold;f.italic=italic;f.color.rgb=rgb(color)
        self.labels.append((name,x,y,w,h,'\n'.join(lines)))
        return s
    def text(self,x,y,w,h,txt,size=20,bold=False,italic=False,align='center',color='ink',font='Times New Roman',name=None,math_color=None):
        lines=txt.split('\n');name=name or txt.replace('\n',' ')
        font=resolve_font(font,bold,italic)[1]
        if '$' not in txt:
            widest=max((text_width(l,1,font,bold,italic) for l in lines),default=1)
            fs=min(size,w*.993/max(widest,.02),h/(.91+1.09*(len(lines)-1)))
            return self._plain(x,y,w,h,lines,fs,bold,italic,align,color,font,name)
        chunks=[re.split(r'(\$[^$]+\$)',line) for line in lines]
        def width_at_one(part):
            if part.startswith('$'):
                dims=self.metrics[part[1:-1]] if part[1:-1] in self.metrics else math_size(part[1:-1])
                return dims[0]+.11
            return text_width(part,1,font,bold,italic)
        widest=max(sum(width_at_one(t) for t in row if t) for row in chunks)
        fs=min(size,w*.98/max(widest,.01),h/(1.02+1.09*(len(lines)-1)))
        with self.group(name):
            for j,row in enumerate(chunks):
                total=sum(width_at_one(t)*fs for t in row if t)
                xx=x if align=='left' else x+(w-total)/2 if align=='center' else x+w-total
                yc=y+h/2+(j-(len(lines)-1)/2)*1.09*fs
                for t in filter(None,row):
                    tw=width_at_one(t)*fs
                    if t.startswith('$'):
                        dims=self.metrics[t[1:-1]] if t[1:-1] in self.metrics else math_size(t[1:-1])
                        mh=max(1.20,dims[1])
                        self.eq(xx,yc-fs*mh/2,tw,fs*mh,t[1:-1],fs,name=name+' • equation',color=math_color or color)
                    else:self._plain(xx,yc-fs*.57,tw,fs*1.14,[t],fs,bold,italic,'left',color,font,t.strip() or 'Space')
                    xx+=tw
    def eq(self,x,y,w,h,tex,size=22,name=None,color='ink'):
        tex=tex.strip().strip('$')
        mw,mh=self.metrics[tex] if tex in self.metrics else math_size(tex)
        size=min(size,w/max(.02,mw),h/max(.5,mh))
        # Size the equation object to its natural typeset bounds. Impress fits native
        # math to the object extent; modern PowerPoint additionally reads run sizes.
        ew=mw*size;eh=mh*size
        left=x+(w-ew)/2;top=y+(h-eh)/2
        label=name or 'Equation • '+tex
        s=self._textbox(left,top,ew,eh,label)
        p=s.text_frame.paragraphs[0];p.alignment=PP_ALIGN.CENTER;p.font.name='Cambria Math';p.font.size=Pt(size*self.pt_per_px);p.font.color.rgb=rgb(color)
        p._p.append(math_run(tex,size*self.pt_per_px,col(color).lstrip('#')))
        self.formulas.append((label,left,top,ew,eh,tex))
        return s



def new_deck(width_px=1300,height_px=743,width_mm=183,formula_metrics=None):
    """Create a one-slide editable deck and its reference-coordinate canvas."""
    from pptx import Presentation
    from pptx.util import Mm
    if width_px<=0 or height_px<=0 or width_mm<=0:raise ValueError('Deck dimensions must be positive')
    presentation=Presentation()
    presentation.slide_width=Mm(width_mm)
    presentation.slide_height=Mm(width_mm*height_px/width_px)
    slide=presentation.slides.add_slide(presentation.slide_layouts[6])
    canvas=Canvas(slide,width_px,height_px,width_mm,formula_metrics)
    return presentation,canvas
