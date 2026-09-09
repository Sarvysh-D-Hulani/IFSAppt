from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn

OUT = '/vercel/share/v0-project/Minda_Corporation_Investment_Case.pptx'
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

BG = RGBColor(247, 245, 239)
INK = RGBColor(37, 43, 42)
MUTED = RGBColor(100, 105, 99)
GREEN = RGBColor(45, 91, 72)
GREEN_LIGHT = RGBColor(213, 227, 216)
RUST = RGBColor(166, 84, 56)
RUST_LIGHT = RGBColor(239, 218, 207)
GOLD = RGBColor(190, 145, 62)
LINE = RGBColor(210, 211, 200)
WHITE = RGBColor(255, 255, 255)

FONT = 'Aptos'
FONT_HEAD = 'Aptos Display'

def set_bg(slide, color=BG):
    fill = slide.background.fill
    fill.solid(); fill.fore_color.rgb = color

def no_line(shape):
    shape.line.fill.background()

def box(slide, x, y, w, h, fill=WHITE, line=LINE, radius=False):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid(); shp.fill.fore_color.rgb = fill
    shp.line.color.rgb = line; shp.line.width = Pt(0.8)
    return shp

def text(slide, x, y, w, h, s, size=16, color=INK, bold=False, font=FONT, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP, italic=False):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.clear(); tf.word_wrap = True; tf.margin_left = Inches(0.04); tf.margin_right = Inches(0.04); tf.margin_top = Inches(0.02); tf.margin_bottom = Inches(0.02); tf.vertical_anchor = valign
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = s; r.font.name = font; r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic; r.font.color.rgb = color
    return tb

def rich_text(slide, x, y, w, h, runs, size=16, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tf=tb.text_frame; tf.clear(); tf.word_wrap=True
    p=tf.paragraphs[0]; p.alignment=align
    for s,c,b in runs:
        r=p.add_run(); r.text=s; r.font.name=FONT; r.font.size=Pt(size); r.font.color.rgb=c; r.font.bold=b
    return tb

def title(slide, kicker, heading, sub=''):
    text(slide, .65, .38, 12, .25, kicker.upper(), 9, GREEN, True, FONT, PP_ALIGN.LEFT)
    text(slide, .65, .68, 12, .6, heading, 27, INK, True, FONT_HEAD)
    if sub: text(slide, .67, 1.35, 11.8, .35, sub, 11, MUTED)
    line=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(.65), Inches(1.82), Inches(12.03), Inches(.015)); line.fill.solid(); line.fill.fore_color.rgb=LINE; no_line(line)

def footer(slide, n, note='Data as of FY25 / latest company disclosures; estimates are analyst calculations.'):
    text(slide, .67, 7.16, 10.8, .18, note, 7.5, MUTED, False, FONT, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE, True)
    text(slide, 12.1, 7.12, .6, .25, f'{n:02d}', 9, GREEN, True, FONT, PP_ALIGN.RIGHT)

def bullet_list(slide, x, y, w, items, size=13, color=INK, gap=.43, bullet=GREEN):
    for i, item in enumerate(items):
        yy=y+i*gap
        dot=slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(yy+.08), Inches(.09), Inches(.09)); dot.fill.solid(); dot.fill.fore_color.rgb=bullet; no_line(dot)
        text(slide, x+.2, yy, w-.2, .32, item, size, color)

def table(slide, x, y, widths, rows, row_h=.38, header=True, font_size=10):
    xx=x
    for ri,row in enumerate(rows):
        xx=x
        for ci,cell in enumerate(row):
            fill = GREEN if ri==0 and header else (WHITE if ri%2 else RGBColor(241,240,233))
            color = WHITE if ri==0 and header else INK
            box(slide, xx, y+ri*row_h, widths[ci], row_h, fill, LINE)
            text(slide, xx+.08, y+ri*row_h+.04, widths[ci]-.16, row_h-.06, str(cell), font_size, color, ri==0 and header, FONT, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
            xx += widths[ci]

def add_slide():
    s=prs.slides.add_slide(prs.slide_layouts[6]); set_bg(s); return s

# 1 title
s=add_slide();
text(s,.7,.65,2.4,.25,'EQUITY RESEARCH / INDIA AUTO COMPONENTS',10,GREEN,True)
text(s,.7,1.35,9.2,1.2,'Minda Corporation',38,INK,True,FONT_HEAD)
text(s,.7,2.5,8.8,.7,'A quality franchise, but the price already reflects much of the recovery.',22,MUTED,False,FONT_HEAD)
box(s,.7,4.15,4.1,1.15,GREEN,GREEN)
text(s,.95,4.38,3.55,.25,'RECOMMENDATION',9,GREEN_LIGHT,True)
text(s,.95,4.68,3.55,.42,'HOLD / ACCUMULATE ON DIPS',20,WHITE,True,FONT_HEAD)
text(s,9.1,5.75,3.5,.3,'Prepared for investment committee review',10,MUTED,False,FONT,PP_ALIGN.RIGHT)
text(s,9.1,6.08,3.5,.3,'9 September 2026',10,MUTED,False,FONT,PP_ALIGN.RIGHT)
footer(s,1,'Investment view for discussion only. Not investment advice.')

# 2 business
s=add_slide(); title(s,'01 / Business model','A diversified auto-component platform with wiring at its core','Minda earns from content per vehicle, not from vehicle volumes alone.')
box(s,.7,2.18,3.2,3.85,GREEN,GREEN)
text(s,.98,2.48,2.6,.26,'FY25 REVENUE MIX',10,GREEN_LIGHT,True)
text(s,.98,2.92,2.5,.75,'₹4,700 cr',29,WHITE,True,FONT_HEAD)
text(s,.98,3.7,2.4,.35,'approx. consolidated revenue',11,GREEN_LIGHT)
text(s,.98,4.35,2.35,1.2,'Wiring harness 32%\nInterior & electronics 28%\nSwitches / lighting / others 40%',13,WHITE)
text(s,4.25,2.22,3.55,.25,'HOW THE MODEL WORKS',10,GREEN,True)
steps=[('1','Win platform business','Design-in with OEMs creates sticky, multi-year programs.'),('2','Scale content','Localization and volume improve purchasing leverage and utilization.'),('3','Compound adjacencies','Electronics, sensors and EV content raise value per vehicle.')]
for i,(num,head,body) in enumerate(steps):
    yy=2.7+i*1.05
    text(s,4.25,yy,0.35,.35,num,16,GREEN,True,FONT_HEAD)
    text(s,4.75,yy,2.8,.25,head,14,INK,True)
    text(s,4.75,yy+.3,2.85,.45,body,10,MUTED)
box(s,8.25,2.18,4.45,3.85,RGBColor(241,240,233),LINE)
text(s,8.55,2.48,3.7,.25,'FINANCIAL SHAPE',10,GREEN,True)
rows=[['Metric','FY23','FY25'],['Revenue','₹3,800 cr','₹4,700 cr'],['EBITDA margin','10.2%','11.1%'],['ROCE','15.6%','17.8%'],['Net debt / EBITDA','1.6x','1.4x']]
table(s,8.55,2.88,[1.85,1.05,1.05],rows,.42,True,9)
text(s,8.55,5.05,3.55,.65,'The positives are steady growth and operating leverage. The constraint is that returns remain below a premium industrial benchmark.',11,INK)
footer(s,2)

# 3 industry
s=add_slide(); title(s,'02 / Industry','The runway is attractive; competition keeps the economics honest','Indian auto-components are moving from mechanical content to electronics-heavy systems.')
# left narrative
text(s,.72,2.18,5.25,.3,'WHAT IS CHANGING',10,GREEN,True)
bullet_list(s,.72,2.62,5.2,['Premiumisation increases content per vehicle.','EVs remove some legacy parts but add sensors, electronics and harness complexity.','OEM localisation and supply-chain de-risking favour scaled Indian vendors.','Passenger vehicles and 2W exports broaden the addressable market.'],13,gap=.63)
# right chart
box(s,6.45,2.2,6.18,3.7,WHITE,LINE)
text(s,6.75,2.48,5.5,.25,'INDUSTRY SIGNALS',10,GREEN,True)
for label,val,fill in [('Content / vehicle','High',GREEN),('Localisation tailwind','High',GREEN),('Pricing power','Medium',GOLD),('Competition','High',RUST)]:
    yy=3.02+['Content / vehicle','Localisation tailwind','Pricing power','Competition'].index(label)*.62
    text(s,6.75,yy,2.5,.25,label,12,INK,True)
    box(s,9.35,yy+.03,2.6,.18,RGBColor(231,230,221),RGBColor(231,230,221))
    w={'High':2.2,'Medium':1.45}[val]
    box(s,9.35,yy+.03,w,.18,fill,fill)
    text(s,12.0,yy-.02,.4,.25,val,10,fill,True,align=PP_ALIGN.RIGHT)
box(s,6.75,5.3,5.55,.38,RUST_LIGHT,RUST_LIGHT)
text(s,6.92,5.38,5.2,.2,'Bottom line: secular growth does not guarantee superior returns.',10,RUST,True)
footer(s,3)

# 4 peers
s=add_slide(); title(s,'03 / Relative valuation','MSUMI closes the most important gap in the peer set','Peer median is used as the reference line; round-number cutoffs are avoided.')
rows=[['Company','Why it belongs','Rev growth','EBITDA mgn','ROCE','P/E'],['Minda Corp','Target: diversified auto-electricals','12%','11.1%','17.8%','34x'],['MSUMI','Pure-play wiring harness','15%','13.2%','23.5%','42x'],['Lumax Ind.','Lighting / mechatronics','9%','10.5%','18.9%','31x'],['Fiem Ind.','Lighting and signalling','11%','12.1%','20.8%','27x'],['Peer median','Reference point','11%','12.1%','20.8%','31x']]
table(s,.72,2.15,[1.45,3.15,1.15,1.25,1.0,0.75],rows,.48,True,9)
text(s,.75,5.45,7.35,.65,'Interpretation: Minda is not the cheapest on earnings, while its return profile sits below the peer median. MSUMI deserves a premium for purity, but also sets a higher bar for execution.',11,INK)
box(s,8.55,5.0,3.95,1.05,GREEN_LIGHT,GREEN_LIGHT)
text(s,8.8,5.2,3.45,.22,'PEER-SET LIMIT',9,GREEN,True)
text(s,8.8,5.48,3.35,.38,'Listed pure-play wiring peers are scarce; adjacent lighting players are included with lower weight.',10,INK)
footer(s,4,'Multiples are indicative, rounded and should be refreshed from the latest filings / market close before publication.')

# 5 quadrant
s=add_slide(); title(s,'03 / Relative valuation','Minda sits near the middle of the quality–valuation map','The reference lines are peer medians, not reverse-engineered thresholds.')
box(s,.92,2.2,6.4,3.85,WHITE,LINE)
# axes
line=s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.55), Inches(5.45), Inches(4.95), Inches(.012)); line.fill.solid(); line.fill.fore_color.rgb=INK; no_line(line)
line=s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.55), Inches(2.78), Inches(.012), Inches(2.68)); line.fill.solid(); line.fill.fore_color.rgb=INK; no_line(line)
text(s,3.1,5.68,3.3,.25,'P/E → higher valuation',9,MUTED)
text(s,.96,3.05,.35,1.3,'ROCE\n↑',10,MUTED,True,align=PP_ALIGN.CENTER)
# median dotted-ish
for x in [3.8]:
    l=s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(2.78), Inches(.012), Inches(2.68)); l.fill.solid(); l.fill.fore_color.rgb=GOLD; no_line(l)
for y in [4.1]:
    l=s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.55), Inches(y), Inches(4.95), Inches(.012)); l.fill.solid(); l.fill.fore_color.rgb=GOLD; no_line(l)
for name,x,y,c in [('MSUMI',5.3,3.38,GREEN),('Fiem',3.15,3.7,GREEN),('Lumax',3.55,4.0,GREEN),('Minda',4.2,4.35,RUST)]:
    dot=s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(.28), Inches(.28)); dot.fill.solid(); dot.fill.fore_color.rgb=c; dot.line.color.rgb=WHITE
    text(s,x+.35,y-.01,1.0,.25,name,10,INK,True)
text(s,1.65,2.95,1.8,.25,'Higher return / cheaper',9,GREEN,True)
text(s,4.75,2.95,1.8,.25,'Higher return / richer',9,GREEN,True)
text(s,1.65,5.15,1.8,.25,'Lower return / cheaper',9,RUST,True)
text(s,4.75,5.15,1.8,.25,'Lower return / richer',9,RUST,True)
box(s,7.75,2.2,4.8,3.85,RGBColor(241,240,233),LINE)
text(s,8.05,2.5,4.1,.25,'READ-THROUGH',10,GREEN,True)
bullet_list(s,8.05,2.92,4.1,['MSUMI earns its premium through purity and higher ROCE.','Minda has more optionality, but the market already values the recovery.','A rerating needs sustained margin delivery, not just revenue growth.'],12,gap=.78,bullet=GOLD)
footer(s,5,'Peer medians shown for directionality; datapoints are indicative analyst estimates.')

# 6 call
s=add_slide(); title(s,'04 / Investment call','HOLD: a good business, but not a bargain','The stock becomes more interesting when the price gives us room for execution risk.')
box(s,.72,2.2,3.25,3.55,GREEN,GREEN)
text(s,1.0,2.55,2.55,.25,'CALL',10,GREEN_LIGHT,True)
text(s,1.0,2.98,2.7,.6,'HOLD',35,WHITE,True,FONT_HEAD)
text(s,1.0,3.78,2.5,.65,'Accumulate on dips\ninto the ₹290–310 zone.',14,WHITE,True)
text(s,1.0,4.9,2.55,.5,'Upside needs margin expansion and better capital efficiency.',10,GREEN_LIGHT)
text(s,4.45,2.24,3.35,.25,'TARGET-PRICE FRAMEWORK',10,GREEN,True)
rows=[['Case','FY27E EPS','P/E','Value'],['Bear','₹10.5','28x','₹294'],['Base','₹11.5','31x','₹357'],['Bull','₹12.2','33x','₹403']]
table(s,4.45,2.7,[1.0,1.0,.8,1.0],rows,.48,True,10)
text(s,4.48,4.72,3.45,.7,'Base case implies a fair-value range of ₹340–370. The range is wide because the multiple, not only the earnings, is doing the work.',11,INK)
box(s,8.55,2.2,4.0,3.55,RGBColor(241,240,233),LINE)
text(s,8.85,2.5,3.4,.25,'WHAT WOULD CHANGE THE CALL',10,GREEN,True)
bullet_list(s,8.85,2.95,3.35,['BUY: ROCE clears 20% and EBITDA margin sustains above 12%.','SELL: leverage rises while growth falls below high single digits.','WATCH: EV wins, export mix, and 2W concentration.'],12,gap=.78,bullet=RUST)
footer(s,6,'Analyst framework; not a live price target. Refresh with current price, shares outstanding and latest earnings before use.')

# 7 risks
s=add_slide(); title(s,'05 / Risks','The debate is execution, not the existence of demand','Risk is manageable when monitored against explicit operating triggers.')
rows=[['Risk','Probability','Impact','Monitor / trigger','Mitigation'],['2W concentration','High','High','2W mix; OEM volumes','Broaden PV / exports'],['Leverage & capex','Medium','High','Net debt / EBITDA >2x','Stage-gate investments'],['Low ROCE','Medium','High','ROCE below 17%','Mix + working-capital discipline'],['EV execution','Medium','Medium','Launch delays / win loss','Pilot before scale'],['Input costs / competition','Medium','Medium','Gross margin compression','Pass-through + sourcing']]
table(s,.72,2.15,[1.7,1.0,.9,3.0,2.3],rows,.58,True,8.7)
box(s,.75,5.75,11.7,.45,RUST_LIGHT,RUST_LIGHT)
text(s,.95,5.86,11.2,.2,'Key watch item for the next two results: can margin gains outrun the capital required to fund the next leg of growth?',11,RUST,True)
footer(s,7)

# 8 appendix
s=add_slide(); title(s,'Appendix','Definitions and analyst notes','A short audit trail makes the recommendation easier to defend.')
box(s,.72,2.15,5.75,3.9,WHITE,LINE)
text(s,1.0,2.45,5.1,.25,'METHOD',10,GREEN,True)
bullet_list(s,1.0,2.87,5.0,['Peer set is selected for component exposure, listed status and Indian operating context.','Peer median is used for the valuation map divider.','Target value = forward EPS × selected P/E band.','Figures are rounded; estimates are clearly labelled as analyst-derived.'],12,gap=.68)
box(s,6.75,2.15,5.75,3.9,RGBColor(241,240,233),LINE)
text(s,7.03,2.45,5.1,.25,'DEFINITIONS',10,GREEN,True)
bullet_list(s,7.03,2.87,5.0,['ROCE: operating profit relative to capital employed.','P/E: market price divided by forward earnings per share.','EV/EBITDA: enterprise value relative to operating cash earnings.','Pure-play: company with a more concentrated exposure to the relevant product.'],12,gap=.68,bullet=GOLD)
footer(s,8)

# 9 references
s=add_slide(); title(s,'References','Sources used and data caveats','The deck is designed to be refreshed, not treated as a static data dump.')
text(s,.8,2.25,11.5,.45,'Primary sources',14,GREEN,True)
bullet_list(s,.8,2.85,11.5,['Minda Corporation annual report and investor presentations, FY23–FY25.','Company filings and exchange disclosures for segment mix, debt and operating commentary.','Peer company annual reports and investor presentations: MSUMI, Lumax Industries and Fiem Industries.','Market multiples and prices: indicative snapshot; verify against the latest exchange close / Screener or equivalent before publication.'],13,gap=.72)
box(s,.8,5.75,11.5,.55,RUST_LIGHT,RUST_LIGHT)
text(s,1.05,5.89,11.0,.25,'Important: valuation figures are rounded and partly analyst-estimated. They are not a substitute for current market data or independent diligence.',10,RUST,True)
footer(s,9,'Data as of 9 September 2026 for presentation purposes; confirm all figures before external circulation.')

# 10 thanks
s=add_slide();
text(s,.75,.75,4.2,.25,'MINDA CORPORATION / INVESTMENT CASE',10,GREEN,True)
text(s,.75,2.1,7.8,.8,'Thank you.',38,INK,True,FONT_HEAD)
text(s,.78,3.05,7.8,.55,'Questions, challenges and what would make us change the call.',18,MUTED,False,FONT_HEAD)
box(s,.78,4.6,3.8,.9,GREEN,GREEN)
text(s,1.03,4.83,3.3,.28,'HOLD / ACCUMULATE ON DIPS',15,WHITE,True,FONT_HEAD)
text(s,9.4,6.35,3.0,.3,'Investment committee review',10,MUTED,False,FONT,PP_ALIGN.RIGHT)
footer(s,10,'For discussion only. Please verify all market data and assumptions before relying on this presentation.')

prs.save(OUT)
print(f'Wrote {OUT} with {len(prs.slides)} slides')
