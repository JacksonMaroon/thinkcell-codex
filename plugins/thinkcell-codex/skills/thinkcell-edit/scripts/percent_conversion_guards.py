"""Saved native data and binding gates for the bounded percentage converter."""
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import re
from chart_geometry import inventory, xml
from multi_chart_update import cells_matrix, model_of

def need(condition, message):
    if not condition: raise ValueError(message)

def write_new(path, content):
    with Path(path).open('xb') as handle: handle.write(content)

def ratio(matrix):
    need(len(matrix)==6 and all(len(row)==4 for row in matrix),'Unsupported 4-series by 3-category profile')
    need(matrix[0][1:]==['Gross Profit','Revenues','Customer Count'] and [r[0] for r in matrix[2:]]==['Not Mapped','High risk','Medium Risk','Low risk'],'Unsupported category or series order')
    numerator=Decimal(str(matrix[5][1]));denominator=Decimal(str(matrix[1][1]))
    need(numerator.is_finite() and denominator.is_finite() and numerator>=0 and denominator>0,'Finite nonnegative numerator and positive denominator required')
    return numerator,denominator,f"{(numerator*100/denominator).quantize(Decimal('1'),rounding=ROUND_HALF_UP)}%"

def binding(state):
    text=state['sources']['m_varsrcRelative']['text_variable_xml']
    need(text,'Native relative binding missing')
    node=xml(text.encode())
    active=lambda n:int(n.get('reqver','0'))<=38764<int(n.get('endver','999999'))
    formats=[n for n in node.findall('m_bstrFormat') if active(n)]
    precisions=[n for n in node.findall('m_prec17834') if active(n)]
    need(len(formats)==len(precisions)==1 and formats[0].text,'Ambiguous relative format')
    digits=precisions[0].find('m_nDecimalDigits17909')
    need(precisions[0].findtext('m_strSuffix17909')=='%' and digits is not None and digits.get('val')=='0','Only zero-decimal percent precision is verified')
    return 'datetime'+formats[0].text

def check(state, expected, phase, require_unbound=False):
    shapes=state['visible_shapes'];need(len(shapes)==1,phase+': selected label missing/ambiguous')
    fields=shapes[0]['fields'];need(len(fields)==1 and fields[0]['text']==expected,phase+': expected one native '+expected)
    need(shapes[0]['literal_texts']==['(',expected,')'],phase+': wrappers changed')
    need(fields[0]['type']==binding(state),phase+': relative field binding mismatch')
    if require_unbound:
        variable=xml(state['sources']['m_varsrcAbsolute']['xml'].encode()).find('m_ctextvar')
        need(variable is not None and len(variable)==0,phase+': absolute text binding remains')

def dual(state):
    shapes=state['visible_shapes'];need(len(shapes)==1,'Selected dual label missing')
    fields=shapes[0]['fields'];need(len(fields)==2,'Two original fields required')
    absolute,relative=fields
    need(shapes[0]['literal_texts']==[absolute['text'],'(',relative['text'],')'],'Unsupported dual grammar')
    need(re.fullmatch(r'\d+',absolute['text']) and re.fullmatch(r'\d+%',relative['text']),'Only nonnegative integer dual text is verified')
    source=state['sources']['m_varsrcAbsolute']['text_variable_xml'];need(source,'Absolute binding missing')
    need(absolute['type']=='datetime'+xml(source.encode()).findtext('m_bstrFormat'),'Absolute field binding mismatch')
    need(relative['type']==binding(state),'Relative field binding mismatch')
    return absolute['text'],relative['text']

def saved_ratio(path,state):
    _,charts,_=inventory(Path(path).read_bytes());need(len(charts)==1,'Exactly one native chart required')
    chart=charts[0]
    need(chart['doc']['part']=='ppt/embeddings/oleObject13.bin' and chart['owner'].tag=='CSequenceChartSE' and chart['owner'].find('m_ect').get('val')=='0','Unsupported chart profile')
    need(chart['doc']['root'].find('version').get('val')=='38764','Unsupported native model version')
    actual=ratio(cells_matrix(chart)[0]);model=model_of(chart)
    need(model['percent_axis'] and model['series_names']==['Not Mapped','High risk','Medium Risk','Low risk'] and model['categories']==['Gross Profit','Revenues','Customer Count'],'Unsupported model semantics')
    need(Decimal(str(model['series_values'][3][0]))==actual[0] and Decimal(str(model['category_extents'][0]))==actual[1],'Saved model and datasheet ratio differ')
    scalar=xml(state['scalar_xml'].encode());need(scalar.get('id')=='191','Selected scalar changed')
    value=xml(state['sources']['m_varsrcAbsolute']['xml'].encode()).find('m_varval')
    need(value is not None and Decimal(value.get('val'))==actual[0],'Selected scalar differs from saved numerator')
    label=xml(state['label_xml'].encode())
    need(label.findtext('m_ppttb/m_bstrShapeName')=='tbUpiE_yCia4NQkLBVWFCIA','Selected label tag changed')
    return actual
