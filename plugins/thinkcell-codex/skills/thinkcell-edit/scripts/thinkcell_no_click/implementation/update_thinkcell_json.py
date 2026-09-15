"""Single-slide, guarded JSON update runner. Dry-run default; native certification remains separate."""
from pathlib import Path
import argparse,collections,io,json,math,os,subprocess,sys,zipfile
from lxml import etree as E
from prepare_thinkcell_name import prepare,inventory,choose,link_contract,need,sha,xml,SKILL,NS,logical_slides,relationship_map,pie_tables,sequence_tables,scatter_tables,inspect_presentation
from read_named_datasheet import read as read_datasheet, read_blob
from extract_thinkcell_named_datasheet import wrap_biff_workbook_stream
from audit_thinkcell_integrity import notes_substance
from runtime import find_ppttc
PPTTC=None

def equal(a,b):
 if isinstance(a,bool) or isinstance(b,bool):return type(a)==type(b) and a==b
 if isinstance(a,(int,float)) and isinstance(b,(int,float)):return math.isclose(a,b,rel_tol=1e-10,abs_tol=1e-9)
 if isinstance(a,list) and isinstance(b,list):return len(a)==len(b) and all(equal(x,y) for x,y in zip(a,b))
 if isinstance(a,dict) and isinstance(b,dict):return set(a)==set(b) and all(equal(a[k],b[k]) for k in a)
 return a==b

def validate_request(request,family):
 need(isinstance(request,dict) and set(request)=={'matrix','expected_model'},'Data JSON requires exactly matrix and expected_model')
 matrix=request['matrix'];need(isinstance(matrix,list) and matrix and all(isinstance(r,list) for r in matrix),'Matrix must contain rows')
 width=len(matrix[0]);need(width>0 and all(len(r)==width for r in matrix),'Matrix must be rectangular')
 need(any(v not in (None,'') for r in matrix for v in r),'Empty matrix rejected')
 for row in matrix:
  for v in row:
   need(not isinstance(v,bool) and (v is None or isinstance(v,(str,int,float))),'Cells must be null, string, or number; JSON boolean cells are unsupported')
   if isinstance(v,(float,int)) and not isinstance(v,bool):need(math.isfinite(v),'Nonfinite value rejected')
 need(family in {'CPieChartSE','CSequenceChartSE','CScatterChartSE'},'Unsupported JSON target family')
 expected=request['expected_model'];required={'categories','values'} if family=='CPieChartSE' else ({'point_labels','group_labels','x_values','y_values','size_values'} if family=='CScatterChartSE' else {'categories','series_names','series_values'})
 need(isinstance(expected,dict) and (set(expected)==required or (family=='CSequenceChartSE' and set(expected) in [required|{'column_widths'},required|{'category_extents'}])),'Expected model must provide exact categories and full raw values/series')
 need(all(isinstance(expected[k],list) for k in required),'Expected model fields must be lists')
 if family=='CSequenceChartSE':
  need(all(isinstance(n,str) for n in expected['series_names']),'Expected series names must be strings')
  need(len(set(expected['series_names']))==len(expected['series_names']),'Expected series names must be unique')
  rows=expected['series_values']
  need(len(rows)==len(expected['series_names']) and all(isinstance(r,list) and len(r)==len(expected['categories']) for r in rows),'Expected series values must match category and series dimensions')
  need(all(v is None or (not isinstance(v,bool) and isinstance(v,(int,float)) and math.isfinite(v)) for r in rows for v in r),'Expected series values must be finite numbers or null, including calculated equals totals')
 if family=='CPieChartSE':
  need(len(expected['values'])==len(expected['categories']) and all(v is None or (not isinstance(v,bool) and isinstance(v,(int,float)) and math.isfinite(v)) for v in expected['values']),'Expected pie values must be finite numbers or null and match categories')
 if family=='CScatterChartSE':
  count=len(expected['x_values']);need(count>0 and all(len(expected[k])==count for k in required),'All scatter fields must have matching nonzero point counts')
  need(all(isinstance(v,str) for k in ['point_labels','group_labels'] for v in expected[k]),'Scatter labels/groups must be strings')
  need(all(not isinstance(v,bool) and isinstance(v,(int,float)) and math.isfinite(v) for k in ['x_values','y_values'] for v in expected[k]),'Scatter coordinates must be finite numbers')
  sizes=expected['size_values'];need(all(v is None for v in sizes) or all(not isinstance(v,bool) and isinstance(v,(int,float)) and math.isfinite(v) and v>0 for v in sizes),'Sizes must be all null for scatter or all positive finite numbers for bubble')
 return request

def canonical_sequence_contract(path, request, target=None, name=None):
 """Verify canonical JSON rows using the source's actual datasheet orientation.
 Current bounded contract keeps series/category counts and optional rows fixed.
 """
 data=Path(path).read_bytes();docs,cs,names=inventory(data)
 if target is not None:
  c=choose(cs,slide_id=target['slide_id'],shape_id=target['shape_id'],shape_tag=target['shape_tag'])
 else:
  hits=[c for c in cs if c['owner_name']==name];need(len(hits)==1,'Canonical target ambiguous');c=hits[0]
 if c['owner'].tag=='CScatterChartSE':return canonical_scatter_contract(c,request)
 if c['owner'].tag=='CPieChartSE':return canonical_pie_contract(c,request)
 ds=link_contract(c);storage=ds['storage'];ss=c['doc']['streams']
 meta=ss.get((storage,'think-cellXML'));need(meta is not None,'Missing datasource orientation metadata')
 orient=xml(meta).find('PersistentType/m_eorient');need(orient is not None and orient.get('val') in {'0','1'},'Unknown datasource orientation')
 orientation=int(orient.get('val'))
 if (storage,'Package') in ss:sheets=read_blob(ss[(storage,'Package')],'xlsb_package')
 elif (storage,'Workbook') in ss:sheets=read_blob(wrap_biff_workbook_stream(ss[(storage,'Workbook')]),'legacy_biff_cfb')
 else:raise ValueError('Unsupported datasource storage')
 need(len(sheets)==1,'Expected one source datasheet');cells=sheets[0]['nonempty_cells'];need(cells,'Empty source datasheet')
 rows=max(x['row'] for x in cells);cols=max(x['column'] for x in cells);source=[[None]*cols for _ in range(rows)]
 for x in cells:source[x['row']-1][x['column']-1]=x['value']
 canonical=[list(r) for r in zip(*source)] if orientation==1 else source
 baseline=model_of(c);expected=request['expected_model'];matrix=request['matrix']
 from chart_semantics import kind,details
 chart_kind=kind(c);semantic=details(c)
 def row_matches(raw,calculated):
  if len(raw)!=len(calculated):return False
  return all((chart_kind=='waterfall' and v=='e' and isinstance(w,(int,float)) and math.isfinite(w)) or equal(v,w) for v,w in zip(raw,calculated))
 need(equal(canonical[0][1:],baseline['categories']),'Source canonical categories disagree with model')
 need(len(matrix)==len(canonical) and all(len(r)==len(canonical[0]) for r in matrix),'JSON matrix must use canonical sequence orientation and fixed slot dimensions')
 need(equal(matrix[0][1:],expected['categories']),'JSON category row does not match expected model')
 names0=baseline['series_names'];need(len(set(names0))==len(names0),'Duplicate source series labels require a richer contract')
 need(len(expected['series_names'])==len(names0) and len(expected['series_values'])==len(names0),'Series count changes need a separate layout contract')
 need(len(expected['categories'])==len(baseline['categories']),'Category count changes need a separate layout contract')
 matched=[]
 for ri,row in enumerate(canonical[1:],1):
  if row[0] in names0:
   i=names0.index(row[0]);slot=len(matched);matched.append(i)
   need(row_matches(row[1:],baseline['series_values'][i]),'Source canonical series values disagree with model')
   need(matrix[ri][0]==expected['series_names'][i] and row_matches(matrix[ri][1:],expected['series_values'][i]),'JSON series row order/values do not match expected model at slot '+str(ri+1))
   if chart_kind=='waterfall':need([v=='e' for v in row[1:]]==[v=='e' for v in matrix[ri][1:]],'Waterfall equals slots must stay fixed')
  elif chart_kind=='CSequenceChartSE' and ri==1 and all(isinstance(v,(int,float)) for v in row[1:]):
   need(equal(row[1:],baseline['category_extents']),'Source 100%= row disagrees with model')
   need(matrix[ri][0]==row[0] and equal(matrix[ri][1:],expected.get('category_extents')),'Explicit 100%= row must match expected category_extents')
   need(all(not isinstance(v,bool) and isinstance(v,(int,float)) and math.isfinite(v) and v>0 for v in matrix[ri][1:]),'Explicit 100%= values must be positive')
  elif chart_kind=='mekko-units' and ri==1:
   need(equal(row[1:],baseline['column_widths']),'Source Mekko width row disagrees with model')
   need(matrix[ri][0]==row[0] and equal(matrix[ri][1:],expected.get('column_widths')),'Mekko X extent row must match expected column_widths')
  else:need(equal(matrix[ri],row),'Optional/reserved datasource row changed without explicit contract')
 need(sorted(matched)==list(range(len(names0))),'Source series slots unresolved')
 if chart_kind.startswith('mekko'):
  widths=expected.get('column_widths');need(isinstance(widths,list) and len(widths)==len(expected['categories']) and all(not isinstance(v,bool) and isinstance(v,(int,float)) and math.isfinite(v) and v>0 for v in widths),'Mekko requires positive expected column_widths')
  need(all(v is None or (not isinstance(v,bool) and isinstance(v,(int,float)) and math.isfinite(v) and v>=0) for row in expected['series_values'] for v in row),'Mekko contract supports nonnegative absolute data only')
  if chart_kind=='mekko-percent':
   need(all(v is None for v in canonical[1]),'Mekko percent-input/100%= mode requires a separate contract')
   need(equal(widths,[sum(row[i] or 0 for row in expected['series_values']) for i in range(len(widths))]),'Percent Mekko widths must equal absolute category totals')
 else:need('column_widths' not in expected,'column_widths is only valid for Mekko')
 return {'family':chart_kind,'semantics':semantic,'source_datasheet_orientation':orientation,'json_orientation':'series_rows_category_columns','source_was_transposed_for_contract':orientation==1,'slot_dimensions':[len(canonical),len(canonical[0])],'fixed_optional_rows_preserved':True,'source_series_row_order':[r[0] for r in canonical[1:] if r[0] in names0],'json_series_row_order':[matrix[ri][0] for ri,row in enumerate(canonical) if ri>0 and row[0] in names0],'source_slots_define_layout_not_series_order':True}

def canonical_pie_contract(c,request):
 """Pie JSON uses category rows, unlike a horizontal embedded datasheet."""
 baseline=model_of(c);expected=request['expected_model'];matrix=request['matrix']
 need(len(expected['categories'])==len(baseline['categories']) and len(expected['values'])==len(baseline['values']),'Pie category count changes require a separate layout contract')
 need(len(matrix)==len(expected['categories'])+1 and all(len(row)==2 for row in matrix),'Pie JSON requires a header row and category/value rows')
 need(equal([row[0] for row in matrix[1:]],expected['categories']) and equal([row[1] for row in matrix[1:]],expected['values']),'Pie JSON rows disagree with expected categories/values')
 ds=link_contract(c);storage=ds['storage'];ss=c['doc']['streams']
 if (storage,'Package') in ss:sheets=read_blob(ss[(storage,'Package')],'xlsb_package')
 elif (storage,'Workbook') in ss:sheets=read_blob(wrap_biff_workbook_stream(ss[(storage,'Workbook')]),'legacy_biff_cfb')
 else:raise ValueError('Unsupported pie datasource storage')
 need(len(sheets)==1,'Expected one pie datasheet');cells=sheets[0]['nonempty_cells'];need(cells,'Empty pie datasheet')
 raw=[[None]*max(x['column'] for x in cells) for _ in range(max(x['row'] for x in cells))]
 for x in cells:raw[x['row']-1][x['column']-1]=x['value']
 options=[raw,[list(row) for row in zip(*raw)]]
 matching=[rows for rows in options if len(rows)==len(baseline['categories'])+1 and all(len(row)==2 for row in rows) and equal([row[0] for row in rows[1:]],baseline['categories']) and equal([row[1] for row in rows[1:]],baseline['values'])]
 need(matching and all(equal(rows,matching[0]) for rows in matching),'Source pie layout is unsupported or ambiguous')
 need(equal(matrix[0],matching[0][0]),'Pie header row changed without explicit contract')
 need(baseline.get('input_mode','absolute')!='mixed','Mixed absolute/percentage pie cells need a separate contract')
 if baseline.get('input_mode')=='percentage':need(all(isinstance(v,(int,float)) and not isinstance(v,bool) and v>=0 for v in expected['values']) and equal(sum(expected['values']),1) and baseline['total']==100,'Percentage pie values must be nonnegative fractions summing to 1, with source total 100')
 return {'family':'pie','input_mode':baseline.get('input_mode','absolute'),'json_orientation':'category_rows_label_value_columns','fixed_category_count':len(expected['categories']),'header_preserved':True}

def canonical_scatter_contract(c,request):
 """Five-column rows: label, X, Y, size, group. Preserve reserved rows and fixed point slots.
 Resolve row/column orientation from exact baseline values, never sequence metadata.
 """
 ds=link_contract(c);storage=ds['storage'];ss=c['doc']['streams']
 if (storage,'Package') in ss:sheets=read_blob(ss[(storage,'Package')],'xlsb_package')
 elif (storage,'Workbook') in ss:sheets=read_blob(ss[(storage,'Workbook')],'legacy_biff_cfb')
 else:raise ValueError('Unsupported scatter datasource storage')
 need(len(sheets)==1,'Expected one scatter datasheet');cells=sheets[0]['nonempty_cells'];need(cells,'Empty source datasheet')
 rows=max(x['row'] for x in cells);cols=max(x['column'] for x in cells);raw=[[None]*cols for _ in range(rows)]
 for x in cells:raw[x['row']-1][x['column']-1]=x['value']
 baseline=model_of(c);expected=request['expected_model'];fields=['point_labels','x_values','y_values','size_values','group_labels']
 baseline_rows=[list(v) for v in zip(*(baseline[k] for k in fields))]
 expected_rows=[list(v) for v in zip(*(expected[k] for k in fields))]
 need(len(expected_rows)==len(baseline_rows),'Scatter point-count changes need a separate layout contract')
 need(all((a is None)==(b is None) for a,b in zip(baseline['size_values'],expected['size_values'])),'Changing scatter/bubble kind is outside this contract')
 candidates=[]
 for transposed,source in [(False,raw),(True,[list(r) for r in zip(*raw)])]:
  if len(source[0])!=5:continue
  slots=[i for i,r in enumerate(source) if any(equal(r,b) for b in baseline_rows)]
  if len(slots)==len(baseline_rows) and equal([source[i] for i in slots],baseline_rows):candidates.append((transposed,source,slots))
 need(len(candidates)==1,'Scatter datasource orientation/point slots unresolved or ambiguous')
 transposed,source,slots=candidates[0];matrix=request['matrix']
 need(len(matrix)==len(source) and all(len(r)==5 for r in matrix),'Scatter JSON requires fixed five-column point-row layout')
 for i,row in enumerate(source):
  if i in slots:need(equal(matrix[i],expected_rows[slots.index(i)]),'Scatter JSON point row disagrees with expected model at row '+str(i+1))
  else:need(equal(matrix[i],row),'Scatter header/reserved row changed without explicit contract')
 return {'family':'bubble' if any(v is not None for v in baseline['size_values']) else 'scatter','json_orientation':'point_rows_label_x_y_size_group_columns','source_was_transposed_for_contract':transposed,'point_rows_1based':[i+1 for i in slots],'fixed_optional_rows_preserved':True,'point_count':len(slots),'size_semantics_preserved':baseline['z_value_is_area']}

def payload(matrix,contract=None):
 def cell(v):
  if v is None:return None
  if isinstance(v,bool):return {'boolean':v}
  if isinstance(v,(int,float)):return {'number':v}
  return {'string':v}
 result=[[cell(v) for v in row] for row in matrix]
 if contract and contract.get('family')=='pie' and contract.get('input_mode')=='percentage':
  for i,row in enumerate(matrix[1:],1):
   if row[1] is not None:result[i][1]={'percentage':row[1]*100}
 return result

def model_of(c):
 tables={'CPieChartSE':pie_tables,'CSequenceChartSE':sequence_tables,'CScatterChartSE':scatter_tables}[c['owner'].tag](c['doc']['root'])
 hits=[x for x in tables if str(x['id'])==c['table'].get('id')];need(len(hits)==1,'Ambiguous model')
 from chart_semantics import kind,details
 if kind(c).startswith('mekko'):hits[0]['column_widths']=details(c)['column_widths']
 return hits[0]

def notes_prose(root):
 """Retain paragraph/text boundaries; omit only cached slidenum field text."""
 paragraphs=[]
 for paragraph in root.findall('.//a:p',NS):
  chunks=[]
  for text in paragraph.findall('.//a:t',NS):
   if any(E.QName(a).localname=='fld' and a.get('type','').lower()=='slidenum' for a in text.iterancestors()):continue
   chunks.append(text.text or '')
  paragraphs.append(''.join(chunks))
 return paragraphs

def snapshot(path,name):
 data=path.read_bytes();docs,cs,names=inventory(data)
 need(not [n.tag for d in docs for n in d['root'] if (n.tag.endswith('ChartSE') or n.tag=='CGanttSE') and n.tag not in {'CPieChartSE','CSequenceChartSE','CScatterChartSE'}],'Unsupported sibling chart family on slide')
 with zipfile.ZipFile(io.BytesIO(data)) as z:
  slides=logical_slides(z);need(len(slides)==1,'Runner requires exactly one slide')
  rel=relationship_map(z,slides[0]['part']);layout=next(x['resolved'] for x in rel.values() if x['type'].endswith('/slideLayout'));master=next(x['resolved'] for x in relationship_map(z,layout).values() if x['type'].endswith('/slideMaster'));theme=next(x['resolved'] for x in relationship_map(z,master).values() if x['type'].endswith('/theme'))
  branding={'theme_sha256':sha(z.read(theme)),'theme_name':xml(z.read(theme)).get('name'),'notes_prose_except_slide_number':[notes_prose(xml(z.read(x['resolved']))) for x in rel.values() if x['type'].endswith('/notesSlide')],'notes_substance_sha256':notes_substance(z,slides[0]['part'])}
 need(all(c['owner'].tag in {'CPieChartSE','CSequenceChartSE','CScatterChartSE'} for c in cs),'Unsupported sibling element family on slide')
 need(all(c['exact'] for c in cs),'Unresolved chart ownership on slide')
 need(len({c['frames'][0]['shape_tag'] for c in cs})==len(cs),'Ambiguous native chart ownership')
 hits=[c for c in cs if c['owner_name'].casefold()==name.casefold()];need(len(hits)==1,'Named target not unique')
 target=hits[0];siblings=[]
 target_ds=link_contract(target);target_meta=target['doc']['streams'].get((target_ds['storage'],'think-cellXML'));orientation_node=xml(target_meta).find('PersistentType/m_eorient') if target_meta else None
 orientation=orientation_node.get('val') if orientation_node is not None else None
 for c in cs:
  ds=link_contract(c)
  if c is target:continue
  logical=model_of(c);logical={k:v for k,v in logical.items() if k not in {'id','chart_id','storage'}}
  siblings.append({'family':c['owner'].tag,'data':logical,'datasource_streams':ds['stream_sha256']})
 # Compare the complete sibling multiset, not IDs that native JSON may regenerate.
 siblings=collections.Counter(json.dumps(s,sort_keys=True) for s in siblings)
 from chart_semantics import details
 return {'target_semantics':details(target),'branding':branding,'target_model':model_of(target),'target_family':target['owner'].tag,'datasheet_orientation':orientation,'target_native_frame':target['frames'][0],'siblings':dict(siblings),'source_sha256':sha(data)}

def validate_output(prepared_path,output_path,automation_name,request):
 """Read-only callable for externally generated/native-reopened candidates."""
 prepared_path=Path(prepared_path);output_path=Path(output_path)
 before=prepared_path.read_bytes();after=output_path.read_bytes()
 base=snapshot(prepared_path,automation_name);actual=snapshot(output_path,automation_name)
 validate_request(request,base['target_family']);contract=canonical_sequence_contract(prepared_path,request,name=automation_name);need(actual['target_family']==base['target_family'],'Target family changed')
 need(actual['target_native_frame']['shape_tag']==base['target_native_frame']['shape_tag'],'Resolved target think-cell tag changed; refusing retargeting')
 need(actual['branding']==base['branding'],'Bound theme or notes changed')
 need(actual['siblings']==base['siblings'],'Untargeted chart model or datasource changed')
 before_sem={k:v for k,v in base['target_semantics'].items() if k!='column_widths'};after_sem={k:v for k,v in actual['target_semantics'].items() if k!='column_widths'}
 need(before_sem==after_sem,'Chart kind, waterfall equals slots, connectors or grounding changed')
 if actual['target_family']=='CPieChartSE':
  need(all(actual['target_model'].get(k)==base['target_model'].get(k) for k in ['input_mode','hole_percent','exploded']),'Pie input mode, hole size or exploded slices changed')
 if actual['target_family']=='CScatterChartSE':need(actual['target_model']['z_value_is_area']==base['target_model']['z_value_is_area'],'Bubble size interpretation changed')
 if actual['target_family']=='CSequenceChartSE':
  expected_series=dict(zip(request['expected_model']['series_names'],request['expected_model']['series_values']))
  actual_series=dict(zip(actual['target_model']['series_names'],actual['target_model']['series_values']))
  need(len(actual_series)==len(actual['target_model']['series_names']) and equal(actual_series,expected_series),'Expected target series data mismatch')
 for k,v in request['expected_model'].items():
  if actual['target_family']=='CSequenceChartSE' and k in {'series_names','series_values'}:continue
  need(equal(actual['target_model'].get(k),v),'Expected target model mismatch: '+k)
 from chart_semantics import audit_scope
 audit=inspect_presentation(output_path,True);integrity_scope=audit_scope(output_path,audit)
 sheet=read_datasheet(output_path,1,automation_name);need(len(sheet['sheets'])==1,'Expected one target datasheet')
 cells=sheet['sheets'][0]['nonempty_cells'];need(all(c.get('cell_type')!=5 for c in cells),'Datasource contains Excel errors')
 actual_cells={(c['row'],c['column']):c['value'] for c in cells}
 if actual['target_family']=='CSequenceChartSE':
  need(actual['datasheet_orientation'] in {'0','1'},'Unknown generated datasheet orientation')
  if actual['datasheet_orientation']=='1':actual_cells={(col,row):v for (row,col),v in actual_cells.items()}
 expected={(r+1,c+1):v for r,row in enumerate(request['matrix']) for c,v in enumerate(row) if v not in (None,'')}
 need(set(actual_cells)==set(expected),'Datasource nonempty cell coordinates differ')
 need(all(equal(actual_cells[k],v) for k,v in expected.items()),'Datasource cell values differ')
 need(prepared_path.read_bytes()==before and output_path.read_bytes()==after,'Input changed during validation')
 return {'status':'JSON_OUTPUT_VALIDATED_PENDING_NATIVE_CERTIFICATION','output_sha256':sha(after),'resolved_target_tag_preserved':True,'actual_datasheet_nonempty_cells_match':True,'expected_logical_model_matches':True,'integrity_scope':integrity_scope,'strict_model_native_cache_integrity':integrity_scope['native_cache_parity']=='pass','untargeted_chart_data_and_streams_unchanged':True,'visual_geometry_requires_native_render_check':True,'bound_theme_and_notes_unchanged':True,'notes_preservation_scope':'Exact authored substance and paragraph prose; only cached slidenum field text excluded','target_native_frame':actual['target_native_frame'],'expected_model':request['expected_model'],'datasheet_storage_kind':sheet['storage_kind'],'generated_datasheet_orientation':actual['datasheet_orientation'],'sequence_cells_compared_in_canonical_orientation':actual['target_family']=='CSequenceChartSE','json_layout_contract':contract,'read_only':True}

def run(input_path,expected_sha256,output_path,request,slide_id=None,slide_number=None,shape_id=None,shape_tag=None,execute=False,ppttc_path=None):
 src=Path(input_path).resolve();out=Path(output_path).resolve();stage=out.parent/(out.stem+'_thinkcell_work')
 need(src!=out and out.suffix.lower()=='.pptx','Output must be a distinct PPTX')
 need(not out.exists(),'Output exists; refusing overwrite')
 with zipfile.ZipFile(src) as z:need(len(logical_slides(z))==1,'Runner requires a single-slide source; clone selected slide natively first')
 prep=prepare(src,expected_sha256,stage/'prepared.pptx',slide_id,slide_number,shape_id,shape_tag,execute=False)
 validate_request(request,prep['target']['model_type'])
 contract=canonical_sequence_contract(src,request,target=prep['target'])
 report={'status':'DRY_RUN_PASS','source':str(src),'source_sha256':expected_sha256.upper(),'output':str(out),'preparation':prep,'data_request':request,'json_layout_contract':contract,'native_certification_required':True}
 if not execute:return report
 ppttc=Path(ppttc_path) if ppttc_path else find_ppttc();need(ppttc is not None and ppttc.is_file(),'Official ppttc executable not found; run doctor or set THINKCELL_PPTTC');need(not stage.exists(),'Staging directory exists; refusing reuse');need(out.parent.exists(),'Output parent missing')
 stage.mkdir();prepared=stage/'prepared.pptx';generated=stage/'generated.pptx'
 prep=prepare(src,expected_sha256,prepared,slide_id,slide_number,shape_id,shape_tag,execute=True)
 snapshot(prepared,prep['automation_name'])  # Reject unsupported/link-bearing siblings before starting Office.
 (stage/'preparation.json').write_text(json.dumps(prep,indent=2));job=stage/'update.ppttc'
 job.write_text(json.dumps([{'template':str(prepared),'data':[{'name':prep['automation_name'],'table':payload(request['matrix'],contract)}]}],indent=2))
 startup=subprocess.STARTUPINFO();startup.dwFlags|=subprocess.STARTF_USESHOWWINDOW;startup.wShowWindow=0
 with (stage/'generator.stdout.txt').open('w') as stdout,(stage/'generator.stderr.txt').open('w') as stderr:
  process=subprocess.Popen([str(ppttc),str(job),'-o',str(generated)],stdout=stdout,stderr=stderr,startupinfo=startup,creationflags=subprocess.CREATE_NO_WINDOW)
  try:code=process.wait(timeout=180)
  except subprocess.TimeoutExpired:raise RuntimeError(f'Generator still running, PID {process.pid}; no process was killed. Inspect staging directory before retrying.')
 need(code==0,'Official JSON generator failed, exit '+str(code));need(generated.exists(),'Generator produced no presentation')
 validation=validate_output(prepared,generated,prep['automation_name'],request)
 need(sha(src.read_bytes())==expected_sha256.upper(),'Source changed during run')
 with out.open('xb') as f:f.write(generated.read_bytes())
 report.update(status='JSON_OUTPUT_VALIDATED_PENDING_NATIVE_CERTIFICATION',preparation=prep,validation=validation,output_sha256=sha(out.read_bytes()),staging_directory=str(stage),source_unchanged=True,generator_exit_code=code)
 return report

def main():
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('--input',required=True,type=Path);p.add_argument('--expected-sha256',required=True);p.add_argument('--output',required=True,type=Path);p.add_argument('--data-json',required=True,type=Path);p.add_argument('--report',required=True,type=Path)
 p.add_argument('--slide-id',type=int);p.add_argument('--slide-number',type=int);p.add_argument('--shape-id',type=int);p.add_argument('--shape-tag');p.add_argument('--execute',action='store_true');a=p.parse_args()
 try:
  need(a.report.resolve() not in {a.input.resolve(),a.output.resolve(),a.data_json.resolve()},'Report overlaps an input/output')
  need(a.report.suffix.lower()=='.json' and not a.report.resolve().is_relative_to(SKILL.resolve()),'Report must be local JSON outside installed assets')
  need(not a.report.exists(),'Report exists; use a new report path')
  request=json.loads(a.data_json.read_text(encoding='utf-8-sig'));report=run(a.input,a.expected_sha256,a.output,request,a.slide_id,a.slide_number,a.shape_id,a.shape_tag,a.execute)
  a.report.write_text(json.dumps(report,indent=2),encoding='utf-8');print(json.dumps({'status':report['status'],'report':str(a.report.resolve())}))
 except Exception as e:print(json.dumps({'status':'REJECTED','error':str(e)}),file=sys.stderr);sys.exit(1)
if __name__=='__main__':main()
