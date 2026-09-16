"""Integrity scope for a slide containing fixed-topology waterfall plus ordinary sequences."""
from pathlib import Path
import sys,zipfile
S=Path(__file__).resolve().parent;sys.path[:0]=[str(S),str(S/'thinkcell_no_click'/'implementation')]
from prepare_thinkcell_name import inventory,need
from update_thinkcell_json import model_of
from chart_semantics import kind
from audit_thinkcell_integrity import chart_details,compare_sequence_model_to_chart

def mixed_integrity_scope(path,audit):
 """Reject package errors; validate ordinary charts by their exact native-chart parts.
 Waterfall cache parity is excluded only for that owner; its data/topology is checked by caller.
 """
 bad=[k for k,v in audit['assertions'].items() if not v and k not in {'no_model_visible_chart_mismatch','strict_parity_available','strict_parity_pass'}]
 need(not bad,'Integrity checks failed: '+', '.join(bad))
 _,charts,_=inventory(Path(path).read_bytes());water=[c for c in charts if kind(c)=='waterfall'];ordinary=[c for c in charts if kind(c)=='CSequenceChartSE']
 need(len(water)==1 and len(ordinary)+1==len(charts),'mixed integrity needs one waterfall and ordinary sequence siblings')
 with zipfile.ZipFile(path) as z:
  checks=[]
  for c in ordinary:
   part=c['frames'][0]['native_chart_part'];need(part in z.namelist(),'ordinary native cache missing')
   comparison=compare_sequence_model_to_chart(model_of(c),chart_details(z.read(part)))
   need(comparison['supported'] and comparison['values_match'] and comparison['names_match'],'ordinary model/cache mismatch: '+c['owner_name'])
   checks.append({'name':c['owner_name'],'chart_part':part,'parity':'pass'})
 return {'native_cache_parity':'mixed:ordinary-exact;waterfall-excluded','ordinary_checks':checks,'waterfall_exclusion':{'name':water[0]['owner_name'],'reason':'waterfall cache offset representation; caller must validate exact datasource and equals/connectors/grounds'},'excluded_inapplicable_checks':['waterfall native cache parity only']}
