"""Independent nonlinear regression and native forecast contract tests, synthetic data only."""
from pathlib import Path
import math,sys,unittest
from lxml import etree as E
SCRIPTS=Path(__file__).resolve().parents[1]/'skills/thinkcell-edit/scripts';sys.path.insert(0,str(SCRIPTS))
import native_insertions as a
class NonlinearProfiles(unittest.TestCase):
 def data(self,kind):
  x=[1.,2.,4.];y=[3*math.exp(.5*v) for v in x] if kind=='exp' else [2+3*math.log(v) for v in x]
  return {'group_labels':['A']*3+['B'],'x_values':x+[99],'y_values':y+[999]}
 def axes(self,ymax):return E.fromstring(f'<chart><valAx><axPos val="b"/><scaling><min val="0"/><max val="9"/></scaling></valAx><valAx><axPos val="l"/><scaling><max val="{ymax}"/></scaling></valAx></chart>')
 def test_exponential_known_fit_and_sibling_exclusion(self):
  c,s=a.nonlinear_fit(self.data('exp'),'A','exp');self.assertAlmostEqual(c,3);self.assertAlmostEqual(s,.5)
 def test_logarithmic_known_fit(self):
  c,s=a.nonlinear_fit(self.data('log'),'A','log');self.assertAlmostEqual(c,2);self.assertAlmostEqual(s,3)
 def test_exponential_changed_data(self):
  data=self.data('exp');data['y_values']=[v*2 for v in data['y_values']];c,s=a.nonlinear_fit(data,'A','exp');self.assertAlmostEqual(c,6);self.assertAlmostEqual(s,.5)
 def test_logarithmic_changed_data(self):
  data=self.data('log');data['y_values']=[v*2 for v in data['y_values']];c,s=a.nonlinear_fit(data,'A','log');self.assertAlmostEqual(c,4);self.assertAlmostEqual(s,6)
 def test_nonpositive_log_x_and_exp_y_rejected(self):
  for kind,field in [('exp','y_values'),('log','x_values')]:
   data=self.data(kind);data[field][0]=0
   with self.assertRaisesRegex(ValueError,'positive'):a.nonlinear_fit(data,'A',kind)
 def test_nonfinite_and_rank_deficient_rejected(self):
  data=self.data('exp');data['y_values'][0]=float('nan')
  with self.assertRaisesRegex(ValueError,'finite'):a.nonlinear_fit(data,'A','exp')
  data=self.data('exp');data['x_values'][:3]=[1,1,1]
  with self.assertRaisesRegex(ValueError,'distinct X'):a.nonlinear_fit(data,'A','exp')
 def test_decreasing_curve_outside_verified_profile(self):
  data=self.data('exp');data['y_values'][:3]=[9,5,1]
  with self.assertRaisesRegex(ValueError,'increasing'):a.nonlinear_fit(data,'A','exp')
 def test_unverified_polynomial_mapping_not_accepted(self):
  self.assertEqual(a.TRENDLINE_PROFILES['quadratic-trendline'], ('6','poly'))
  self.assertNotIn(('4','poly'),a.TRENDLINE_PROFILES.values())
  with self.assertRaisesRegex(ValueError,'Unverified'):a.nonlinear_fit(self.data('exp'),'A','poly')
 def test_both_observed_clipping_endpoints_and_stale_cache_rejected(self):
  for kind,c,s,ymax in [('exp',3,.5,3*math.exp(2.5)),('log',2,3,2+3*math.log(5))]:
   trend=E.fromstring(b'<trendline><forward val="1"/></trendline>');result=a.verify_nonlinear_forecast(self.axes(ymax),trend,4,c,s,kind);self.assertAlmostEqual(result['reference_upper_x'],5);self.assertTrue(result['fit_limits_visible_endpoint']);trend[0].set('val','5')
   with self.assertRaisesRegex(ValueError,'contradicts'):a.verify_nonlinear_forecast(self.axes(ymax),trend,4,c,s,kind)
 def test_duplicate_forecast_and_logarithmic_axes_rejected(self):
  chart=self.axes(10);trend=E.fromstring(b'<trendline><forward val="1"/><forward val="2"/></trendline>')
  with self.assertRaisesRegex(ValueError,'ambiguous'):a.verify_nonlinear_forecast(chart,trend,4,3,.5,'exp')
  E.SubElement(chart.find('valAx/scaling'),'logBase',val='10')
  with self.assertRaisesRegex(ValueError,'linear X/Y'):a.verify_nonlinear_forecast(chart,E.fromstring(b'<trendline/>'),4,3,.5,'exp')

class PolynomialProfiles(unittest.TestCase):
 def data(self,order):
  x=[-2.,-1.,0.,1.,2.,3.];coeff=[2,-1,.5,-.25,.1][:order+1]
  return {'group_labels':['A']*len(x),'x_values':x,'y_values':[sum(c*v**j for j,c in enumerate(coeff)) for v in x]},coeff
 def test_each_authenticated_degree_recovers_known_polynomial(self):
  for order in [2,3,4]:
   data,wanted=self.data(order);actual=a.polynomial_fit(data,'A',order)
   for x,y in zip(actual,wanted):self.assertAlmostEqual(x,y,places=9)
 def test_changed_shape_recovered_and_unrelated_group_excluded(self):
  data,wanted=self.data(4);data['y_values']=[y+.5*x for x,y in zip(data['x_values'],data['y_values'])];wanted[1]+=.5
  data['group_labels'].append('B');data['x_values'].append(99);data['y_values'].append(999)
  for x,y in zip(a.polynomial_fit(data,'A',4),wanted):self.assertAlmostEqual(x,y,places=9)
 def test_rank_deficient_and_nonfinite_poly_rejected(self):
  data,_=self.data(4);data['x_values']=[1]*len(data['x_values'])
  with self.assertRaisesRegex(ValueError,'rank'):a.polynomial_fit(data,'A',4)
  data,_=self.data(2);data['y_values'][0]=float('inf')
  with self.assertRaisesRegex(ValueError,'finite'):a.polynomial_fit(data,'A',2)
 def test_repeated_root_and_all_curve_intersections(self):
  from trendline_math import roots_in_domain
  self.assertEqual(roots_in_domain([4,-4,1],0,5),[2.0])
  roots=roots_in_domain([-6,11,-6,1],0,5)
  for x,y in zip(roots,[1,2,3]):self.assertAlmostEqual(x,y)
 def test_quadratic_fit_dependent_upper_boundary_and_stale_forecast(self):
  from trendline_math import verify_polynomial_forecast
  chart=E.fromstring(b'<chart><valAx><axPos val="b"/><scaling><min val="0"/><max val="9"/></scaling></valAx><valAx><axPos val="l"/><scaling><min val="0"/><max val="8"/></scaling></valAx></chart>')
  trend=E.fromstring(b'<trendline><forward val="2"/></trendline>')
  result=verify_polynomial_forecast(chart,trend,3,[0,5,-1]);self.assertAlmostEqual(result['reference_upper_x'],5);self.assertTrue(result['fit_limits_visible_endpoint']);trend[0].set('val','6')
  with self.assertRaisesRegex(ValueError,'contradicts'):verify_polynomial_forecast(chart,trend,3,[0,5,-1])
 def test_wrong_polynomial_order_not_registered(self):
  self.assertEqual(a.POLYNOMIAL_ORDERS,{'quadratic-trendline':2,'cubic-trendline':3,'quartic-trendline':4})
  with self.assertRaisesRegex(ValueError,'Unverified'):a.polynomial_fit(self.data(2)[0],'A',5)

class NonlinearDomainRegression(unittest.TestCase):
 def axes(self,xmax,ymax,ymin=0):
  return E.fromstring(f'<chart><valAx><axPos val="b"/><scaling><min val="0"/><max val="{xmax}"/></scaling></valAx><valAx><axPos val="l"/><scaling><min val="{ymin}"/><max val="{ymax}"/></scaling></valAx></chart>')
 def test_all_nonlinear_helpers_reject_backward_domain_tampering(self):
  calls=[lambda t:a.verify_power_forecast(self.axes(9,48),t,4,3,2,xmin=1),
         lambda t:a.verify_nonlinear_forecast(self.axes(9,3*math.exp(2.5)),t,4,3,.5,'exp',xmin=1),
         lambda t:a.verify_nonlinear_forecast(self.axes(9,2+3*math.log(5)),t,4,2,3,'log',xmin=1),
         lambda t:a.verify_polynomial_forecast(self.axes(9,8),t,3,[0,5,-1],xmin=1)]
  for call in calls:
   for xml in [b'<trendline><backward val="1"/></trendline>',b'<trendline><backward val="-1"/></trendline>',b'<trendline><backward val="NaN"/></trendline>',b'<trendline><backward val="0"/><backward val="0"/></trendline>',b'<trendline><backward/></trendline>']:
    with self.assertRaisesRegex(ValueError,'backward|Lower'):call(E.fromstring(xml))
 def test_missing_forward_value_rejected_even_when_data_max_matches_clip(self):
  trend=E.fromstring(b'<trendline><forward/></trendline>')
  with self.assertRaisesRegex(ValueError,'missing'):a.verify_power_forecast(self.axes(9,48),trend,4,3,2)
  with self.assertRaisesRegex(ValueError,'missing'):a.verify_polynomial_forecast(self.axes(9,8),trend,5,[0,5,-1])
 def test_power_unchanged_lower_domain_is_reported(self):
  r=a.verify_power_forecast(self.axes(9,48),E.fromstring(b'<trendline><forward val="0"/></trendline>'),4,3,2,xmin=1)
  self.assertEqual(r['actual_lower_x'],1);self.assertEqual(r['reference_lower_x'],1)
 def test_tiny_y_scaled_equivalence_rejects_false_full_extent(self):
  for scale in [1,1e-12,1e12]:
   chart=self.axes(2,scale);coeff=[0,0,scale];t=E.fromstring(b'<trendline><forward val="0"/></trendline>')
   self.assertAlmostEqual(a.verify_polynomial_forecast(chart,t,1,coeff)['reference_upper_x'],1)
   t[0].set('val','1')
   with self.assertRaisesRegex(ValueError,'contradicts'):a.verify_polynomial_forecast(chart,t,1,coeff)
 def test_tiny_x_polynomial_domain_retains_roots_and_rejects_wrong_endpoint(self):
  from trendline_math import roots_in_domain
  root=roots_in_domain([-1,0,1e26],0,2e-13)[0];self.assertAlmostEqual(root/1e-13,1)
  chart=self.axes(2e-13,1);t=E.fromstring(b'<trendline><forward val="0"/></trendline>')
  self.assertAlmostEqual(a.verify_polynomial_forecast(chart,t,1e-13,[0,0,1e26])['reference_upper_x']/1e-13,1)
  t[0].set('val','1e-13')
  with self.assertRaisesRegex(ValueError,'contradicts'):a.verify_polynomial_forecast(chart,t,1e-13,[0,0,1e26])
 def test_tiny_x_power_and_exponential_endpoint_guards(self):
  calls=[lambda t:a.verify_power_forecast(self.axes(2e-13,1e-26),t,1e-13,1,2),
         lambda t:a.verify_nonlinear_forecast(self.axes(2e-13,math.e),t,1e-13,1,1e13,'exp')]
  for call in calls:
   self.assertAlmostEqual(call(E.fromstring(b'<trendline><forward val="0"/></trendline>'))['reference_upper_x']/1e-13,1)
   with self.assertRaisesRegex(ValueError,'contradicts'):call(E.fromstring(b'<trendline><forward val="1e-13"/></trendline>'))
 def test_extreme_coordinate_transform_underflow_fails_closed(self):
  chart=self.axes(2e-200,1e-100)
  with self.assertRaisesRegex(ValueError,'underflow|unsupported'):
   a.verify_polynomial_forecast(chart,E.fromstring(b'<trendline/>'),2e-200,[0,0,1e300])
  from trendline_math import roots_in_domain
  with self.assertRaisesRegex(ValueError,'underflow|unsupported'):
   roots_in_domain([0,0,1],0,2e-200)
 def test_power_nonfinite_axis_and_fit_fail_closed(self):
  trend=E.fromstring(b'<trendline><forward val="5"/></trendline>')
  with self.assertRaisesRegex(ValueError,'finite'):a.verify_power_forecast(self.axes(9,float('inf')),trend,4,3,2)
  with self.assertRaisesRegex(ValueError,'finite'):a.verify_power_forecast(self.axes(9,48),trend,4,3,float('inf'))
 def test_zero_width_tangent_has_no_visible_curve(self):
  with self.assertRaisesRegex(ValueError,'visible interval'):
   a.verify_polynomial_forecast(self.axes(2,0,-1),E.fromstring(b'<trendline/>'),1,[1,-2,1])
if __name__=='__main__':unittest.main()
