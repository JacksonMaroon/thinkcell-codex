"""Independent bounded polynomial regression and plot-boundary readback."""
import math
import sys


def need(ok,message):
 if not ok:raise ValueError(message)

def evaluate(coefficients,x):return math.fsum(c*x**j for j,c in enumerate(coefficients))

def polynomial_fit(data,group,order):
 need(order in {2,3,4},'Unverified polynomial order')
 pairs=[(x,y) for g,x,y in zip(data['group_labels'],data['x_values'],data['y_values']) if g==group]
 need(len(pairs)>=order+1 and all(math.isfinite(x) and math.isfinite(y) for x,y in pairs),'Polynomial fit requires finite data and enough points')
 need(len({x for x,y in pairs})>=order+1,'Polynomial fit requires full rank distinct X values')
 scale=max(abs(x) for x,y in pairs);need(scale>0,'Polynomial fit has no nonzero X extent');x=[p[0]/scale for p in pairs];y=[p[1] for p in pairs];m=order+1
 matrix=[[sum(v**(j+k) for v in x) for j in range(m)]+[sum(v**k*w for v,w in zip(x,y))] for k in range(m)]
 for k in range(m):
  piv=max(range(k,m),key=lambda j:abs(matrix[j][k]));matrix[k],matrix[piv]=matrix[piv],matrix[k];need(abs(matrix[k][k])>1e-13,'Polynomial fit is numerically rank deficient');div=matrix[k][k];matrix[k]=[v/div for v in matrix[k]]
  for j in range(m):
   if j!=k:
    mul=matrix[j][k];matrix[j]=[v-mul*w for v,w in zip(matrix[j],matrix[k])]
 coeff=[row[-1]/scale**j for j,row in enumerate(matrix)];need(all(math.isfinite(v) for v in coeff),'Polynomial fit is nonfinite');return coeff

def _unit_roots(coefficients):
    coeff = list(coefficients)
    while len(coeff) > 1 and coeff[-1] == 0:
        coeff.pop()
    scale = max(abs(v) for v in coeff)
    if not scale or len(coeff) == 1:
        return []
    coeff = [v / scale for v in coeff]
    if len(coeff) == 2:
        root = -coeff[0] / coeff[1]
        return [root] if 0 <= root <= 1 else []
    critical = _unit_roots([j * c for j, c in enumerate(coeff) if j])
    boundaries = sorted([0.0, *critical, 1.0])
    roots = []
    def zero_residual(point, value):
        error_scale = math.fsum(abs(c * point**j) for j, c in enumerate(coeff))
        return abs(value) <= 128 * math.ulp(1.0) * error_scale
    for point in boundaries:
        residual = evaluate(coeff, point)
        if zero_residual(point, residual):
            roots.append(point)
    for lower, upper in zip(boundaries, boundaries[1:]):
        left, right = evaluate(coeff, lower), evaluate(coeff, upper)
        if zero_residual(lower, left) or zero_residual(upper, right) or (left < 0) == (right < 0):
            continue
        for _ in range(65):
            middle = (lower + upper) / 2
            value = evaluate(coeff, middle)
            if value == 0:
                lower = upper = middle
                break
            if (left < 0) != (value < 0):
                upper = middle
            else:
                lower, left = middle, value
        roots.append((lower + upper) / 2)
    result = []
    for point in sorted(roots):
        if not result or abs(point - result[-1]) > 128 * math.ulp(1.0):
            result.append(point)
    return result


def roots_in_domain(coefficients, lower, upper):
    """Scale both coordinates and residuals; retain tiny-domain root positions."""
    need(math.isfinite(lower) and math.isfinite(upper) and lower < upper,
         "Polynomial root domain invalid")
    need(all(math.isfinite(c) for c in coefficients), "Polynomial coefficients nonfinite")
    span = upper - lower
    need(math.isfinite(span), "Polynomial root domain span nonfinite")
    coefficient_scale = max(abs(c) for c in coefficients)
    if not coefficient_scale:
        return []
    normalized = [c / coefficient_scale for c in coefficients]
    need(all(c == 0 or (v != 0 and abs(v) >= sys.float_info.min)
             for c, v in zip(coefficients, normalized)),
         "Polynomial coefficient normalization underflow; numerical profile unsupported")
    transformed = []
    for j in range(len(normalized)):
        terms = []
        for k, c in enumerate(normalized):
            if k < j or c == 0 or (lower == 0 and k > j):
                continue
            term = c * math.comb(k, j) * lower**(k-j) * span**j
            need(math.isfinite(term) and abs(term) >= sys.float_info.min,
                 "Polynomial domain transformation underflow or overflow; numerical profile unsupported")
            terms.append(term)
        transformed.append(math.fsum(terms))
    need(all(math.isfinite(c) for c in transformed), "Polynomial domain transformation nonfinite")
    return [lower + span * root for root in _unit_roots(transformed)]


def verify_lower_forecast(trend, xmin=None):
    """Authenticated graft controls begin at the selected series minimum X."""
    backward = trend.xpath('./*[local-name()="backward"]/@val')
    elements = trend.xpath('./*[local-name()="backward"]')
    need(len(elements) <= 1 and len(backward) == len(elements), "Lower forecast is ambiguous or missing its value")
    extension = float(backward[0]) if backward else 0.0
    need(math.isfinite(extension) and extension == 0,
         "Nonzero backward forecast is outside the authenticated native profile")
    if xmin is not None:
        need(math.isfinite(xmin), "Lower forecast X is nonfinite")
    return {"actual_lower_x": xmin, "reference_lower_x": xmin,
            "lower_domain": "selected_series_minimum_x", "backward_extension": extension}


def forecast_close(actual, wanted, lower, upper):
    """Compare on axis span, avoiding a fixed tolerance on tiny domains."""
    return (all(math.isfinite(v) for v in (actual, wanted, lower, upper))
            and lower < upper and abs(actual - wanted) <= (upper - lower) * 1e-7)


def verify_polynomial_forecast(cache, trend, xmax, coefficients, xmin=None):
    lower_domain = verify_lower_forecast(trend, xmin)
    bounds = {}
    for kind, positions in [('x', {'b', 't'}), ('y', {'l', 'r'})]:
        axes = [a for a in cache.xpath('.//*[local-name()="valAx"]')
                if a.xpath('./*[local-name()="axPos"]/@val')
                and a.xpath('./*[local-name()="axPos"]/@val')[0] in positions]
        need(len(axes) == 1 and not axes[0].xpath('./*[local-name()="scaling"]/*[local-name()="logBase"]'),
             "Polynomial forecast requires unique linear X/Y axes")
        values = [axes[0].xpath('./*[local-name()="scaling"]/*[local-name()="'+field+'"]/@val')
                  for field in ['min', 'max']]
        need(all(len(v) == 1 for v in values), "Polynomial forecast requires explicit axis bounds")
        bounds[kind] = [float(v[0]) for v in values]
        need(all(math.isfinite(v) for v in bounds[kind]) and bounds[kind][0] < bounds[kind][1],
             "Polynomial forecast axis bounds invalid")
    lower, upper = bounds['x']
    boundaries = [lower, upper]
    for edge in bounds['y']:
        shifted = list(coefficients)
        shifted[0] -= edge
        boundaries.extend(roots_in_domain(shifted, lower, upper))
    boundaries = sorted(set(boundaries))
    # A solitary tangent point has no visible interval and cannot establish a
    # curve extent. Include every genuine interval when the fit reenters the plot.
    visible = [(left, right) for left, right in zip(boundaries, boundaries[1:])
               if left < right and bounds['y'][0] <= evaluate(coefficients, (left+right)/2) <= bounds['y'][1]]
    need(visible, "Polynomial fit has no visible interval in plot")
    wanted = max(right for left, right in visible)
    forward = trend.xpath('./*[local-name()="forward"]/@val')
    need(len(forward) <= 1 and len(forward) == len(trend.xpath('./*[local-name()="forward"]')),
         "Polynomial forecast ambiguous or missing its value")
    actual = xmax + (float(forward[0]) if forward else 0)
    need(forecast_close(actual, wanted, lower, upper),
         "Physical polynomial forecast endpoint contradicts independent fitted curve and plot bounds")
    return {"actual_upper_x": actual, "reference_upper_x": wanted,
            "fit_limits_visible_endpoint": wanted < upper, **lower_domain}
