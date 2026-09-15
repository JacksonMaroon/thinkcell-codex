"""Regression checks for native data modes, without proprietary fixtures."""
import io
from pathlib import Path
import struct
import sys
import unittest

SCRIPTS=Path(__file__).resolve().parents[1]/'skills/thinkcell-edit/scripts'
sys.path[:0]=[str(SCRIPTS),str(SCRIPTS/'thinkcell_no_click/implementation')]
import olefile
from extract_thinkcell_named_datasheet import wrap_biff_workbook_stream
from audit_thinkcell_integrity import normalized_percent_series,compare_pie_model_to_chart
from update_thinkcell_json import payload


class DataModes(unittest.TestCase):
    def test_short_biff_uses_valid_regular_stream(self):
        raw=struct.pack('<HHHH',0x0809,4,0x0600,5)+b'\0'*100
        with olefile.OleFileIO(io.BytesIO(wrap_biff_workbook_stream(raw))) as ole:
            extracted=ole.openstream('Workbook').read()
        self.assertEqual(extracted[:len(raw)],raw)
        self.assertEqual(len(extracted),4096)

    def test_explicit_denominator_is_not_silently_renormalized(self):
        self.assertEqual(normalized_percent_series([[20],[30]],[100]),[[20],[30]])
        self.assertEqual(normalized_percent_series([[20],[30]]),[[40],[60]])

    def test_pie_fraction_payload_uses_percentage_points(self):
        result=payload([[None,None],['A',.3],['B',.7]],{'family':'pie','input_mode':'percentage'})
        self.assertEqual(result[1],[{'string':'A'},{'percentage':30}])
        self.assertEqual(result[2],[{'string':'B'},{'percentage':70}])

    def test_percentage_pie_cache_must_use_correct_scale(self):
        model={'series_count':1,'values':[.3,.7],'categories':['A','B'],
               'calculated_total':1,'total':100,'input_mode':'percentage'}
        chart={'types':['doughnutChart'],'series':[{'values':[30,70],'categories':['A','B']}]}
        self.assertTrue(compare_pie_model_to_chart(model,chart)['values_match'])
        chart['series'][0]['values']=[.3,.7]
        self.assertFalse(compare_pie_model_to_chart(model,chart)['values_match'])


if __name__=='__main__':unittest.main()
