"""Actual converted native fields, with stale-text and binding negative controls."""
import copy
from pathlib import Path
import sys
import unittest
from unittest.mock import patch

S=Path(__file__).resolve().parents[1]/'skills/thinkcell-edit/scripts'
sys.path.insert(0,str(S))
import thinkcell
from chart_geometry import inventory
from percent_labels import native_relative_contract as contract
from percent_labels.discover_percent_semantics import discover,select
LAB=Path(__file__).resolve().parents[3]/'research/native-percent-labels'

def chart(path):
    return inventory(path.read_bytes())[1][0]


class NativeRelativeContractTests(unittest.TestCase):
    def test_real_field_backed_fixture_is_selectable(self):
        p=LAB/'conversion/00-one-native-percent.pptx'
        row=select(discover(p),category=2024.0,series='Series 3')
        self.assertEqual(row['physical_shapes'][0]['fields'][0]['text'],'63%')
        self.assertIsNone(row['absolute_text_variable'])
        self.assertEqual(contract.geometry(LAB/'00-native-absolute.pptx'),contract.geometry(p))

    def test_all_nine_remain_dynamic_across_two_actual_updates(self):
        original=contract.geometry(LAB/'00-native-absolute.pptx')
        converted=contract.geometry(LAB/'conversion/01-all-native-percent.pptx')
        self.assertEqual({k:v for k,v in original.items() if k!='bounds_emu'},
                         {k:v for k,v in converted.items() if k!='bounds_emu'})
        before=None
        for name,expected in [('01-all-native-percent.pptx','63%'),('02-numerator.pptx','74%'),('03-denominator.pptx','57%')]:
            p=LAB/'conversion'/name
            result=contract.snapshot(p,chart(p))
            self.assertEqual(len(result['labels']),9)
            target=next(r for r in result['labels'] if r['category_index']==0 and r['series_index']==2)
            self.assertEqual(target['display'],expected)
            if before:
                self.assertEqual(contract.verify(before,p,chart(p))['status'],'native_relative_fields_preserved')
            before=result

    def test_stale_text_is_rejected(self):
        p=LAB/'conversion/01-all-native-percent.pptx';rows=discover(p)
        row=rows[0];row['physical_shapes'][0]['fields'][0]['text']='999%'
        row['physical_shapes'][0]['literal_texts']=['999%']
        with patch.object(contract,'discover',return_value=rows),self.assertRaisesRegex(ValueError,'current native ratio'):
            contract.snapshot(p,chart(p))

    def test_disappearing_field_and_source_identity_change_are_rejected(self):
        p=LAB/'conversion/01-all-native-percent.pptx';before=contract.snapshot(p,chart(p))
        for change in ('missing','identity'):
            damaged=copy.deepcopy(before)
            if change=='missing':damaged['labels'].pop()
            else:damaged['labels'][0]['source_guid']='different-source'
            with patch.object(contract,'snapshot',return_value=damaged),self.assertRaisesRegex(ValueError,'lost identity'):
                contract.verify(before,p,chart(p))


if __name__=='__main__':unittest.main()
