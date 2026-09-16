"""Offline behavioral and rejection checks for the bounded percentage route."""
import copy, unittest
from pathlib import Path
from percent_conversion_guards import ratio, dual, check, saved_ratio
from trace_percent_semantics import trace

ROOT=Path(__file__).resolve().parents[1]
FIXTURE=ROOT/'assets/feature-fixtures/percent-dual-38764.pptx'

class Guards(unittest.TestCase):
    def setUp(self):
        self.state=trace(FIXTURE,'ppt/embeddings/oleObject13.bin','191','tbUpiE_yCia4NQkLBVWFCIA')

    def test_actual_saved_ratio_and_half_up(self):
        self.assertEqual(saved_ratio(FIXTURE,self.state)[2],'71%')
        matrix=[[None,'Gross Profit','Revenues','Customer Count'],[None,200,200,200],['Not Mapped',50,20,20],['High risk',70,30,30],['Medium Risk',55,40,40],['Low risk',25,110,110]]
        self.assertEqual(ratio(matrix)[2],'13%')
        matrix[1][1]=0
        with self.assertRaises(ValueError):ratio(matrix)

    def test_stale_native_scalar_rejected(self):
        state=copy.deepcopy(self.state)
        state['sources']['m_varsrcAbsolute']['xml']=state['sources']['m_varsrcAbsolute']['xml'].replace('7.10000000000000000000E+01','2.50000000000000000000E+01')
        self.assertNotEqual(state,self.state)
        with self.assertRaisesRegex(ValueError,'numerator'):saved_ratio(FIXTURE,state)

    def test_genuine_field_required(self):
        self.assertEqual(dual(self.state),('71','71%'))
        self.state['visible_shapes'][0]['fields'][1]['type']='datetimeUNBOUND'
        with self.assertRaisesRegex(ValueError,'binding'):dual(self.state)

if __name__=='__main__':unittest.main()
