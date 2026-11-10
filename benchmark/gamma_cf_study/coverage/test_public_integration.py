"""Source-level public wiring, exact private-port extraction and historical isolation."""
import re
import unittest
from install_scaled_fidelity import documented
from public_integration import b64,ROOT,PORT,SRC,TEST,BEGIN,END,EXTRA,RERAISE,RERAISE_LINES,AMENDMENTS,amend,code_only,original,procedures,arithmetic,adapter,test_block,restore,live_total


class PublicIntegrationTests(unittest.TestCase):
    maxDiff=1600
    def test_only_intended_public_source_changes(self):
        current=(ROOT/SRC).read_text()
        self.assertEqual(restore(current),original(SRC))
        block,types,names=arithmetic()
        self.assertEqual(len(names),46)
        self.assertEqual(current.count(block),1);self.assertEqual(current.count(types),1)
        self.assertNotIn('AssertTrue',block)
        source=procedures(current)
        for tail in ('P','Q'):
            self.assertEqual(source['PROB_TryGammaRegularized'+tail],adapter(tail))
        # No unreviewed public/API/module expansion: only the two internal
        # binary64 routines of #307 (b64_route_vba.py), in the private layer.
        self.assertEqual(re.findall(r'^Public (?:Function|Sub|Const) (\w+)',current,re.M),
                         re.findall(r'^Public (?:Function|Sub|Const) (\w+)',original(SRC),re.M)+list(b64.PUBLIC))
        for path in (ROOT/'src').glob('*.bas'):
            rel=path.relative_to(ROOT).as_posix()
            if rel!=SRC:
                self.assertEqual(amend(rel,path.read_text(),reverse=True),original(rel))

    def test_old_regressions_and_rejected_controls_preserved(self):
        text=(ROOT/TEST).read_text()
        self.assertEqual(restore(text),original(TEST))
        self.assertEqual(text.count(test_block()),1)
        self.assertEqual(text.count('    Test_PublicGamma_Integration\n'),1)
        procs=procedures(text)
        for name in ('EG_RoutingControl','WG_RoutingControl'):
            self.assertIn('PG_B6_',procs[name])
            self.assertNotIn('PROB_TryGammaRegularized',procs[name])
        for tail in ('P','Q'):
            old=procedures(original(SRC))['PROB_TryGammaRegularized'+tail]
            self.assertIn(documented(old.replace('Public Function','Private Function',1).replace('PROB_TryGammaRegularized'+tail,'PG_B6_'+tail)),test_block())
        self.assertEqual(live_total(),10848+EXTRA)
        self.assertEqual(EXTRA,290*8+7*2+206*8+24*2+2*5)

    def test_review_fixes_in_production(self):
        """#242 review: no comment-reached dead code, unexpected runtime errors
        propagate (#VALUE!), and the binary64 route answers at the pair's cap."""
        current=(ROOT/SRC).read_text();source=procedures(current)
        for dead in ('PROB_DD_PV2_Pair','PROB_DD_PV2_LogGamma','PROB_DD_PV2_Log(',
                     'PROB_DD_PV2_Log1p','PROB_DD_PV2_Expm1','PROB_DD_PV2_Exp(','PROB_DD_PV2_Ln2'):
            self.assertNotIn(dead,current)
        self.assertNotIn('test-only arithmetic port',current)
        self.assertNotIn('not a public kernel',current)
        for name in RERAISE:
            self.assertIn('Err_Handler:\n'+RERAISE_LINES,source['PROB_DD_'+name])
        for tail in ('P','Q'):
            body=source['PROB_TryGammaRegularized'+tail]
            self.assertIn('If Not PROB_TryGammaSeriesP(A, X, Value, FailMsg) Then Exit Function',body)
            self.assertIn('If Not PROB_TryGammaContinuedFractionQ(A, X, Value, FailMsg) Then Exit Function',body)
            self.assertNotIn('acceptance pending',body)
        tests=procedures((ROOT/TEST).read_text())
        self.assertIn('    Test_PublicGamma_Reach\n',tests['Test_PublicGamma_Integration'])
        self.assertEqual(tests['Test_PublicGamma_Reach'].count('PG_Reach "PG reach'),2)

    def test_amendments_change_comments_only(self):
        """The amendments outside the generated block touch comments, never code."""
        self.assertEqual(set(AMENDMENTS),{SRC,'src/M_STATS_PROBDIST_TFAMILY.bas',
                                          'src/M_STATS_PROBDIST_CONTINUOUS.bas'})
        for path,pairs in AMENDMENTS.items():
            current=(ROOT/path).read_text()
            for old,new in pairs:
                self.assertEqual(code_only(old).rstrip(),code_only(new).rstrip(),(path,old[:60]))
                self.assertEqual(current.count(new),1,(path,new[:60]))
                self.assertNotEqual(old,new)
        with self.assertRaises(ValueError):
            amend('src/M_STATS_PROBDIST_TFAMILY.bas','no amendment here')

    def test_reversal_refuses_arithmetic_or_adapter_mutations(self):
        source=(ROOT/SRC).read_text()
        mutations=[source.replace('Result = P\n','Result = Q\n',1),
                   source.replace('Private Function PROB_DD_WG_Pair(', 'Private Function PROB_DD_WG_PairBroken(',1),
                   source.replace('Lower Q is reconstructed directly','Lower Q is changed',1)]
        for mutant in mutations:
            self.assertNotEqual(mutant,source)
            with self.assertRaises(ValueError):restore(mutant)
        text=(ROOT/TEST).read_text()
        with self.assertRaises(ValueError):restore(text.replace('PG invalid P cleared','PG invalid P ignored',1))

    def test_historical_strip_does_not_remove_unrelated_drift(self):
        source=(ROOT/SRC).read_text().replace('Public Function PROB_LogGamma(', 'Public Function PROB_LogGammaChanged(',1)
        self.assertNotEqual(restore(source),original(SRC))
        text=(ROOT/TEST).read_text().replace('Option Explicit','Option Explicit changed',1)
        self.assertTrue(restore(text)!=original(TEST))
        self.assertEqual(PORT,'b8efaf3c7810c48fadbfece38950ded190f79a56')
        self.assertIn(BEGIN,test_block());self.assertIn(END,test_block())


if __name__=='__main__':unittest.main()
