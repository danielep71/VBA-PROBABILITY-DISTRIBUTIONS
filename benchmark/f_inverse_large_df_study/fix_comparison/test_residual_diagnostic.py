"""Check isolation and the diagnostic's exact-vs-rounded argument distinction."""
import unittest
import diagnose_residual as diagnostic
import vba_order_model as model


class DiagnosticTests(unittest.TestCase):
    def test_substitutions_restore_model(self):
        original=model.forward.beta_regularized
        cf=model.forward.beta_continued_fraction
        for mode in ['forward','prefactor_supplied','cf','cf_exact_complement','joint_prefactor_cf']:
            diagnostic.run_one(.9,100.,1e10,mode,50)
            self.assertIs(model.forward.beta_regularized,original)
            self.assertIs(model.forward.beta_continued_fraction,cf)

    def test_complement_matters_even_with_high_precision_cf(self):
        rounded=diagnostic.run_one(.9,100.,1e10,'cf',80)
        exact=diagnostic.run_one(.9,100.,1e10,'cf_exact_complement',80)
        self.assertNotEqual(rounded[0],exact[0])
        self.assertLess(exact[1],.5)
        self.assertTrue(exact[3][-1]['settled'])


if __name__=='__main__':
    unittest.main()
