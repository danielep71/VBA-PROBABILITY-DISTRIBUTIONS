"""Source isolation guards for the exact-value cache; Excel tests runtime behavior."""
import hashlib
import json
import unittest
import cache_vba as C


class CacheSourceTests(unittest.TestCase):
    def test_exact_reference_and_round_trip(self):
        source = C.SRC.read_text()
        original = C.strip(source)
        self.assertEqual(C.apply(original), source)
        prior = json.loads((C.ROOT / 'benchmark/harness_timing/run_43017348872_profile.json').read_text())
        self.assertEqual(hashlib.sha256(original.encode()).hexdigest(), prior['instrumentation']['source_sha256_lf'])

    def test_wrapper_drift_and_double_install_are_refused(self):
        source = C.SRC.read_text()
        with self.assertRaises(ValueError):
            C.strip(source.replace('If Key = PreviousKey Then', 'If True Then', 1))
        with self.assertRaises(ValueError):
            C.apply(source)

    def test_stripping_does_not_hide_arithmetic_changes(self):
        source = C.SRC.read_text()
        changed = source.replace('For K = 1 To N\n', 'For K = 2 To N\n', 1)
        self.assertNotEqual(changed, source)
        self.assertNotEqual(C.strip(changed), C.strip(source))


if __name__ == '__main__':
    unittest.main()
