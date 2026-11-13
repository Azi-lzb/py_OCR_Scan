"""Fast orientation must retain the complete probe for uncertain inputs."""
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'core/src'))
from scan2excel.service import Scan2ExcelService


def quality(q, mean=0.95):
    return dict(q=q, tall_ratio=0.0, mean_score=mean, n_items=12)


class OrientationProbeTests(unittest.TestCase):
    def setUp(self):
        self.service = Scan2ExcelService()
        self.image = np.arange(24).reshape(2, 4, 3)

    def test_clear_upright_uses_sample_only(self):
        with patch.object(self.service, '_probe_quality',
                          side_effect=[quality(100), quality(30)]) as probe:
            result = self.service._fix_page_rotation(self.image, None)
        self.assertIs(result, self.image)
        self.assertEqual(probe.call_count, 2)
        self.assertTrue(all(c.kwargs.get('sampled') for c in probe.call_args_list))

    def test_clear_upside_down_is_rotated(self):
        with patch.object(self.service, '_probe_quality',
                          side_effect=[quality(40), quality(100)]):
            result = self.service._fix_page_rotation(self.image, None)
        np.testing.assert_array_equal(result, np.rot90(self.image, 2))

    def test_ambiguous_samples_use_complete_probe(self):
        with patch.object(self.service, '_probe_quality', side_effect=[
                quality(100), quality(90), quality(100), quality(30)]) as probe:
            result = self.service._fix_page_rotation(self.image, None)
        self.assertIs(result, self.image)
        self.assertEqual(probe.call_count, 4)
        self.assertFalse(probe.call_args_list[-1].kwargs.get('sampled', False))

    def test_low_confidence_samples_use_complete_probe(self):
        with patch.object(self.service, '_probe_quality', side_effect=[
                quality(100, 0.7), quality(30), quality(40), quality(100)]):
            result = self.service._fix_page_rotation(self.image, None)
        np.testing.assert_array_equal(result, np.rot90(self.image, 2))

    def test_sample_failure_uses_complete_probe(self):
        with patch.object(self.service, '_probe_quality', side_effect=[
                RuntimeError('sample unavailable'), quality(100), quality(30)]):
            result = self.service._fix_page_rotation(self.image, None)
        self.assertIs(result, self.image)


if __name__ == '__main__':
    unittest.main()
