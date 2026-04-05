import unittest

import topup_jobs_qa as app


class JobScriptTests(unittest.TestCase):
    def test_norm_url(self):
        self.assertEqual(
            app.norm('https://www.Example.com/path/to/job///?a=1'),
            'https://example.com/path/to/job',
        )

    def test_parse_salary_range(self):
        self.assertEqual(app.parse_salary_range('$120,000 - $150,000'), ('$120,000', '$150,000'))
        self.assertEqual(app.parse_salary_range('$95,000 yearly'), ('$95,000', '$95,000'))
        self.assertEqual(app.parse_salary_range('not listed'), ('Not specified', 'Not specified'))

    def test_looks_like_us_remote(self):
        self.assertTrue(app.looks_like_us_remote('Senior QA Engineer remote United States', ''))
        self.assertFalse(app.looks_like_us_remote('Senior QA Engineer hybrid Berlin', 'Berlin'))

    def test_dedupe_candidates(self):
        existing_url = {'https://example.com/jobs/1'}
        existing_tc = {('senior qa engineer', 'acme')}
        cands = [
            {'Job Title': 'Senior QA Engineer', 'Company': 'Acme', 'Job URL': 'https://example.com/jobs/2', '_score': 5},
            {'Job Title': 'SDET', 'Company': 'Beta', 'Job URL': 'https://example.com/jobs/1', '_score': 8},
            {'Job Title': 'QA Automation Engineer', 'Company': 'Gamma', 'Job URL': 'https://example.com/jobs/3', '_score': 10},
        ]
        unique, stats = app.dedupe_candidates(cands, existing_url, existing_tc)
        self.assertEqual(len(unique), 1)
        self.assertEqual(unique[0]['Company'], 'Gamma')
        self.assertEqual(stats['kept'], 1)


if __name__ == '__main__':
    unittest.main()
