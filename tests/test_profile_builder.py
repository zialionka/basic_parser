import os
import tempfile
import unittest

import profile_builder as pb


class ProfileBuilderTests(unittest.TestCase):
    def test_build_profile_from_resume_txt(self):
        with tempfile.TemporaryDirectory() as td:
            resume = os.path.join(td, 'resume.txt')
            with open(resume, 'w', encoding='utf-8') as f:
                f.write('Senior Python QA Engineer with Playwright, API testing, CI/CD and Selenium.')

            profile = pb.build_profile_from_resume(resume)
            self.assertIn('python', profile['stack_keywords'])
            self.assertIn('qa engineer', profile['role_keywords'])
            self.assertEqual(profile['seniority'], 'senior')
            self.assertTrue(profile['queries'])

    def test_build_profile_file(self):
        with tempfile.TemporaryDirectory() as td:
            resume_dir = os.path.join(td, 'resume')
            out_dir = os.path.join(td, 'results')
            os.makedirs(resume_dir, exist_ok=True)
            resume = os.path.join(resume_dir, 'cv.txt')
            out = os.path.join(out_dir, 'profile.yaml')

            with open(resume, 'w', encoding='utf-8') as f:
                f.write('Data Engineer with SQL, Python, AWS and Docker. Senior level.')

            profile = pb.build_profile(resume_dir, out)
            self.assertIsNotNone(profile)
            self.assertTrue(os.path.exists(out))


if __name__ == '__main__':
    unittest.main()
