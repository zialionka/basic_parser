import os
import re
from collections import Counter

try:
    import yaml
except Exception:  # pragma: no cover
    yaml = None

try:
    from pypdf import PdfReader
except Exception:  # pragma: no cover
    PdfReader = None

try:
    import docx
except Exception:  # pragma: no cover
    docx = None


ROLE_TERMS = [
    'software engineer', 'software developer', 'backend engineer', 'frontend engineer', 'full stack engineer',
    'qa engineer', 'quality assurance', 'sdet', 'test automation engineer', 'test engineer',
    'quality engineer', 'quality assurance engineer', 'qa automation engineer', 'qa analyst',
    'software test engineer', 'software tester', 'test analyst', 'automation test engineer',
    'software development engineer in test', 'software engineer in test', 'test architect',
    'qa lead', 'qa manager', 'test lead', 'test manager', 'quality lead', 'quality manager',
    'validation engineer', 'verification engineer', 'reliability test engineer', 'performance test engineer',
    'api test engineer', 'mobile test engineer', 'security test engineer', 'release validation engineer',
    'product quality engineer', 'customer quality engineer', 'quality specialist', 'quality control',
    'software qa analyst', 'software qa engineer', 'quality assurance analyst', 'quality assurance specialist',
    'quality assurance tester', 'quality verification engineer', 'testing specialist', 'testing engineer',
    'software testing analyst', 'automated testing engineer', 'manual qa engineer', 'manual tester',
    'test automation analyst', 'test automation lead', 'integration test engineer', 'system test engineer',
    'end to end test engineer', 'e2e test engineer', 'uat analyst', 'platform quality engineer',
    'cloud qa engineer', 'backend qa engineer', 'frontend qa engineer', 'web qa engineer',
    'mobile qa engineer', 'ios qa engineer', 'android qa engineer', 'embedded test engineer',
    'firmware test engineer', 'hardware test engineer', 'build and release engineer',
    'quality operations engineer', 'test operations engineer', 'compliance test engineer',
    'accessibility test engineer', 'localization qa tester', 'data qa engineer',
    'data quality engineer', 'data quality analyst', 'etl test engineer', 'bi test engineer',
    'analytics qa engineer', 'ml test engineer', 'ai quality engineer', 'devops test engineer',
    'security qa engineer', 'staff quality engineer', 'principal quality engineer',
    'staff test engineer', 'principal test engineer', 'senior test engineer', 'sdet ii', 'sdet iii',
    'devops engineer', 'site reliability engineer', 'data engineer', 'data analyst', 'data scientist',
    'ml engineer', 'machine learning engineer', 'product manager', 'project manager', 'business analyst',
    'security engineer', 'cloud engineer', 'mobile engineer', 'ios developer', 'android developer',
    'platform engineer', 'solutions engineer', 'application engineer', 'systems engineer',
]

STACK_TERMS = [
    'python', 'java', 'javascript', 'typescript', 'go', 'golang', 'ruby', 'c#', 'c++', 'sql', 'nosql',
    'playwright', 'cypress', 'selenium', 'webdriverio', 'appium', 'robot framework', 'cucumber', 'gherkin',
    'testng', 'junit', 'postman', 'insomnia', 'soapui', 'karate', 'rest api', 'graphql', 'api testing',
    'contract testing', 'pytest', 'unittest', 'tox', 'allure',
    'jenkins', 'github actions', 'gitlab ci', 'circleci', 'azure devops', 'teamcity', 'bamboo', 'ci/cd',
    'docker', 'kubernetes', 'helm', 'aws', 'gcp', 'azure', 'terraform', 'ansible',
    'linux', 'datadog', 'grafana', 'prometheus',
    'playwright test', 'selenium webdriver', 'selenium grid', 'browserstack', 'saucelabs',
    'cypress io', 'espresso', 'xcuitest', 'xctest', 'detox', 'maestro',
    'k6', 'jmeter', 'gatling', 'locust', 'artillery', 'wiremock', 'mockserver', 'testcontainers',
    'pytest bdd', 'behave', 'hypothesis', 'hamcrest', 'mocha', 'jest', 'chai',
    'nunit', 'mstest', 'xunit', 'puppeteer', 'newman', 'readyapi', 'testrail', 'zephyr', 'xray',
    'sonarqube', 'snyk', 'owasp zap', 'burp suite', 'splunk', 'elasticsearch', 'kibana',
    'rabbitmq', 'kafka', 'redis', 'postgresql', 'mysql', 'sqlite', 'oracle', 'mssql',
    'snowflake', 'bigquery', 'dbt', 'airflow', 'spark', 'pandas', 'numpy',
    'fastapi', 'flask testing', 'django testing', 'spring boot', 'dotnet', 'node',
    'gradle', 'maven', 'bash', 'powershell', 'argo cd', 'openshift', 'rancher',
    'cloudformation', 'pulumi', 'vault', 'new relic', 'opentelemetry',
    'react', 'node.js', 'django', 'flask', 'spring',
]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
KEYWORDS_PATH = os.getenv('KEYWORDS_FILE', os.path.join(BASE_DIR, 'keywords.yaml'))

if yaml is not None and os.path.exists(KEYWORDS_PATH):
    try:
        data = yaml.safe_load(open(KEYWORDS_PATH, 'r', encoding='utf-8')) or {}
        role_terms = data.get('role_keywords')
        stack_terms = data.get('stack_keywords')
        if isinstance(role_terms, list) and role_terms:
            ROLE_TERMS = role_terms
        if isinstance(stack_terms, list) and stack_terms:
            STACK_TERMS = stack_terms
    except Exception:
        pass


SENIORITY_PATTERNS = [
    ('lead', r'\blead\b|\bprincipal\b|\bstaff\b'),
    ('senior', r'\bsenior\b|\bsr\.?\b'),
    ('mid', r'\bmid\b|\bintermediate\b'),
    ('junior', r'\bjunior\b|\bentry\b'),
]


def _read_pdf(path):
    if PdfReader is None:
        return ''
    try:
        reader = PdfReader(path)
        return '\n'.join((p.extract_text() or '') for p in reader.pages)
    except Exception:
        return ''


def _read_docx(path):
    if docx is None:
        return ''
    try:
        d = docx.Document(path)
        return '\n'.join(p.text for p in d.paragraphs)
    except Exception:
        return ''


def _read_txt(path):
    try:
        with open(path, 'r', encoding='utf-8', errors='ignore') as f:
            return f.read()
    except Exception:
        return ''


def read_resume_text(path):
    low = path.lower()
    if low.endswith('.pdf'):
        return _read_pdf(path)
    if low.endswith('.docx'):
        return _read_docx(path)
    if low.endswith('.txt'):
        return _read_txt(path)
    return ''


def pick_resume_file(resume_dir):
    if not os.path.isdir(resume_dir):
        return None
    candidates = []
    for n in os.listdir(resume_dir):
        p = os.path.join(resume_dir, n)
        if not os.path.isfile(p):
            continue
        if n.lower().endswith(('.pdf', '.docx', '.txt')):
            candidates.append(p)
    if not candidates:
        return None
    candidates.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return candidates[0]


def find_terms(text, terms):
    text_l = (text or '').lower()
    found = []
    for t in terms:
        if t in text_l:
            found.append(t)
    return found


def detect_seniority(text):
    t = (text or '').lower()
    for level, pat in SENIORITY_PATTERNS:
        if re.search(pat, t):
            return level
    return 'not_specified'


def extract_title_keywords(text, limit=10):
    t = re.sub(r'[^a-zA-Z0-9\s\-/]', ' ', (text or '').lower())
    words = [w for w in t.split() if len(w) > 2]
    counts = Counter(words)

    boosts = {
        'engineer', 'developer', 'analyst', 'manager', 'architect', 'specialist',
        'tester', 'qa', 'sdet', 'automation', 'data', 'security', 'cloud',
    }
    ranked = [w for w, c in counts.most_common(200) if w in boosts]

    out = []
    seen = set()
    for w in ranked:
        if w in seen:
            continue
        seen.add(w)
        out.append(w)
        if len(out) >= limit:
            break
    return out


def build_queries(profile):
    roles = profile.get('role_keywords', [])[:4]
    stack = profile.get('stack_keywords', [])[:4]
    seniority = profile.get('seniority', 'not_specified')
    seniority_part = '' if seniority == 'not_specified' else f'{seniority} '
    queries = []

    for r in roles[:3]:
        s = ' '.join(stack[:2]) if stack else ''
        q = f'remote {seniority_part}{r} united states {s}'.strip()
        queries.append(' '.join(q.split()))

    if stack:
        q = f'remote {seniority_part}engineer united states {" ".join(stack[:3])}'
        queries.append(' '.join(q.split()))

    queries.append(f'remote {seniority_part}software engineer united states'.strip())

    dedup = []
    seen = set()
    for q in queries:
        if q not in seen:
            dedup.append(q)
            seen.add(q)
    return dedup[:5]


def build_profile_from_resume(resume_path):
    text = read_resume_text(resume_path)
    role_keywords = find_terms(text, ROLE_TERMS)
    stack_keywords = find_terms(text, STACK_TERMS)
    title_keywords = extract_title_keywords(text)
    seniority = detect_seniority(text)

    if not role_keywords:
        role_keywords = ['software engineer', 'qa engineer', 'sdet']
    if not stack_keywords:
        stack_keywords = ['python', 'api testing', 'ci/cd']

    profile = {
        'resume_file': resume_path,
        'seniority': seniority,
        'role_keywords': role_keywords[:12],
        'stack_keywords': stack_keywords[:16],
        'title_keywords': title_keywords[:12],
        'remote_keywords': ['remote'],
        'us_keywords': ['united states', 'usa', 'us', 'u.s.', 'us only'],
    }
    profile['queries'] = build_queries(profile)
    return profile


def save_profile(profile, path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    if yaml is not None:
        with open(path, 'w', encoding='utf-8') as f:
            yaml.safe_dump(profile, f, sort_keys=False, allow_unicode=True)
        return

    # Minimal YAML-like fallback when PyYAML is not available.
    lines = []
    for key, value in profile.items():
        if isinstance(value, list):
            lines.append(f'{key}:')
            for item in value:
                lines.append(f'  - "{str(item)}"')
        else:
            lines.append(f'{key}: "{str(value)}"')
    with open(path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')


def build_profile(resume_dir, output_path):
    resume_path = pick_resume_file(resume_dir)
    if not resume_path:
        return None
    profile = build_profile_from_resume(resume_path)
    save_profile(profile, output_path)
    return profile


if __name__ == '__main__':
    base = os.path.dirname(os.path.abspath(__file__))
    resume_dir = os.getenv('RESUME_DIR', os.path.join(base, 'resume'))
    out_path = os.getenv('PROFILE_FILE', os.path.join(base, 'results', 'profile.yaml'))
    profile = build_profile(resume_dir, out_path)
    if profile:
        print(f'PROFILE_WRITTEN {out_path}')
        print(f'RESUME_FILE {profile.get("resume_file", "")}')
    else:
        print('NO_RESUME_FOUND')
