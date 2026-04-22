import json
d = json.load(open('data.json', 'r', encoding='utf-8'))
depts = ['Design','Project Management','Product Management','Technical Writing','Product Marketing','EIT']
for e in d['employees']:
    if e['department'] in depts:
        skills = [(s['skill'], s['proficiency']) for v in e['skills'].values() for s in v]
        print(f"{e['name']} ({e['department']}): {len(skills)} skills")
        for sk, pr in skills:
            print(f"  - {sk} [{pr}]")
        print()
