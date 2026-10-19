import openpyxl, json
from urllib.parse import urlparse

wb = openpyxl.load_workbook(r'c:\Users\trong\Downloads\Telegram Desktop\accounts_OK.xlsx')
ws = wb.active

hospitals = {}
for row in ws.iter_rows(min_row=2, values_only=True):
    url, username, password, domain = row
    if not url:
        continue
    parsed = urlparse(str(url).strip())
    domain = parsed.netloc.lower().replace('www.', '')
    if not domain:
        continue
    subdomain = domain.split('.')[0]
    base_url = parsed.scheme + '://' + domain

    if domain not in hospitals:
        hospitals[domain] = {
            'name': subdomain.upper(),
            'domain': domain,
            'link': base_url + '/vnpthis'
        }

result = sorted(hospitals.values(), key=lambda x: x['domain'])

output = {
    'total_hospitals': len(result),
    'hospitals': result
}

out_path = r'd:\Workspaces\Python_Code\Bot_vncare\hospitals.json'
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print(f'Done! {len(result)} benh vien -> hospitals.json')
for h in result:
    print(f'  {h["link"]}')
