with open('generate_json_hierarchy.py', 'r', encoding='utf-8') as f:
    text = f.read()

start_idx = text.find('def _write_exact_summary_layout(ws):')
if start_idx != -1:
    header = text[:start_idx]
    body = text[start_idx:]
    
    body = body.replace('value=\"Total Sub modules\"', 'value=\"Total Sub module\'s\"')
    body = body.replace('value=\"Carry-Over Sub modules\"', 'value=\"Carry-Over Sub module\'s\"')
    body = body.replace('value=\"Carryover Sub modules\"', 'value=\"Carryover Sub module\'s\"')
    body = body.replace('value=\"Exclusive Sub modules (A+B+C)\"', 'value=\"Exclusive Sub module\'s (A+B+C)\"')
    body = body.replace('value=\"Exclusive Sub modules\"', 'value=\"Exclusive Sub module\'s\"')
    body = body.replace('value=\"A. Modified Sub modules\"', 'value=\"A. Modified Sub module\'s\"')
    body = body.replace('value=\"B. Newly Added Sub modules\"', 'value=\"B. Newly Added Sub module\'s\"')
    body = body.replace('value=\"C. Replaced Sub modules\"', 'value=\"C. Replaced Sub module\'s\"')
    body = body.replace('value=\"Deleted Sub modules\"', 'value=\"Deleted Sub module\'s\"')
    body = body.replace('value=\"Modules\"', 'value=\"Module\'s\"')
    body = body.replace('value=\"Modules *\"', 'value=\"Module\'s *\"')
    
    with open('generate_json_hierarchy.py', 'w', encoding='utf-8') as f:
        f.write(header + body)
    print('Replacements done.')
