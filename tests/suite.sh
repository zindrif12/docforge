#!/bin/bash
# DocForge test suite — run against a MOCK_AI=1 server on :8090
B=http://localhost:8090
P=0; F=0
ck(){ if [ "$1" = "$2" ]; then P=$((P+1)); echo "  PASS $3"; else F=$((F+1)); echo "  FAIL $3 (want [$2] got [$1])"; fi; }

echo "== health =="
H=$(curl -s $B/api/health)
ck "$(echo $H | grep -o '"ok":true')" '"ok":true' "health ok"
ck "$(echo $H | grep -o '"provider":"mock"')" '"provider":"mock"' "mock provider active"
ck "$(echo $H | grep -o '"modules":\["fill","sheet","doc","deck"\]')" '"modules":["fill","sheet","doc","deck"]' "four modules advertised"

echo "== template analyze =="
# build a tiny docx template with placeholders via python-docx
python3 - <<'EOF'
import docx
d = docx.Document()
d.add_heading('Offer of Employment', 0)
p = d.add_paragraph('Dear {{candidate_name}}, we offer you the role of {{job_title}} at {{company}}.')
d.add_paragraph('Salary: {{salary}}. Start: {{start_date}}. Manager: {{manager}}.')
d.save('/tmp/df_tpl.docx')
EOF
B64=$(base64 -w0 /tmp/df_tpl.docx)
R=$(curl -s -X POST $B/api/template/analyze -H 'Content-Type: application/json' -d "{\"filename\":\"offer.docx\",\"fileBase64\":\"$B64\"}")
TID=$(echo $R | grep -o '"id":"[^"]*' | cut -d'"' -f4)
ck "$(echo $R | grep -o 'candidate_name')" "candidate_name" "placeholders discovered"
ck "$(echo $R | grep -o '"kind":"docx"')" '"kind":"docx"' "kind detected docx"
ck "$([ -n "$TID" ] && echo yes)" "yes" "template id returned"

R=$(curl -s -X POST $B/api/template/analyze -H 'Content-Type: application/json' -d "{\"filename\":\"offer.txt\",\"fileBase64\":\"$B64\"}")
ck "$(echo $R | grep -o 'Only .docx and .xlsx')" "Only .docx and .xlsx" "rejects unsupported extension"

# a docx with no placeholders → 422
python3 -c "import docx; d=docx.Document(); d.add_paragraph('plain'); d.save('/tmp/df_plain.docx')"
B64P=$(base64 -w0 /tmp/df_plain.docx)
CODE=$(curl -s -o /dev/null -w '%{http_code}' -X POST $B/api/template/analyze -H 'Content-Type: application/json' -d "{\"filename\":\"p.docx\",\"fileBase64\":\"$B64P\"}")
ck "$CODE" "422" "422 when no placeholders"

R=$(curl -s $B/api/templates)
ck "$(echo $R | grep -o 'offer.docx' | head -1)" "offer.docx" "template listed in library"

echo "== fill =="
R=$(curl -s -X POST $B/api/template/fill -H 'Content-Type: application/json' -d "{\"templateId\":\"$TID\",\"sourceText\":\"Email thread: We agreed to hire Nadeesha Perera as Senior Software Engineer at Meridian Labs, salary LKR 480,000, starting 1 Sept 2026, reporting to R. Fernando.\"}")
ck "$(echo $R | grep -o '"confidence":92')" '"confidence":92' "fill returns confidence"
ck "$(echo $R | grep -o 'Nadeesha Perera' | head -1)" "Nadeesha Perera" "mapped values returned"
echo $R | python3 -c "
import sys, json, base64
d = json.load(sys.stdin)
open('/tmp/df_filled.docx','wb').write(base64.b64decode(d['fileBase64']))
print('saved', d['filename'])"
python3 - <<'EOF'
import docx
t = "\n".join(p.text for p in docx.Document('/tmp/df_filled.docx').paragraphs)
assert 'Nadeesha Perera' in t and 'Senior Software Engineer' in t and '{{' not in t
print('  PASS filled docx opens + placeholders gone (python-docx verified)')
EOF
[ $? -eq 0 ] && P=$((P+1)) || F=$((F+1))

CODE=$(curl -s -o /dev/null -w '%{http_code}' -X POST $B/api/template/fill -H 'Content-Type: application/json' -d '{"templateId":"nope","sourceText":"some text long enough"}')
ck "$CODE" "404" "404 for unknown template"

echo "== sheet =="
R=$(curl -s -X POST $B/api/generate/sheet -H 'Content-Type: application/json' -d '{"prompt":"12 month bakery cash flow"}')
ck "$(echo $R | grep -o '"filename":"bakery_cash_flow.xlsx"')" '"filename":"bakery_cash_flow.xlsx"' "sheet filename"
echo $R | python3 -c "
import sys, json, base64
d = json.load(sys.stdin)
open('/tmp/df_gen.xlsx','wb').write(base64.b64decode(d['fileBase64']))"
python3 - <<'EOF'
import openpyxl
wb = openpyxl.load_workbook('/tmp/df_gen.xlsx')
ws = wb['Cash Flow']
assert ws['D2'].value == '=B2-C2', ws['D2'].value
assert ws['B4'].value == '=SUM(B2:B3)'
assert wb.sheetnames == ['Cash Flow','Assumptions']
print('  PASS generated xlsx has LIVE formulas (openpyxl verified)')
EOF
[ $? -eq 0 ] && P=$((P+1)) || F=$((F+1))

echo "== doc =="
R=$(curl -s -X POST $B/api/generate/doc -H 'Content-Type: application/json' -d '{"prompt":"proposal for meridian project"}')
ck "$(echo $R | grep -o '"title":"Meridian Project Proposal"')" '"title":"Meridian Project Proposal"' "doc title"
echo $R | python3 -c "
import sys, json, base64
d = json.load(sys.stdin)
open('/tmp/df_gen.docx','wb').write(base64.b64decode(d['fileBase64']))"
python3 - <<'EOF'
import docx
d = docx.Document('/tmp/df_gen.docx')
styles = [p.style.name for p in d.paragraphs]
assert 'Title' in styles and 'Heading 1' in styles
assert len(d.tables) == 1 and d.tables[0].cell(0,0).text == 'Item'
print('  PASS generated docx has styles + table (python-docx verified)')
EOF
[ $? -eq 0 ] && P=$((P+1)) || F=$((F+1))

echo "== deck =="
R=$(curl -s -X POST $B/api/generate/deck -H 'Content-Type: application/json' -d '{"prompt":"kickoff deck for meridian","accent":"7C3AED"}')
ck "$(echo $R | grep -o '"filename":"meridian_kickoff.pptx"')" '"filename":"meridian_kickoff.pptx"' "deck filename"
echo $R | python3 -c "
import sys, json, base64
d = json.load(sys.stdin)
open('/tmp/df_gen.pptx','wb').write(base64.b64decode(d['fileBase64']))"
python3 - <<'EOF'
import pptx
pr = pptx.Presentation('/tmp/df_gen.pptx')
assert len(pr.slides) == 4
texts = [sh.text_frame.text for s in pr.slides for sh in s.shapes if sh.has_text_frame]
assert any('Why now' in t for t in texts)
print('  PASS generated pptx opens with', len(pr.slides), 'slides (python-pptx verified)')
EOF
[ $? -eq 0 ] && P=$((P+1)) || F=$((F+1))

echo "== history & hygiene =="
R=$(curl -s $B/api/history)
ck "$(echo $R | grep -o '"module":"fill"' | head -1)" '"module":"fill"' "history records fill"
ck "$(echo $R | grep -o '"module":"deck"' | head -1)" '"module":"deck"' "history records deck"
# privacy: filled VALUES must never appear in history
ck "$(echo $R | grep -o 'Nadeesha' | head -1)" "" "history contains no filled values (privacy)"

CODE=$(curl -s -o /dev/null -w '%{http_code}' -X POST $B/api/generate/sheet -H 'Content-Type: application/json' -d '{"prompt":"x"}')
ck "$CODE" "400" "400 for too-short prompt"

R=$(curl -s -X DELETE $B/api/templates/$TID)
ck "$(echo $R | grep -o '"deleted"')" '"deleted"' "template delete"

R=$(curl -s $B/)
ck "$(echo $R | grep -o '<title>DocForge' | head -1)" "<title>DocForge" "frontend served at /"

echo "========================"
echo "FINAL: $P passed, $F failed"
[ $F -eq 0 ]
