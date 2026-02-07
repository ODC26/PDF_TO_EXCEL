import pdfplumber

pdf_path = "NOMENCLATURE_NATIONALE_ANRP _ 2024.pdf"
page_num = 6

with pdfplumber.open(pdf_path) as pdf:
    page = pdf.pages[page_num - 1]
    words = page.extract_words(x_tolerance=2, y_tolerance=2)

lines = {}
for w in words:
    y = round(w['top'])
    lines.setdefault(y, []).append(w)

header = None
for y, ws in sorted(lines.items()):
    text = " ".join(w['text'] for w in sorted(ws, key=lambda x: x['x0']))
    tl = text.lower()
    if 'designation' in tl or 'désignation' in tl:
        header = (y, ws, text)
        break

print("HEADER:", header[2] if header else "None")
if header:
    y, ws, text = header
    for w in sorted(ws, key=lambda x: x['x0']):
        print(f"{w['text']:<20} x0={w['x0']:.2f} x1={w['x1']:.2f}")

print("\nSAMPLE LINES")
if header:
    count = 0
    for y, ws in sorted(lines.items()):
        if y <= header[0] + 5:
            continue
        line_text = " ".join(w['text'] for w in sorted(ws, key=lambda x: x['x0']))
        if line_text.strip():
            print(f"y={y} -> {line_text}")
            for w in sorted(ws, key=lambda x: x['x0']):
                print(f"   {w['text']:<20} x0={w['x0']:.2f} x1={w['x1']:.2f}")
            count += 1
        if count >= 5:
            break
