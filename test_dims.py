from base_scraper import parse_dimensions

tests = [
    ('Height 19.7" x Width 12.2" x Width 6.7"',  {"Width":"6.7","Depth":"12.2","Height":"19.7"}),
    ('Low: 59.1" W x 17.7" D x 23.6" H',          {"Width":"59.1","Depth":"17.7","Height":"23.6"}),
    ('W 25" x D 12" x H 22.5"',                    {"Width":"25","Depth":"12","Height":"22.5"}),
    ('16.00" W x 5.00" H x 10.75" D',              {"Width":"16.00","Depth":"10.75","Height":"5.00"}),
    ('6.3" W x 7.9" H',                            {"Width":"6.3","Height":"7.9"}),
    ('Canyon Vase Large',                           {}),  # no dims → should return only Dimensions key
]

all_pass = True
for raw, expected in tests:
    r = parse_dimensions(raw)
    ok = all(r.get(k) == v for k, v in expected.items())
    status = "PASS" if ok else "FAIL"
    if not ok:
        all_pass = False
    print(f"[{status}] {raw[:55]}")
    print(f"       W={r.get('Width','-')} D={r.get('Depth','-')} H={r.get('Height','-')}")
    if not ok:
        print(f"  EXPECTED: {expected}")
    print()

print("ALL PASS" if all_pass else "SOME FAILED")
