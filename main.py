import re

aha = {
    '%02%': 'Ergene Belediyesi',
    '%03%': 'İkitelli Vergi Dairesi',
    '%04%': '548621598723',
    '%05%': 'ergene@belediye.com',
    '%06%': '+905052615688',
    '%07%': '59772264',
    '%09%': 'Belediyecilik işte',
    '%10%': 'ERGENE BELEDİYESİ',
    '%11%': 'ergenebelediyesi@hs01.kep.tr',
    '%12%': 'T.C. ERGENE BELEDİYE BAŞKANLIĞI',
    '%13%': 'Ulaş Mahallesi Atatürk Bulvarı No: 84/ERGENE /TEKİRDAĞ',
    '%14%': 'ticaret müdürlüğü',
    '%15%': '54564678',
    '%16%': '1458 nolu kanun',
}


def multiple_replace(string, rep_dict):
    pattern = re.compile("|".join([re.escape(k) for k in sorted(rep_dict, key=len, reverse=True)]), flags=re.DOTALL)
    return pattern.sub(lambda x: rep_dict[x.group(0)], string)


with open('h01.txt', 'r', encoding='utf-8') as rf:
    with open('test_ho1.txt', 'w', encoding='utf-8') as wf:
        for line in rf:
            x = re.findall(r"\B%\d\d\b%", line)
            if x:
                wf.write(multiple_replace(line, aha))
            else:
                wf.write(line)
