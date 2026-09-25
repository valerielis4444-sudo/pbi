# Генерирует Макрос_ПБЕ.bas: встраивает M-код запросов из power_query/*.m
import glob, os
Q = [("ДПБ_История","01_ДПБ_История.m"),("ДПБ_Файлы","02_ДПБ_Файлы.m"),("ДПБ_Факт","03_ДПБ_Факт.m"),
     ("ДПБ_Календарь","04_ДПБ_Календарь.m"),("ДПБ_ПБЕ","05_ДПБ_ПБЕ.m"),("ДПБ_Параметры","06_ДПБ_Параметры.m"),
     ("ДПБ_Сверка","07_ДПБ_Сверка.m")]
def vb(s): return s.replace('"','""')
out=[]
for i,(n,f) in enumerate(Q):
    lines=[l for l in open('power_query/'+f,encoding='utf8').read().splitlines() if not l.strip().startswith('//')]
    out.append(f"Private Function QM{i}() As String\n    Dim s As String")
    for l in lines:
        out.append(f'    s = s & "{vb(l)}" & vbLf')
    out.append(f"    QM{i} = s\nEnd Function\n")
qfuncs="\n".join(out)

P = "'ДПБ_Календарь'[МесяцИндекс]"
def prev(expr):
    return (f"VAR mn = MIN({P}) VAR n = MAX({P}) - mn + 1 VAR perv = CALCULATE(MIN({P});ALL('ДПБ_Календарь')) "
            f"RETURN IF(mn - n >= perv;CALCULATE({expr};FILTER(ALL('ДПБ_Календарь');{P} >= mn - n && {P} <= mn - 1)))")
def ly(expr):
    return (f"VAR mn = MIN({P}) VAR mx = MAX({P}) VAR perv = CALCULATE(MIN({P});ALL('ДПБ_Календарь')) "
            f"RETURN IF(mn - 12 >= perv;CALCULATE({expr};FILTER(ALL('ДПБ_Календарь');{P} >= mn - 12 && {P} <= mx - 12)))")
M = [
 ("Оборот Дт","SUM('ДПБ_Факт'[ОборотДтТг])","D0"),
 ("Оборот Кт","SUM('ДПБ_Факт'[ОборотКтТг])","D0"),
 ("Месяцев ПБЕ","COUNTROWS('ДПБ_Календарь')","W"),
 ("Дней ПБЕ","SUM('ДПБ_Календарь'[Дней])","W"),
 ("Средняя ДЗ","DIVIDE(SUM('ДПБ_Факт'[ДЗТг]);[Месяцев ПБЕ])","D0"),
 ("ДЗ на конец","VAR d = MAX('ДПБ_Календарь'[Дата]) RETURN CALCULATE(SUM('ДПБ_Факт'[ДЗТг]);'ДПБ_Календарь'[Дата] = d)","D0"),
 ("Авансы на конец","VAR d = MAX('ДПБ_Календарь'[Дата]) RETURN CALCULATE(SUM('ДПБ_Факт'[АвансыТг]);'ДПБ_Календарь'[Дата] = d)","D0"),
 ("DSO","VAR t = [Оборот Дт] RETURN IF(t > 0;DIVIDE([Средняя ДЗ] * [Дней ПБЕ];t))","D1"),
 ("Инкассация","DIVIDE([Оборот Кт];[Оборот Дт])","P1"),
 ("Оборот Дт компании","CALCULATE([Оборот Дт];ALL('ДПБ_ПБЕ');ALL('ДПБ_Факт'[ПБЕ]))","D0"),
 ("Доля оборота","DIVIDE([Оборот Дт];[Оборот Дт компании])","P1"),
 ("Средняя ДЗ компании","CALCULATE([Средняя ДЗ];ALL('ДПБ_ПБЕ');ALL('ДПБ_Факт'[ПБЕ]))","D0"),
 ("Доля ДЗ","DIVIDE([Средняя ДЗ];[Средняя ДЗ компании])","P1"),
 ("Перекос ДЗ","[Доля ДЗ] - [Доля оборота]","P1"),
 ("Цель DSO","MAX('ДПБ_Параметры'[Цель_DSO])","W"),
 ("Норматив ДЗ","DIVIDE([Оборот Дт];[Дней ПБЕ]) * [Цель DSO]","D0"),
 ("Излишек ДЗ","SUMX(VALUES('ДПБ_ПБЕ'[ПБЕ]);VAR s = [Средняя ДЗ] VAR n = [Норматив ДЗ] VAR n0 = IF(n > 0;n;0) RETURN IF(s > n0;s - n0;0))","D0"),
 ("Оборот Дт пред период",prev("[Оборот Дт]"),"D0"),
 ("Прирост оборота","VAR b = [Оборот Дт пред период] RETURN IF(b > 0;DIVIDE([Оборот Дт] - b;b))","P1"),
 ("DSO пред период",prev("[DSO]"),"D1"),
 ("Дельта DSO","VAR a = [DSO] VAR b = [DSO пред период] RETURN IF(NOT ISBLANK(a) && NOT ISBLANK(b);a - b)","D1"),
 ("Оборот Дт год назад",ly("[Оборот Дт]"),"D0"),
 ("Прирост оборота гг","VAR b = [Оборот Дт год назад] RETURN IF(b > 0;DIVIDE([Оборот Дт] - b;b))","P1"),
 ("Дельта DSO значимая","IF([Доля оборота] >= MAX('ДПБ_Параметры'[Порог_доли]);[Дельта DSO])","D1"),
 ("Период ПБЕ текст","FORMAT(MIN('ДПБ_Календарь'[Дата]);\"MM.yyyy\") & \" – \" & FORMAT(MAX('ДПБ_Календарь'[Дата]);\"MM.yyyy\") & \"  ·  \" & [Месяцев ПБЕ] & \" мес.\"","G"),
 ("Ранг ПБЕ","IF(HASONEVALUE('ДПБ_ПБЕ'[ПБЕ]) && NOT ISBLANK([Оборот Дт]);RANKX(ALLSELECTED('ДПБ_ПБЕ'[ПБЕ]);[Оборот Дт]))","W"),
 ("Оборот Дт млн","DIVIDE([Оборот Дт];1000000)","D1"),
 ("Средняя ДЗ млн","DIVIDE([Средняя ДЗ];1000000)","D1"),
 ("ДЗ на конец млн","DIVIDE([ДЗ на конец];1000000)","D1"),
 ("Авансы на конец млн","DIVIDE([Авансы на конец];1000000)","D1"),
 ("Излишек ДЗ млн","DIVIDE([Излишек ДЗ];1000000)","D1"),
]
mlines="\n".join(f'    AddM "{vb(n)}", "{vb(f)}", f{fmt}' for n,f,fmt in M)
tpl=open('macro_template.bas',encoding='utf8').read()
res=tpl.replace('%%QFUNCS%%',qfuncs).replace('%%MEASURES%%',mlines)
open('Макрос_ПБЕ.bas','w',encoding='utf8').write(res)
res.encode('cp1251')
print('ok', len(res), 'chars; measures', len(M))
# документация мер
with open('Меры_DAX_ПБЕ.md','w',encoding='utf8') as fo:
    fo.write('# Меры DAX для дашборда ПБЕ (создаются макросом автоматически)\n\nТаблица мер: `ДПБ_Факт`. Формат: D0/D1 — десятичное 0/1 знак, P1 — процент, W — целое, G — общий.\n\n')
    for n,f,fmt in M: fo.write(f'**{n}** — {fmt}\n```\n{f}\n```\n\n')
