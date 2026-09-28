Q=["01_Тек_Параметры.m","02_Тек_Остатки.m","03_Тек_Продажи.m","04_Тек_Итог.m"]
def vb(s): return s.replace('"','""')
out=[]
for i,f in enumerate(Q):
    lines=[l for l in open('power_query/'+f,encoding='utf8').read().splitlines() if not l.strip().startswith('//')]
    out.append(f"Private Function QM{i}() As String\n    Dim s As String")
    out+= [f'    s = s & "{vb(l)}" & vbLf' for l in lines]
    out.append(f"    QM{i} = s\nEnd Function\n")
res=open('macro_template.bas',encoding='utf8').read().replace('%%QFUNCS%%',"\n".join(out))
open('Макрос_Текущая_оборачиваемость.bas','w',encoding='utf8').write(res); res.encode('cp1251'); print('ok',len(res))
