import subprocess,json,shutil,os
from pathlib import Path
R=Path(__file__).parent; S=R/'sources';B=R/'build';B.mkdir(exist_ok=True)
units=[('status',['2_4_01_status.c'],'available=7\n'),('price',['2_4_02_main.c','2_4_02_price.c'],'total=3600\n'),('label_v1',['2_5_01_label_v1.c'],'revision=1\n'),('label_v2',['2_5_02_label_v2.c'],'revision=2\n'),('portable',['2_6_01_portable.c'],'total=11\n'),('warning_fixed',['2_6_02_warning_fixed.c'],'difference=3\n')]
results=[]
for compiler in ['gcc','clang']:
    if not shutil.which(compiler):results.append({'compiler':compiler,'status':'unavailable'});continue
    for name,files,expected in units:
        exe=B/(compiler+'_'+name)
        cmd=[compiler,'-std=c11','-Wall','-Wextra','-Wpedantic','-Werror','-O0']+[str(S/f) for f in files]+['-o',str(exe)]
        c=subprocess.run(cmd,text=True,capture_output=True)
        r=subprocess.run([str(exe)],text=True,capture_output=True) if c.returncode==0 else None
        ok=c.returncode==0 and r.returncode==0 and r.stdout==expected
        results.append({'compiler':compiler,'unit':name,'sources':files,'command':cmd,'compile_rc':c.returncode,'compile_stderr':c.stderr,'run_rc':None if r is None else r.returncode,'stdout':None if r is None else r.stdout,'expected':expected,'passed':ok})
Path(R/'c11_verification_2_4-2_6.json').write_text(json.dumps({'compiler_matrix':results,'passed':sum(x.get('passed',False) for x in results),'total':len(results)},ensure_ascii=False,indent=2),encoding='utf-8')
print('passed',sum(x.get('passed',False) for x in results),'/',len(results));
for x in results:
    if not x.get('passed'):print(x)