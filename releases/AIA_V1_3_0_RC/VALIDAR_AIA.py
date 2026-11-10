# -*- coding: utf-8 -*-
from pathlib import Path
from types import SimpleNamespace, MethodType
import importlib.util, py_compile, sys, math, tempfile

BASE = Path(__file__).parent
SRC = BASE / "AIA_V1_3_0_RC.py"

checks=[]
def ok(name, cond, detail=""):
    checks.append((name, bool(cond), detail))
    if not cond:
        raise AssertionError(name + (": " + detail if detail else ""))

print("AIA V1.3.0 RC - validacao local")
print("Arquivo:", SRC)
print()

py_compile.compile(str(SRC), doraise=True)
ok("Sintaxe Python", True)

spec=importlib.util.spec_from_file_location("aia_oficial", SRC)
mod=importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
ok("Importacao do modulo", True)
ok("Versao", getattr(mod,"VERSION",None)=="1.3.0 RC", str(getattr(mod,"VERSION",None)))

# C x D-I deliberadamente divergentes.
records=[
    {"vision":"NOK","reference":"Defeito confirmado","real":0,"detected":2},
    {"vision":"OK","reference":"OK confirmado","real":4,"detected":0},
    {"vision":"NOK","reference":"Defeito confirmado","real":3,"detected":3},
    {"vision":"OK","reference":"OK confirmado","real":0,"detected":0},
]
c=mod.matrix_c_from_records(records)
di=mod.matrix_di_from_records(records)
ok("C e D-I independentes", (c["TP"],c["TN"],c["FP"],c["FN"]) != (di["TP"],di["TN"],di["FP"],di["FN"]), f"C={c} D-I={di}")

# D-O por linha: nao pode compensar FP de uma linha com FN de outra.
dummy=SimpleNamespace()
dummy.rows=[
    {"id":"A","real":4,"det":8,"detection_observed":True,"detection_available":True,"defects":[],"tp_d":None,"fp_d":None,"fn_d":None,"detection_explicit_matrix":False},
    {"id":"B","real":5,"det":2,"detection_observed":True,"detection_available":True,"defects":[],"tp_d":None,"fp_d":None,"fn_d":None,"detection_explicit_matrix":False},
]
dummy.source_analysis={}; dummy.source_file=""
dummy.detection_for_row=MethodType(mod.App.detection_for_row,dummy)
do=mod._aia_v130_detection_occurrence(dummy)
ok("D-O soma por linha", (do["real"],do["det"],do["tp"],do["fp"],do["fn"])==(9,10,6,4,3), str(do))

# Pareamento explicito conservativo deve ser preservado.
dummy2=SimpleNamespace()
dummy2.rows=[{"id":"E","real":4,"det":4,"correspondentes":3,"detection_observed":True,"detection_available":True,"defects":[],"tp_d":3,"fp_d":1,"fn_d":1,"detection_explicit_matrix":True}]
dummy2.source_analysis={}; dummy2.source_file=""
dummy2.detection_for_row=MethodType(mod.App.detection_for_row,dummy2)
do2=mod._aia_v130_detection_occurrence(dummy2)
ok("D-O explicito conservativo", (do2["tp"],do2["fp"],do2["fn"])==(3,1,1), str(do2))

# Correspondentes permite FP e FN na mesma inspeção.
dummy3=SimpleNamespace()
dummy3.rows=[{"id":"C","real":5,"det":7,"correspondentes":3,"detection_observed":True,"detection_available":True,"defects":[]}]
do3=mod._aia_v130_detection_occurrence(dummy3)
ok("D-O com FP e FN simultaneos", (do3["tp"],do3["fp"],do3["fn"],do3["status"])==(3,4,2,"EXPLÍCITO"), str(do3))
ok("D-O sem correspondentes e estimado", do["status"]=="ESTIMADO", str(do))

# Entrada operacional TP-D/FP-D/FN-D preserva os mesmos totais do motor.
tp_d, fp_d, fn_d = 1, 2, 1
row_operador=mod.normalize_row({"id":"OP","real":tp_d+fn_d,"det":tp_d+fp_d,"correspondentes":tp_d,"detection_observed":True})
got=mod._aia_v130_detection_for_row(dummy3,row_operador)
ok("Entrada TP FP FN preserva D-O", (got["real"],got["det"],got["tp"],got["fp"],got["fn"])==(2,3,1,2,1), str(got))

# Projeto JSON: campos novos e projetos antigos devem ser aceitos.
payload=mod.aia_v130_validate_project_payload({"version":"1.3.0 RC","metadata":{"projeto":"Teste"},"rows":[{"id":"J1","real":5,"det":7,"correspondentes":3,"detection_observed":True}],"class_thresholds":{"Risco":{"threshold":"0.82","confidence":"82"}},"history":[],"comparison_archive":[],"calibration_history":[]})
ok("Projeto JSON V1.3 restaurável", payload["rows"][0]["correspondentes"]==3 and payload["class_thresholds"]["RISCO"]["threshold"]=="0.82", str(payload))
legacy=mod.aia_v130_validate_project_payload({"metadata":{},"rows":[{"id":"L1","real":1,"det":1}],"history":[],"comparison_archive":[],"calibration_history":[]})
ok("Projeto JSON legado compatível", legacy["rows"][0]["id"]=="L1" and legacy["class_thresholds"]=={}, str(legacy))

# Catálogo global: tipos são rastreáveis, não mudam D-O e Poro vira Porosidade.
catalog=mod.aia_v130_sorted_catalog(["Poro", "Risco", "Tipo especial"])
row_catalog=mod.normalize_row({"id":"CAT1","real":1,"det":1,"defect_types":["Poro", "Risco", "Poro"]})
ok("Catálogo global normaliza Poro", "Porosidade" in catalog and row_catalog["defect_types"]==["Porosidade","Risco"], str(row_catalog))

# CSV da própria aplicação deve voltar por ABRIR EXCEL/CSV com metadados e D-O.
class Var:
    def __init__(self, value=""): self.value=value
    def get(self): return self.value
    def set(self, value): self.value=str(value)
class Status:
    def set(self, _value): pass
meta_keys=("cliente","local","projeto","linha","tecnico","data","threshold","gabarito")
csv_source=SimpleNamespace(rows=[mod.normalize_row({"id":"CSV1","vision":"NOK","ref":"Defeito confirmado","real":2,"det":3,"correspondentes":1,"detection_observed":True,"obs":"teste CSV","defect_types":["Risco","Porosidade"]})],class_thresholds={"TRINCA":{"threshold":"80","confidence":"","notes":""}},meta={k:Var("Projeto CSV" if k=="projeto" else "") for k in meta_keys},status=Status())
csv_source.detection_for_row=MethodType(mod._aia_v130_detection_for_row,csv_source)
with tempfile.TemporaryDirectory() as temp_dir:
    csv_path=Path(temp_dir)/"aia.csv"
    old_save,old_info=mod.filedialog.asksaveasfilename,mod.messagebox.showinfo
    mod.filedialog.asksaveasfilename=lambda **_kwargs:str(csv_path); mod.messagebox.showinfo=lambda *args,**kwargs:None
    try: mod._aia_v130_export_csv(csv_source)
    finally: mod.filedialog.asksaveasfilename=old_save; mod.messagebox.showinfo=old_info
    csv_target=SimpleNamespace(rows=[],history=[],comparison_archive=[],calibration_history=[],class_thresholds={},source_file="",source_analysis={},meta={k:Var() for k in meta_keys},status=Status(),classify=lambda r:r,refresh_all=lambda:None,_archive_current_collection=lambda:None)
    old_info,old_error=mod.messagebox.showinfo,mod.messagebox.showerror
    mod.messagebox.showinfo=lambda *args,**kwargs:None; mod.messagebox.showerror=lambda *args,**kwargs:None
    try: mod._aia_v130_import_csv(csv_target,str(csv_path))
    finally: mod.messagebox.showinfo=old_info; mod.messagebox.showerror=old_error
    ok("CSV AIA reabre com dados", len(csv_target.rows)==1 and csv_target.rows[0]["id"]=="CSV1" and csv_target.rows[0]["correspondentes"]==1 and csv_target.rows[0]["defect_types"]==["Risco","Porosidade"] and csv_target.meta["projeto"].get()=="Projeto CSV", str(csv_target.rows))

# PDF deve ser produzido com a tabela de registros em página paisagem.
pdf_source=SimpleNamespace(rows=csv_source.rows,class_thresholds=csv_source.class_thresholds,status=Status())
pdf_source.detection_for_row=MethodType(mod._aia_v130_detection_for_row,pdf_source)
pdf_source.detection_totals=lambda:{"inspection":{"N":1,"TP":1,"TN":0,"FP":0,"FN":0,"Accuracy":1.0,"Recall / Sensibilidade":1.0,"Precision":1.0,"F1":1.0,"method":"teste"}}
pdf_source.defect_counts=lambda:{"real":2,"det":3,"tp":1,"fp":2,"fn":1,"status":"EXPLÍCITO","method":"teste"}
pdf_source.display_metrics=lambda:{"N":1,"TP":1,"VN/TN":0,"FP":0,"FN":0,"Accuracy":1.0}
with tempfile.TemporaryDirectory() as temp_dir:
    pdf_path=Path(temp_dir)/"aia.pdf"; old_save,old_info=mod.filedialog.asksaveasfilename,mod.messagebox.showinfo
    mod.filedialog.asksaveasfilename=lambda **_kwargs:str(pdf_path); mod.messagebox.showinfo=lambda *args,**kwargs:None
    try: mod._aia_v130_export_pdf(pdf_source)
    finally: mod.filedialog.asksaveasfilename=old_save; mod.messagebox.showinfo=old_info
    ok("PDF com dados registrados", pdf_path.exists() and pdf_path.stat().st_size>1000 and pdf_path.read_bytes().startswith(b"%PDF"), str(pdf_path))

# Perfis de Meta oficiais.
good=mod.aia_v120_meta_profile_targets("BOM")
excellent=mod.aia_v120_meta_profile_targets("EXCELENTE")
ok("Perfil BOM", good=={"accuracy":0.90,"precision":0.90,"recall":0.95,"f1":0.92,"fp100":5.0,"fn100":3.0}, str(good))
ok("Perfil EXCELENTE", excellent=={"accuracy":0.95,"precision":0.95,"recall":0.98,"f1":0.96,"fp100":3.0,"fn100":1.0}, str(excellent))

# Runtime final: pontos de regressao essenciais.
bind=getattr(mod,"FINAL_RUNTIME_BINDINGS",{})
for name in ["view_operacao","view_dashboard","view_graficos","view_comparacao","view_tecnica","view_engenharia","view_calibracao","view_expectativa","view_auditoria","import_excel","load_project","export_excel","export_csv","export_executive_report","detection_totals","defect_counts"]:
    ok("Binding "+name, name in bind)

print("RESULTADO")
for name, passed, detail in checks:
    print(("[OK] " if passed else "[FALHA] ") + name)
print()
print(f"{sum(p for _,p,_ in checks)}/{len(checks)} verificacoes aprovadas.")
print("Validacao concluida com sucesso.")
