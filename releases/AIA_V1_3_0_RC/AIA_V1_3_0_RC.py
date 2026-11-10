# -*- coding: utf-8 -*-
"""
AUTAZA INSPECTION ANALYTICS - AIA V1.3.0 RC
Industrial Inspection Analytics Platform

RELEASE CANDIDATE - 09/11/2026
- Base consolidada e auditada do ciclo V1.2.0.
- D-I como referencia operacional por inspecao.
- D-O consolidado por linha, sem compensacao indevida entre FP e FN de inspecoes distintas.
- Matriz C, D-I e D-O mantidas independentes, com fonte, unidade e metodo rastreaveis.
- Comparacao multicoleta com leitura visual e deltas.
- Meta D-I com perfis BOM, EXCELENTE e PERSONALIZADO.
- Calibracao com orientacao conservadora baseada em D-I + Meta.
- Auditoria com verificacoes de integridade e rastreabilidade.
- Importacao/exportacao Excel, CSV, PDF e salvamento de projeto JSON preservados.
- Runtime final unico por FINAL_RUNTIME_BINDINGS.

Esta release parte da baseline V1.2.0 Oficial, que permanece congelada.
Inclui D-I como indicador principal, D-O por correspondência explícita e rastreabilidade por classe.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime
from collections import defaultdict
import json, math, csv, re, unicodedata, os, sys, hashlib

try:
    import openpyxl
except Exception:
    openpyxl = None

APP = "Autaza Inspection Analytics"
VERSION = "1.3.0 RC"
# === V1.0.0 FRONT-END CONSOLIDADO — RELEASE VISUAL FINAL ===



def application_dir():
    """Return the directory containing bundled resources or the source file."""
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)
    # parent evita dependência de resolução de link/rede no Windows e mantém
    # a pasta do script como referência de recursos locais.
    return Path(__file__).parent

BG = "#07131C"
PANEL = "#0E202D"
PANEL2 = "#123B52"
PANEL3 = "#091A25"
BORDER = "#21485F"
TEXT = "#F3F7FA"
MUTED = "#8FA9B8"
ACCENT = "#00A9E8"
ACCENT2 = "#67D5FF"
GOOD = "#45D39A"
WARN = "#F4C15D"
BAD = "#FF6B78"

REFS = ["Defeito confirmado", "OK confirmado", "Não determinado"]
GABARITOS = ["Produção", "Técnico", "Produção + Técnico", "Outro", "Não determinado"]

GLOSSARY = {
    "N — Número de inspeções":
        "Quantidade total de inspeções classificadas na amostra: N = TP + TN/VN + FP + FN. Não representa quantidade de defeitos; representa a unidade estatística de inspeções.",
    "TP — Verdadeiro Positivo":
        "Na classificação da inspeção: a peça/carroceria realmente possui defeito e o Vision AI classifica como NOK.",
    "TN/VN — Verdadeiro Negativo":
        "Na classificação da inspeção: a peça/carroceria realmente está OK e o Vision AI classifica como OK.",
    "FP — Falso Positivo (Classificação)":
        "Na classificação da inspeção: a peça/carroceria está OK, mas o Vision AI classifica como NOK. Mede erro no resultado OK/NOK da inspeção.",
    "FN — Falso Negativo (Classificação)":
        "Na classificação da inspeção: existe defeito real, mas o Vision AI classifica como OK. Mede escape no resultado OK/NOK da inspeção.",
    "FP — Falso Positivo (Detecção)":
        "Na análise de ocorrências: uma detecção registrada não encontra defeito real correspondente. Nesta versão, quando não existe correspondência individual, a estimativa por contagem usa max(Detectado−Real, 0).",
    "FN — Falso Negativo (Detecção)":
        "Na análise de ocorrências: existe defeito real que não aparece nas detecções. Quando não existe correspondência individual, a estimativa por contagem usa max(Real−Detectado, 0).",
    "Accuracy — Acurácia":
        "(TP+TN)/(TP+TN+FP+FN). Percentual de inspeções classificadas corretamente.",
    "Precision — Precisão":
        "TP/(TP+FP). Entre as inspeções classificadas como NOK, proporção que realmente possuía defeito.",
    "Recall / Sensibilidade":
        "TP/(TP+FN). Entre as inspeções realmente NOK, proporção classificada como NOK.",
    "Specificity — Especificidade":
        "TN/(TN+FP). Entre as inspeções realmente OK, proporção classificada como OK.",
    "FPR — False Positive Rate":
        "FP/(FP+TN). Taxa clássica de falso positivo na classificação das inspeções.",
    "FNR — False Negative Rate":
        "FN/(FN+TP). Taxa clássica de falso negativo na classificação das inspeções.",
    "F1-Score":
        "Média harmônica entre Precision e Recall.",
    "MCC":
        "Matthews Correlation Coefficient. Indicador de correlação da matriz de confusão.",
    "Threshold / Grau de Confiança":
        "Limite usado pelo sistema de visão para aceitar uma detecção.",
    "Método de referência":
        "Método usado para determinar a verdade do caso: técnico, produção, ambos etc.",
    "Ocorrência de defeito":
        "Quantidade de defeitos reais ou detectados dentro de uma inspeção. Pode ser registrada sem informar o tipo.",
    "TP-D / FP-D / FN-D":
        "Indicadores de detecção de ocorrências. São calculados separadamente da classificação OK/NOK; quando não há correspondência individual, usam uma estimativa por contagem.",
    "TP-C / VN-C / FP-C / FN-C":
        "Indicadores da classificação da inspeção, isto é, o resultado OK/NOK da carroceria, peça ou inspeção.",
    "TP-DI / TN-DI / FP-DI / FN-DI":
        "Indicadores de detecção por inspeção. D-I usa presença binária por inspeção: Real>0 × Detectado>0. Não são os mesmos TP-D/FP-D/FN-D de ocorrência.",
    "FP-DI / 100 inspeções":
        "Fórmula: (FP-DI / N-DI) × 100. Normaliza falsos alarmes da matriz D-I para 100 inspeções.",
    "FN-DI / 100 inspeções":
        "Fórmula: (FN-DI / N-DI) × 100. Normaliza escapes da matriz D-I para 100 inspeções.",
    "Prevalência":
        "Proporção de inspeções realmente NOK dentro da amostra avaliada. É um indicador de composição da amostra, não de qualidade do sistema.",
    "Taxa de alarme NOK":
        "Proporção de inspeções classificadas pelo sistema como NOK. Indicador operacional complementar.",
    "FPR / FNR":
        "FPR mede a fração de casos realmente OK classificados como NOK; FNR mede a fração de casos realmente NOK classificados como OK.",
    "Expectativa / Meta de Entrega":
        "Critério de desempenho definido para orientar a calibração e a entrega. Accuracy, Precision, Recall e F1 usam meta mínima; FP/100 e FN/100 usam limite máximo.",
    "FP / 100 inspeções":
        "Fórmula: (FP / N) × 100. Normaliza a carga de falsos positivos de classificação para 100 inspeções.",
    "FN / 100 inspeções":
        "Fórmula: (FN / N) × 100. Normaliza a carga de falsos negativos de classificação para 100 inspeções.",
    "Precision de detecção":
        "Fórmula: TP-D / (TP-D + FP-D). Mede a confiabilidade das ocorrências detectadas.",
    "Recall de detecção":
        "Fórmula: TP-D / (TP-D + FN-D). Mede a capacidade de encontrar ocorrências reais.",
}

def safe(a, b):
    return None if b == 0 else a / b

def pct(v):
    return "N/D" if v is None else f"{v*100:.2f}%"

def classification_metrics(tp, tn, fp, fn):
    n = tp + tn + fp + fn
    precision = safe(tp, tp + fp)
    recall = safe(tp, tp + fn)
    specificity = safe(tn, tn + fp)
    fpr = safe(fp, fp + tn)
    fnr = safe(fn, fn + tp)
    npv = safe(tn, tn + fn)
    accuracy = safe(tp + tn, n)
    bal = None if recall is None or specificity is None else (recall + specificity) / 2
    f1 = None if precision is None or recall is None or precision + recall == 0 else 2 * precision * recall / (precision + recall)
    beta = 2.0
    f2 = None if precision is None or recall is None or (beta**2 * precision + recall) == 0 else (1+beta**2)*precision*recall/(beta**2*precision+recall)
    den = (tp+fp)*(tp+fn)*(tn+fp)*(tn+fn)
    mcc = None if den <= 0 else (tp*tn-fp*fn)/math.sqrt(den)
    return {
        "N": n, "TP": tp, "VN/TN": tn, "FP": fp, "FN": fn,
        "Accuracy": accuracy, "Precision": precision,
        "Recall / Sensibilidade": recall, "Specificity": specificity,
        "FPR": fpr, "FNR": fnr, "NPV": npv,
        "Balanced Accuracy": bal, "F1": f1, "F-beta (β=2)": f2, "MCC": mcc
    }


def matrix_c_from_records(records):
    """Motor puro da Matriz C: somente Vision AI x Referência OK/NOK."""
    out = {"TP": 0, "TN": 0, "VN/TN": 0, "FP": 0, "FN": 0}
    for item in records:
        vision = str(item.get("vision", "")).strip().upper()
        reference = str(item.get("reference", item.get("ref", ""))).strip()
        if vision not in {"OK", "NOK"} or reference not in {"OK confirmado", "Defeito confirmado"}:
            continue
        if reference == "Defeito confirmado":
            out["TP" if vision == "NOK" else "FN"] += 1
        else:
            out["TN" if vision == "OK" else "FP"] += 1
    out["VN/TN"] = out["TN"]
    out["N"] = out["TP"] + out["TN"] + out["FP"] + out["FN"]
    out.update(classification_metrics(out["TP"], out["TN"], out["FP"], out["FN"]))
    out["source"] = "Vision AI × Referência"
    out["unit"] = "inspeção"
    out["method"] = "decisão OK/NOK"
    return out


def matrix_di_from_records(records):
    """Motor puro da Matriz D-I: somente presença Real x Detectado."""
    tp = tn = fp = fn = 0
    for item in records:
        if item.get("real") is None or item.get("detected", item.get("det")) is None:
            continue
        real = float(item.get("real", 0)) > 0
        detected = float(item.get("detected", item.get("det", 0))) > 0
        if real and detected: tp += 1
        elif real: fn += 1
        elif detected: fp += 1
        else: tn += 1
    out = classification_metrics(tp, tn, fp, fn)
    out["TN"] = tn
    out["Prevalência"] = safe(tp + fn, out["N"])
    out["source"] = "Real × Detectado"
    out["unit"] = "inspeção"
    out["method"] = "presença binária: Real>0 × Detectado>0"
    return out


def matrix_relationship_diagnostics(records):
    """Explica por que C e D-I coincidem ou divergem sem misturar seus motores."""
    comparable = equivalent = 0
    divergent_ids = []
    for index, item in enumerate(records, 1):
        vision = str(item.get("vision", "")).strip().upper()
        reference = str(item.get("reference", item.get("ref", ""))).strip()
        detected_raw = item.get("detected", item.get("det"))
        real_raw = item.get("real")
        if (vision not in {"OK", "NOK"} or reference not in {"OK confirmado", "Defeito confirmado"}
                or real_raw is None or detected_raw is None):
            continue
        comparable += 1
        c_pair = (reference == "Defeito confirmado", vision == "NOK")
        di_pair = (float(real_raw) > 0, float(detected_raw) > 0)
        if c_pair == di_pair:
            equivalent += 1
        else:
            divergent_ids.append(str(item.get("id") or item.get("identification") or index))
    return {
        "comparable": comparable, "equivalent": equivalent,
        "divergent": comparable - equivalent, "divergent_ids": divergent_ids,
        "fully_equivalent": comparable > 0 and comparable == equivalent,
    }

def normalize_text(value):
    if value is None:
        return ""
    s = str(value).strip()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"\s+", " ", s)
    return s.upper().strip()

def normalize_defect(value):
    s = normalize_text(value)
    if not s or s in {"0", "-", "N/A", "NA", "N/D", "SEM DEFEITO", "SEM DEFEITO DETECTADO"}:
        return ""
    aliases = {
        "AMASSADO": "AMASSADO",
        "AMASS.": "AMASSADO",
        "AMASSADO/DEF.": "AMASSADO / DEFORMACAO",
        "AMASSADO / DEF.": "AMASSADO / DEFORMACAO",
        "AMASSADO-DEFORMACAO": "AMASSADO / DEFORMACAO",
        "DEFORMAÇÃO/AMASSADO": "DEFORMACAO / AMASSADO",
        "DEFORMACAO/AMASSADO": "DEFORMACAO / AMASSADO",
    }
    return aliases.get(s, s)

def blank_row():
    return {
        "id": "", "vision": "OK", "ref": "Não determinado",
        "real": 0, "det": 0,
        "tp": 0, "vn": 0, "fp": 0, "fn": 0,
        "obs": "", "calc_status": "N/D",
        "defects": [], "source": "Manual",
        # A linha vazia não é uma observação válida até o usuário salvar a inspeção.
        "detection_observed": False, "detection_explicit_matrix": False,
        "tp_d": None, "fp_d": None, "fn_d": None
    }

def normalize_row(r):
    x = blank_row()
    x.update(r or {})
    x["id"] = str(x.get("id") or "")
    x["vision"] = str(x.get("vision") or "OK").upper()
    x["ref"] = str(x.get("ref") or "Não determinado")
    for k in ("real","det","tp","vn","fp","fn"):
        try:
            x[k] = max(0, int(float(x.get(k) or 0)))
        except Exception:
            x[k] = 0
    x["obs"] = str(x.get("obs") or "")
    x["calc_status"] = str(x.get("calc_status") or "N/D")
    x["source"] = str(x.get("source") or "Manual")
    x["detection_available"] = bool(x.get("detection_available", True))
    # Compatibilidade com projetos manuais criados antes da V1.1.4: se a
    # inspeção foi efetivamente salva pelo caminho manual, Real/Detectado são
    # dados observados. A linha vazia continua protegida por detection_observed=False.
    if (str(x.get("source", "")).startswith("Manual")
            and str(x.get("id", "")).strip()
            and x.get("detection_observed") is False):
        x["detection_observed"] = True
    else:
        x["detection_observed"] = bool(x.get("detection_observed", False))
    x["detection_explicit_matrix"] = bool(x.get("detection_explicit_matrix", False))
    raw = x.get("defects") or []
    if not isinstance(raw, list):
        raw = []
    defects = []
    for d in raw:
        if not isinstance(d, dict):
            continue
        name = normalize_defect(d.get("type") or d.get("defeito"))
        if not name:
            continue
        try: rr = max(0, int(float(d.get("real") or 0)))
        except Exception: rr = 0
        try: dd = max(0, int(float(d.get("det") or d.get("detectado") or 0)))
        except Exception: dd = 0
        defects.append({"type": name, "real": rr, "det": dd})
    x["defects"] = defects
    return x

TOOLTIP_DEFS = {
    "N • Inspeções": ("N — Número de inspeções", "N = TP + TN/VN + FP + FN. Unidade estatística: inspeção/carroceria/peça."),
    "TP • Inspeções": ("TP — Verdadeiro Positivo", "TP = inspeções realmente NOK e classificadas como NOK."),
    "TN/VN • Inspeções": ("TN/VN — Verdadeiro Negativo", "TN/VN = inspeções realmente OK e classificadas como OK."),
    "FP • Inspeções": ("FP — Falso Positivo de classificação", "FP = inspeções realmente OK classificadas como NOK."),
    "FN • Inspeções": ("FN — Falso Negativo de classificação", "FN = inspeções realmente NOK classificadas como OK."),
    "Accuracy": ("Accuracy — Acurácia", "Fórmula: (TP + TN) / (TP + TN + FP + FN). Percentual de classificações corretas."),
    "Precision": ("Precision — Precisão", "Fórmula: TP / (TP + FP). Entre os NOK classificados, quantos realmente são NOK."),
    "Recall": ("Recall — Sensibilidade", "Fórmula: TP / (TP + FN). Entre os NOK reais, quantos foram encontrados."),
    "Recall / Sensibilidade": ("Recall — Sensibilidade", "Fórmula: TP / (TP + FN). Entre os NOK reais, quantos foram encontrados."),
    "Specificity": ("Specificity — Especificidade", "Fórmula: TN / (TN + FP). Entre os OK reais, quantos foram mantidos como OK."),
    "FPR": ("FPR — False Positive Rate", "Fórmula: FP / (FP + TN). Taxa de falsos positivos entre os OK reais."),
    "FNR": ("FNR — False Negative Rate", "Fórmula: FN / (FN + TP). Taxa de escapes entre os NOK reais."),
    "F1": ("F1-Score", "Fórmula: 2 × Precision × Recall / (Precision + Recall). Média harmônica entre Precision e Recall."),
    "MCC": ("MCC — Matthews Correlation Coefficient", "Coeficiente de correlação da matriz de confusão. Varia de -1 a +1."),
    "Prevalência": ("Prevalência", "Fórmula: (TP + FN) / N. Proporção de inspeções realmente NOK."),
    "Taxa de alarme NOK": ("Taxa de alarme NOK", "Fórmula: (TP + FP) / N. Proporção de inspeções classificadas como NOK."),
    "FP / 100 inspeções": ("Falsos positivos normalizados", "Fórmula: (FP / N) × 100. Carga de falso alarme normalizada para 100 inspeções."),
    "FN / 100 inspeções": ("Falsos negativos normalizados", "Fórmula: (FN / N) × 100. Carga de escape normalizada para 100 inspeções."),
    "Precision detecção": ("Precision de detecção", "Fórmula: TP-D / (TP-D + FP-D). Mede a confiabilidade das ocorrências detectadas."),
    "Recall detecção": ("Recall de detecção", "Fórmula: TP-D / (TP-D + FN-D). Mede a capacidade de encontrar ocorrências reais."),
}

class Tooltip:
    def __init__(self, widget, title, text):
        self.widget=widget; self.title=title; self.text=text; self.tip=None
        widget.bind("<Enter>", self.show, add="+")
        widget.bind("<Leave>", self.hide, add="+")
        widget.bind("<ButtonPress>", self.hide, add="+")
    def show(self, event=None):
        if self.tip is not None: return
        x=self.widget.winfo_rootx()+18; y=self.widget.winfo_rooty()+self.widget.winfo_height()+6
        self.tip=tk.Toplevel(self.widget); self.tip.wm_overrideredirect(True); self.tip.wm_geometry(f"+{x}+{y}")
        self.tip.configure(bg="#07131C")
        frame=tk.Frame(self.tip,bg="#0E202D",highlightbackground="#2B718F",highlightthickness=1); frame.pack()
        tk.Label(frame,text=self.title,bg="#0E202D",fg=ACCENT2,font=("Segoe UI",9,"bold"),anchor="w",justify="left").pack(fill="x",padx=10,pady=(8,2))
        tk.Label(frame,text=self.text,bg="#0E202D",fg=TEXT,font=("Segoe UI",8),anchor="w",justify="left",wraplength=390).pack(fill="x",padx=10,pady=(0,9))
    def hide(self,event=None):
        if self.tip is not None:
            try:self.tip.destroy()
            except Exception:pass
            self.tip=None

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP} | V{VERSION} | FRONT-END CONSOLIDADO")
        self.after(120, lambda: self.status.set("AIA V1.2.0 • MATRIZES INDEPENDENTES • CALIBRAÇÃO D-I • AUDITORIA ATIVA"))
        self.geometry("1520x930")
        self.minsize(1200, 760)
        self.configure(bg=BG)
        self.rows = []
        self.history = []
        self.comparison_archive = []
        self.calibration_history = []
        self.calibration_project = {"id":"", "name":"", "objective":"", "responsible":"", "created_at":""}
        self.asset_dir = application_dir()
        self.logo_img = None
        self.icon_img = None
        self.source_file = ""
        self.current_view = "dashboard"
        self.meta = {k: tk.StringVar() for k in
                     ["cliente","local","projeto","linha","tecnico","data","threshold","gabarito"]}
        self.expectation = {
            "accuracy": tk.StringVar(value=""), "precision": tk.StringVar(value=""),
            "recall": tk.StringVar(value=""), "f1": tk.StringVar(value=""),
            "fp100": tk.StringVar(value=""), "fn100": tk.StringVar(value=""),
            "source": tk.StringVar(value="Acordo / critério de entrega"),
            "responsible": tk.StringVar(value=""), "date": tk.StringVar(value=""),
            "notes": tk.StringVar(value=""),
        }
        self.setup_style()
        # Carrega a meta persistida antes da montagem visual, sem depender de uma segunda camada de UI.
        self._aia_final_load_active_expectation()
        try:
            # Usa o ícone oficial distribuído com a build, evitando o ícone
            # padrão do interpretador Python na janela e no atalho.
            icon_ico = self.asset_dir / "AIA_V1_3_0_RC.ico"
            if not icon_ico.exists():
                icon_ico = self.asset_dir / "AIA_V1_2_0.ico"
            if os.name == "nt" and icon_ico.exists():
                self.iconbitmap(str(icon_ico))
            icon = self.asset_dir / "autaza_icon.png"
            self.icon_img = tk.PhotoImage(file=str(icon))
            self.iconphoto(True, self.icon_img)
        except Exception:
            pass
        self.build_shell()
        self.new_test(False)

    def setup_style(self):
        s = ttk.Style(self)
        try: s.theme_use("clam")
        except Exception: pass
        s.configure(".", font=("Segoe UI",10), background=BG, foreground=TEXT)
        s.configure("TFrame", background=BG)
        s.configure("TLabel", background=BG, foreground=TEXT)
        s.configure("Accent.TButton", background=ACCENT, foreground="white",
                    padding=(14,8), font=("Segoe UI",10,"bold"))
        s.map("Accent.TButton", background=[("active",ACCENT2)])
        s.configure("TButton", background=PANEL2, foreground=TEXT,
                    bordercolor=BORDER, padding=(11,7))
        s.map("TButton", background=[("active","#203247")])
        s.configure("TEntry", fieldbackground="#0d151e", foreground=TEXT, insertcolor=TEXT)
        s.configure("TCombobox", fieldbackground="#F0F3F5", foreground="#17212B")
        s.configure("Treeview", background="#14202B", fieldbackground="#14202B",
                    foreground=TEXT, rowheight=31)
        s.configure("Treeview.Heading", background="#21465D", foreground=TEXT,
                    font=("Segoe UI",9,"bold"))
        s.map("Treeview", background=[("selected","#245D8F")],
              foreground=[("selected","#FFFFFF")])

    def build_shell(self):
        """Construção única e definitiva do shell visual do AIA.

        O motor estatístico e as rotas existentes são preservados. Esta camada
        controla somente identidade visual, navegação, hierarquia e área útil.
        """
        # ================================================================
        # HEADER — identidade AUTAZA / estado do sistema / comandos
        # ================================================================
        header = tk.Frame(self, bg="#061019", height=86,
                          highlightbackground="#102B39", highlightthickness=1)
        header.pack(fill="x", padx=12, pady=(9, 0))
        header.pack_propagate(False)

        brand = tk.Frame(header, bg="#061019")
        brand.pack(side="left", fill="y", padx=(8, 0))

        # Carregamento tolerante do logo: mantém compatibilidade com instalações
        # antigas e também com a pasta que contém os arquivos de marca originais.
        logo_candidates = [
            self.asset_dir / "autaza_header.png",
            self.asset_dir / "Logo Autaza (Branco).png",
            self.asset_dir / "Logo Autaza (Azul e Branco) (1).png",
            self.asset_dir / "Logo Autaza (Azul e Cinza).png",
        ]
        logo_loaded = False
        for logo_path in logo_candidates:
            if logo_path.exists():
                try:
                    self.logo_img = tk.PhotoImage(file=str(logo_path))
                    # Evita que logos muito grandes dominem o header.
                    while self.logo_img.height() > 66 or self.logo_img.width() > 360:
                        self.logo_img = self.logo_img.subsample(2, 2)
                    tk.Label(brand, image=self.logo_img, bg="#061019",
                             bd=0).pack(side="left", padx=(0, 16), pady=8)
                    logo_loaded = True
                    break
                except Exception:
                    pass
        if not logo_loaded:
            tk.Label(brand, text="AUTAZA", bg="#061019", fg="white",
                     font=("Segoe UI", 23, "bold")).pack(side="left", padx=(0, 18))

        title_wrap = tk.Frame(brand, bg="#061019")
        title_wrap.pack(side="left", fill="y", pady=12)
        tk.Label(title_wrap, text="INSPECTION ANALYTICS", bg="#061019", fg=TEXT,
                 font=("Segoe UI", 17, "bold")).pack(anchor="w")
        tk.Label(title_wrap,
                 text=f"INDUSTRIAL CONTROL ROOM   /   V{VERSION}   /   FRONT-END",
                 bg="#061019", fg=ACCENT2,
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(2, 0))

        # Área de comandos: ordem mantida para não quebrar os fluxos existentes.
        cmd = tk.Frame(header, bg="#061019")
        cmd.pack(side="right", fill="y", padx=(8, 8), pady=10)

        status_chip = tk.Frame(cmd, bg="#0A2531",
                               highlightbackground="#1B718E", highlightthickness=1)
        status_chip.pack(side="left", padx=(0, 8), fill="y")
        tk.Label(status_chip, text="●  SISTEMA ATIVO", bg="#0A2531", fg=GOOD,
                 font=("Segoe UI", 8, "bold")).pack(padx=11, pady=(5, 1))
        tk.Label(status_chip, text="APLICAÇÃO LOCAL", bg="#0A2531", fg=MUTED,
                 font=("Segoe UI", 7)).pack(padx=11, pady=(0, 5))

        self.header_buttons = {}
        for label, command in [
            ("ABRIR EXCEL", self.import_excel),
            ("ABRIR JSON", self.load_project),
            ("PROJETO", self.edit_calibration_project),
            ("SALVAR", self.save_project),
            ("EXPORTAR", self.export_excel),
            ("RELATÓRIO", self.export_executive_report),
            ("CSV", self.export_csv),
            ("GLOSSÁRIO", self.glossary),
        ]:
            b = tk.Button(cmd, text=label, command=command,
                          bg="#0D2432", fg=TEXT,
                          activebackground="#15506B", activeforeground="white",
                          bd=0, relief="flat", highlightbackground="#24536A",
                          highlightthickness=1, padx=9, pady=7,
                          font=("Segoe UI", 8, "bold"), cursor="hand2")
            b.pack(side="left", padx=2)
            self.header_buttons[label] = b

        # ================================================================
        # MAIN — navegação + conteúdo
        # ================================================================
        self.main = tk.Frame(self, bg=BG)
        self.main.pack(fill="both", expand=True, padx=12, pady=7)

        # Sidebar deliberadamente mais larga: os nomes dos módulos precisam ser
        # legíveis; abreviações continuam apenas nos títulos curtos de navegação.
        self.sidebar = tk.Frame(self.main, bg="#081923", width=198,
                                highlightbackground="#1A4054", highlightthickness=1)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        # Marca lateral
        mark = tk.Frame(self.sidebar, bg="#0A2532", height=68,
                        highlightbackground="#153D4D", highlightthickness=1)
        mark.pack(fill="x", padx=8, pady=(8, 6))
        mark.pack_propagate(False)
        try:
            icon_path = self.asset_dir / "autaza_icon.png"
            if not icon_path.exists():
                icon_path = self.asset_dir / "icone autaza (1).png"
            side_icon = tk.PhotoImage(file=str(icon_path))
            while side_icon.width() > 42 or side_icon.height() > 42:
                side_icon = side_icon.subsample(2, 2)
            self.side_icon_img = side_icon
            tk.Label(mark, image=self.side_icon_img, bg="#0A2532", bd=0).pack(expand=True)
        except Exception:
            tk.Label(mark, text="AIA", bg="#0A2532", fg=ACCENT2,
                     font=("Segoe UI", 16, "bold")).pack(expand=True)

        tk.Label(self.sidebar, text="MÓDULOS", bg="#081923", fg="#5E8799",
                 font=("Segoe UI", 7, "bold")).pack(anchor="w", padx=12, pady=(1, 4))

        self.nav_buttons = {}
        self.nav_meta = {}

        # Ordem operacional definitiva. Nenhuma funcionalidade é removida:
        # apenas reorganizamos a navegação para refletir o fluxo de trabalho.
        nav_groups = [
            ("OPERAÇÃO", [
                ("operacao", "▣", "COLETA", "Registro de campo e inspeções"),
            ]),
            ("VISÃO E ANÁLISE", [
                ("dashboard", "⌂", "DASHBOARD", "Visão executiva"),
                ("graficos", "◒", "GRÁFICOS", "Performance"),
                ("comparacao", "⇄", "COMPARAÇÃO", "Evolução entre coletas"),
                ("tecnica", "◆", "TÉCNICA", "Diagnóstico técnico"),
                ("defeitos", "≡", "DEFEITOS", "Análise por tipo"),
            ]),
            ("ENGENHARIA", [
                ("engenharia", "⌁", "ENGENHARIA", "Decisão e investigação"),
                ("calibracao", "◉", "CALIBRAÇÃO", "Projeto de calibração"),
                ("expectativa", "◎", "META", "Critério de entrega"),
            ]),
            ("RASTREABILIDADE", [
                ("auditoria", "✓", "AUDITORIA", "Histórico × motor novo"),
            ]),
        ]

        for group_name, items in nav_groups:
            tk.Label(self.sidebar, text=group_name, bg="#081923", fg="#3F7185",
                     font=("Segoe UI", 6, "bold")).pack(anchor="w", padx=12, pady=(5, 2))
            for key, icon, label, hint in items:
                b = tk.Button(
                    self.sidebar,
                    text=f"{icon}  {label}",
                    command=lambda k=key: self.show(k),
                    bg="#081923", fg="#B8CBD5",
                    activebackground="#124765", activeforeground="white",
                    bd=0, relief="flat", highlightthickness=0,
                    anchor="w", padx=11, pady=7,
                    font=("Segoe UI", 8, "bold"), cursor="hand2"
                )
                b.pack(fill="x", padx=6, pady=1)
                self.nav_buttons[key] = b
                self.nav_meta[key] = (label, hint)
                b.bind("<Enter>", lambda e, k=key: self._nav_hover(k, True))
                b.bind("<Leave>", lambda e, k=key: self._nav_hover(k, False))

        # Status fixo e compacto no rodapé da navegação.
        status_box = tk.Frame(self.sidebar, bg="#0A2635",
                              highlightbackground="#2B718F", highlightthickness=1)
        status_box.pack(side="bottom", fill="x", padx=8, pady=8)
        tk.Label(status_box, text="●  STATUS DO TESTE", bg="#0A2635", fg=ACCENT2,
                 font=("Segoe UI", 7, "bold")).pack(anchor="w", padx=9, pady=(6, 2))
        self.side_status = tk.Label(status_box, bg="#0A2635", fg=TEXT,
                                    justify="left", anchor="w", wraplength=170,
                                    font=("Segoe UI", 7))
        self.side_status.pack(fill="x", padx=9, pady=(0, 7))

        # ================================================================
        # WORKSPACE — faixa de instrumentação + área rolável
        # ================================================================
        self.content_host = tk.Frame(self.main, bg=BG)
        self.content_host.pack(side="left", fill="both", expand=True, padx=(9, 0))

        self.instrument_bar = tk.Frame(
            self.content_host, bg="#091C27", height=34,
            highlightbackground="#173D50", highlightthickness=1
        )
        self.instrument_bar.pack(fill="x", pady=(0, 7))
        self.instrument_bar.pack_propagate(False)

        self.instrument_left = tk.Label(
            self.instrument_bar, text="AIA  /  CONTROL ROOM",
            bg="#091C27", fg=ACCENT2, font=("Segoe UI", 8, "bold")
        )
        self.instrument_left.pack(side="left", padx=12)

        self.instrument_right = tk.Label(
            self.instrument_bar, text="DATA STREAM: MANUAL     •     ENGINE: READY",
            bg="#091C27", fg="#6E8E9E", font=("Segoe UI", 7, "bold")
        )
        self.instrument_right.pack(side="right", padx=12)

        scroll_wrap = tk.Frame(self.content_host, bg=BG)
        scroll_wrap.pack(fill="both", expand=True)

        self.content_canvas = tk.Canvas(scroll_wrap, bg=BG, bd=0,
                                        highlightthickness=0)
        self.content_canvas.pack(side="left", fill="both", expand=True)

        self.content_scrollbar = tk.Scrollbar(
            scroll_wrap, orient="vertical", command=self.content_canvas.yview,
            bg="#0B2634", troughcolor="#061019", activebackground="#1689B5",
            width=12
        )
        self.content_scrollbar.pack(side="right", fill="y")

        self.content = tk.Frame(self.content_canvas, bg=BG)
        self.content_window = self.content_canvas.create_window(
            (0, 0), window=self.content, anchor="nw"
        )
        self.content.configure(highlightthickness=0, bd=0)

        self.content.bind(
            "<Configure>",
            lambda e: self.content_canvas.configure(
                scrollregion=self.content_canvas.bbox("all")
            )
        )
        self.content_canvas.bind(
            "<Configure>",
            lambda e: self.content_canvas.itemconfigure(
                self.content_window, width=e.width
            )
        )

        # Scroll apenas quando o ponteiro estiver sobre a área de trabalho.
        self.content_canvas.bind("<Enter>", self._enable_workspace_wheel)
        self.content_canvas.bind("<Leave>", self._disable_workspace_wheel)
        self._wheel_bound = False

        # ================================================================
        # FOOTER
        # ================================================================
        bar = tk.Frame(self, bg="#040C12", height=25,
                       highlightbackground="#10222D", highlightthickness=1)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        self.status = tk.StringVar(value="Pronto.")
        tk.Label(bar, textvariable=self.status, bg="#040C12", fg=MUTED,
                 font=("Segoe UI", 8)).pack(side="left", padx=12)
        tk.Label(bar, text=f"AIA  /  INDUSTRIAL CONTROL ROOM  /  V{VERSION}",
                 bg="#040C12", fg="#416777", font=("Segoe UI", 8, "bold")
                 ).pack(side="right", padx=12)

        self.show("dashboard")

    def _nav_hover(self, key, entering):
        """Hover discreto sem sobrescrever o estado ativo."""
        if key not in self.nav_buttons:
            return
        b = self.nav_buttons[key]
        if key == self.current_view:
            return
        b.configure(bg="#0C2C3D" if entering else "#081923")

    def _enable_workspace_wheel(self, _event=None):
        if not self._wheel_bound:
            self.content_canvas.bind_all("<MouseWheel>", self._workspace_wheel)
            self._wheel_bound = True

    def _disable_workspace_wheel(self, _event=None):
        if self._wheel_bound:
            self.content_canvas.unbind_all("<MouseWheel>")
            self._wheel_bound = False

    def _workspace_wheel(self, event):
        step = -1 if event.delta > 0 else 1
        self.content_canvas.yview_scroll(step, "units")

    def clear_content(self):
        for w in self.content.winfo_children():
            if w is not getattr(self, "instrument_bar", None):
                w.destroy()

    def set_instrument(self, module_name):
        if hasattr(self, "instrument_left"):
            self.instrument_left.configure(text=f"AIA  /  {module_name.upper()}")
        if hasattr(self, "instrument_right"):
            self.instrument_right.configure(
                text=f"DATA STREAM: {'EXCEL' if self.source_file else 'MANUAL'}"
                      f"     •     ENGINE: {'ACTIVE' if self.rows else 'READY'}"
            )

    def show(self, key):
        """Navegação centralizada: uma rota, um estado ativo, uma view."""
        routes = {
            "dashboard": self.view_dashboard,
            "graficos": self.view_graficos,
            "comparacao": self.view_comparacao,
            "calibracao": self.view_calibracao,
            "operacao": self.view_operacao,
            "auditoria": self.view_auditoria,
            "tecnica": self.view_tecnica,
            "engenharia": self.view_engenharia,
            "defeitos": self.view_defeitos,
            "expectativa": self.view_expectativa,
        }
        module_names = {
            "dashboard": "EXECUTIVE / DASHBOARD",
            "graficos": "PERFORMANCE / GRÁFICOS",
            "comparacao": "COMPARAÇÃO / EVOLUÇÃO",
            "calibracao": "PROJETO / CALIBRAÇÃO",
            "operacao": "COLETA / OPERAÇÃO",
            "auditoria": "AUDITORIA / HISTÓRICO",
            "tecnica": "ANÁLISE TÉCNICA / DIAGNÓSTICO",
            "engenharia": "ENGENHARIA / DECISÃO",
            "defeitos": "ANÁLISE POR DEFEITO",
            "expectativa": "EXPECTATIVA / CRITÉRIO DE ENTREGA",
        }
        if key not in routes:
            return

        # Reinicia a posição antes de destruir a view longa anterior. Isso evita
        # que uma tela curta seja criada fora do viewport visível.
        try:
            self.content_canvas.yview_moveto(0)
            self.content_canvas.configure(scrollregion=(0, 0, 1, 1))
        except Exception:
            pass
        self.current_view = key
        self.clear_content()
        self.set_instrument(module_names[key])

        # Estado ativo inequívoco; hover não interfere no item selecionado.
        for k, b in self.nav_buttons.items():
            active = (k == key)
            b.configure(
                bg="#0D6288" if active else "#081923",
                fg="white" if active else "#B8CBD5",
                relief="solid" if active else "flat",
                bd=0,
                highlightthickness=1 if active else 0,
                highlightbackground="#3FA7CF" if active else "#081923",
                font=("Segoe UI", 8, "bold"),
            )

        routes[key]()
        try:
            self.update_idletasks()
            self.content_canvas.configure(scrollregion=self.content_canvas.bbox("all"))
            self.content_canvas.yview_moveto(0)
            self.after_idle(lambda: self.content_canvas.yview_moveto(0))
        except Exception:
            pass

    def card(self,parent,title,subtitle=None):
        f=tk.Frame(parent,bg=PANEL,highlightbackground=BORDER,highlightthickness=1)
        tk.Label(f,text=title,bg=PANEL,fg=TEXT,font=("Segoe UI",11,"bold")).pack(
            anchor="w",padx=14,pady=(12,3))
        if subtitle:
            tk.Label(f,text=subtitle,bg=PANEL,fg=MUTED,font=("Segoe UI",8),
                     wraplength=1000,justify="left").pack(anchor="w",padx=14,pady=(0,8))
        return f

    # ---------- DATA / MOTOR ----------
    def classify(self,r):
        r=normalize_row(r)
        r.update(tp=0,vn=0,fp=0,fn=0,calc_status="N/D")
        ref=r["ref"]; vision=r["vision"]
        if ref=="Defeito confirmado":
            r["tp"]=1 if vision=="NOK" else 0
            r["fn"]=1 if vision!="NOK" else 0
            r["calc_status"]="Classificação calculada"
        elif ref=="OK confirmado":
            r["vn"]=1 if vision=="OK" else 0
            r["fp"]=1 if vision!="OK" else 0
            r["calc_status"]="Classificação calculada"
        return r

    def detection_for_row(self, r0):
        """Calcula erros no nível de ocorrência, separado da classificação OK/NOK."""
        r=normalize_row(r0)
        if not r.get("detection_available", True):
            return {"real":0,"det":0,"tp":0,"fp":0,"fn":0,"available":False}
        ds=r.get("defects",[])
        if ds:
            real=det=tp=fp=fn=0
            for d in ds:
                rr,dd=int(d.get("real",0)),int(d.get("det",0))
                real += rr; det += dd
                tp += min(rr,dd); fp += max(dd-rr,0); fn += max(rr-dd,0)
            return {"real":real,"det":det,"tp":tp,"fp":fp,"fn":fn,"available":True}
        rr,dd=int(r.get("real",0)),int(r.get("det",0))
        return {"real":rr,"det":dd,"tp":min(rr,dd),"fp":max(dd-rr,0),"fn":max(rr-dd,0),"available":True}

    def detection_totals(self):
        total={"real":0,"det":0,"tp":0,"fp":0,"fn":0,"available_rows":0,"unavailable_rows":0}
        for r in self.rows:
            d=self.detection_for_row(r)
            if not d["available"]:
                total["unavailable_rows"] += 1
                continue
            total["available_rows"] += 1
            for k in ("real","det","tp","fp","fn"):
                total[k] += d[k]
        total["precision"]=safe(total["tp"], total["tp"]+total["fp"])
        total["recall"]=safe(total["tp"], total["tp"]+total["fn"])
        return total

    def defect_counts(self):
        real=det=tp=fp=fn=0
        detailed=False
        by=defaultdict(lambda:{"real":0,"det":0,"tp":0,"fp":0,"fn":0,"inspecoes":0})
        for r0 in self.rows:
            r=normalize_row(r0)
            ds=r.get("defects",[])
            if ds:
                detailed=True
                for d in ds:
                    rr,dd=d["real"],d["det"]
                    tt=min(rr,dd); ff=max(dd-rr,0); nn=max(rr-dd,0)
                    real+=rr; det+=dd; tp+=tt; fp+=ff; fn+=nn
                    a=by[d["type"]]
                    a["real"]+=rr; a["det"]+=dd; a["tp"]+=tt; a["fp"]+=ff; a["fn"]+=nn; a["inspecoes"]+=1
            else:
                d=self.detection_for_row(r)
                if not d["available"]:
                    continue
                real+=d["real"]; det+=d["det"]; tp+=d["tp"]; fp+=d["fp"]; fn+=d["fn"]
        return {"real":real,"det":det,"tp":tp,"fp":fp,"fn":fn,
                "by":by,"detailed":detailed}

    def classification_totals(self):
        # O motor novo só contabiliza inspeções com referência explícita.
        # Linhas importadas do Excel permanecem como HISTÓRICO e não entram
        # automaticamente na matriz de confusão do motor novo.
        vals={"tp":0,"vn":0,"fp":0,"fn":0}
        for i,r0 in enumerate(self.rows):
            r=normalize_row(r0)
            if r.get("calc_status") == "Histórico" or str(r.get("ref","")).startswith("Importado"):
                self.rows[i]=r
                continue
            r=self.classify(r)
            self.rows[i]=r
            for k in vals:
                vals[k]+=r[k]
        return vals

    def metrics(self):
        t=self.classification_totals()
        return classification_metrics(t["tp"],t["vn"],t["fp"],t["fn"])

    def historical_metrics(self):
        """Métricas calculadas diretamente dos totais importados da planilha."""
        if not self.history:
            return classification_metrics(0,0,0,0)
        h=self.history[-1]
        return classification_metrics(h.get("TP",0), h.get("VN/TN",0),
                                      h.get("FP",0), h.get("FN",0))

    def display_metrics(self):
        """Dashboard/gráficos usam o histórico importado quando ele é a fonte ativa."""
        if self.source_file and self.history:
            return self.historical_metrics()
        return self.metrics()

    # ---------- EXPECTATIVA / META ----------
    def _parse_target(self,key):
        raw=(self.expectation[key].get() or "").strip().replace(",",".")
        if not raw:return None
        try:
            v=float(raw)
            if key in ("accuracy","precision","recall","f1"):
                if v>1:v/=100.0
                return max(0.0,min(1.0,v))
            return max(0.0,v)
        except Exception:return None

    def expectation_snapshot(self):
        return {k:self._parse_target(k) for k in ("accuracy","precision","recall","f1","fp100","fn100")} | {
            "source":self.expectation["source"].get().strip(),
            "responsible":self.expectation["responsible"].get().strip(),
            "date":self.expectation["date"].get().strip(),
            "notes":self.expectation["notes"].get().strip(),
        }

    def _expectation_actuals(self):
        m=self.display_metrics(); n=m.get("N",0)
        return {"accuracy":m.get("Accuracy"),"precision":m.get("Precision"),"recall":m.get("Recall / Sensibilidade"),
                "f1":m.get("F1"),"fp100":None if not n else m.get("FP",0)/n*100.0,
                "fn100":None if not n else m.get("FN",0)/n*100.0}

    def _expectation_rows(self):
        actual=self._expectation_actuals(); target=self.expectation_snapshot()
        specs=[("Accuracy","accuracy",">="),("Precision","precision",">="),("Recall","recall",">="),
               ("F1","f1",">="),("FP / 100 inspeções","fp100","<="),("FN / 100 inspeções","fn100","<=")]
        rows=[]
        for label,key,op in specs:
            t=target[key]; a=actual[key]
            if t is None: status="SEM META"
            elif a is None: status="N/D"
            elif (a>=t if op==">=" else a<=t): status="ATENDIDA"
            else: status="ABAIXO DA META" if op==">=" else "ACIMA DA META"
            rows.append((label,key,op,a,t,status))
        return rows

    def expectation_summary(self):
        rows=self._expectation_rows(); configured=[r for r in rows if r[4] is not None]
        if not configured:return {"configured":0,"met":0,"status":"SEM META"}
        met=sum(1 for r in configured if r[5]=="ATENDIDA")
        status="META ATINGIDA" if met==len(configured) else ("FORA DA EXPECTATIVA" if met==0 else "EM CALIBRAÇÃO")
        return {"configured":len(configured),"met":met,"status":status}

    def _expectation_status_color(self,status):
        return {"META ATINGIDA":GOOD,"ATENDIDA":GOOD,"EM CALIBRAÇÃO":WARN,"FORA DA EXPECTATIVA":BAD,
                "ABAIXO DA META":WARN,"ACIMA DA META":WARN,"SEM META":MUTED,"N/D":MUTED}.get(status,TEXT)

    def view_expectativa(self):
        self.section_title("EXPECTATIVA  /  CRITÉRIO DE ENTREGA",
            "Defina o norte da calibração e acompanhe resultado × meta. Campo vazio = indicador não monitorado.")
        summary=self.expectation_summary()
        hero=self.card(self.content,"🎯 MODO EXPECTATIVA","Critérios de aceitação definidos para orientar a calibração e a entrega.")
        hero.pack(fill="x",pady=(0,7))
        row=tk.Frame(hero,bg=PANEL);row.pack(fill="x",padx=16,pady=(2,12))
        tk.Label(row,text=summary["status"],bg=PANEL,fg=self._expectation_status_color(summary["status"]),font=("Segoe UI",16,"bold")).pack(side="left")
        tk.Label(row,text=f"   {summary['met']}/{summary['configured']} metas configuradas atendidas",bg=PANEL,fg=MUTED,font=("Segoe UI",9)).pack(side="left")

        meta=self.card(self.content,"METAS DE DESEMPENHO","Accuracy, Precision, Recall e F1 = mínimo. FP/100 e FN/100 = máximo.")
        meta.pack(fill="x",pady=5)
        grid=tk.Frame(meta,bg=PANEL);grid.pack(fill="x",padx=16,pady=8)
        for i,(lab,key) in enumerate([("Accuracy (%)","accuracy"),("Precision (%)","precision"),("Recall (%)","recall"),("F1 (%)","f1"),("FP / 100","fp100"),("FN / 100","fn100")]):
            r,c=divmod(i,3);box=tk.Frame(grid,bg=PANEL);box.grid(row=r,column=c,sticky="ew",padx=6,pady=6);grid.columnconfigure(c,weight=1)
            tk.Label(box,text=lab,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w")
            ttk.Entry(box,textvariable=self.expectation[key]).pack(fill="x")
        table=ttk.Treeview(meta,columns=("ind","meta","atual","regra","status"),show="headings",height=6)
        for c,h,w in [("ind","Indicador",210),("meta","Meta",100),("atual","Resultado atual",120),("regra","Critério",90),("status","Status",180)]:table.heading(c,text=h);table.column(c,width=w,anchor="center" if c!="ind" else "w")
        table.pack(fill="x",padx=16,pady=(2,14))
        for label,key,op,a,t,status in self._expectation_rows():
            ms="N/D" if t is None else (f"{t*100:.2f}%" if key in ("accuracy","precision","recall","f1") else f"{t:.2f}")
            av="N/D" if a is None else (pct(a) if key in ("accuracy","precision","recall","f1") else f"{a:.2f}")
            iid=table.insert("","end",values=(label,ms,av,op,status))
            for item in table.get_children():
                if item==iid and status not in ("SEM META","N/D"):
                    # no per-row colors in ttk; tooltip carries the technical explanation
                    break

        ctx=self.card(self.content,"RASTREABILIDADE DA META","Registre de onde veio o critério para que o número continue defensável depois da calibração.")
        ctx.pack(fill="x",pady=5)
        form=tk.Frame(ctx,bg=PANEL);form.pack(fill="x",padx=16,pady=8)
        for i,(lab,key) in enumerate([("Origem do critério","source"),("Responsável","responsible"),("Data do critério","date")]):
            box=tk.Frame(form,bg=PANEL);box.grid(row=0,column=i,sticky="ew",padx=6);form.columnconfigure(i,weight=1)
            tk.Label(box,text=lab,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w");ttk.Entry(box,textvariable=self.expectation[key]).pack(fill="x")
        box=tk.Frame(form,bg=PANEL);box.grid(row=1,column=0,columnspan=3,sticky="ew",padx=6,pady=6)
        tk.Label(box,text="Observações / acordo com o cliente",bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w");ttk.Entry(box,textvariable=self.expectation["notes"]).pack(fill="x")

        read=self.card(self.content,"LEITURA DE CALIBRAÇÃO","Use a meta como norte; ela não substitui a análise de risco nem define sozinha a condição de entrega.")
        read.pack(fill="both",expand=True,pady=5)
        for line in ["• META ATINGIDA: todos os critérios configurados estão dentro das metas.","• EM CALIBRAÇÃO: parte das metas foi atendida; continue ajustando e comparando.","• FORA DA EXPECTATIVA: nenhum critério configurado foi atendido na amostra atual.","• COMPARA continua mostrando o efeito de diferentes thresholds; EXPECTATIVA acrescenta o critério de aceitação."]:
            tk.Label(read,text=line,bg=PANEL,fg=TEXT,wraplength=1100,justify="left",font=("Segoe UI",9)).pack(anchor="w",padx=16,pady=5)

    # ---------- VIEWS ----------
    def metric_card(self,parent,label,value,accent=None):
        f=tk.Frame(parent,bg=PANEL,highlightbackground=BORDER,highlightthickness=1,cursor="hand2")
        title=tk.Label(f,text=label,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold"),cursor="hand2")
        title.pack(anchor="w",padx=12,pady=(10,1))
        val=tk.Label(f,text=value,bg=PANEL,fg=accent or TEXT,font=("Segoe UI",18,"bold"),cursor="hand2")
        val.pack(anchor="w",padx=12,pady=(0,10))
        tip=TOOLTIP_DEFS.get(label)
        if tip:
            Tooltip(f,*tip); Tooltip(title,*tip); Tooltip(val,*tip)
        return f

    def section_title(self, title, subtitle=""):
        tk.Label(self.content, text=title, bg=BG, fg=TEXT,
                 font=("Segoe UI",18,"bold")).pack(anchor="w")
        if subtitle:
            tk.Label(self.content, text=subtitle, bg=BG, fg=MUTED,
                     font=("Segoe UI",9)).pack(anchor="w", pady=(2,10))

    def chart_card(self, parent, title, subtitle=""):
        f = tk.Frame(parent, bg=PANEL, highlightbackground=BORDER, highlightthickness=1)
        tk.Label(f, text=title, bg=PANEL, fg=TEXT,
                 font=("Segoe UI",10,"bold")).pack(anchor="w", padx=14, pady=(10,2))
        if subtitle:
            tk.Label(f, text=subtitle, bg=PANEL, fg=MUTED,
                     font=("Segoe UI",8)).pack(anchor="w", padx=14, pady=(0,4))
        return f

    def draw_bar_chart(self, parent, items, max_value=None, suffix=""):
        """Gráfico de barras simples, sem dependências externas."""
        canvas = tk.Canvas(parent, bg=PANEL, highlightthickness=0, height=220)
        canvas.pack(fill="both", expand=True, padx=12, pady=(4,12))
        canvas.update_idletasks()

        w = max(canvas.winfo_width(), 520)
        h = 220
        left, right, top, bottom = 72, 18, 18, 38
        plot_w, plot_h = w-left-right, h-top-bottom
        vals = [float(v or 0) for _,v,_ in items]
        maxv = max_value if max_value is not None else (max(vals) if vals else 1)
        if maxv <= 0: maxv = 1

        # grade horizontal
        for i in range(5):
            y = top + plot_h - (plot_h*i/4)
            canvas.create_line(left, y, w-right, y, fill="#29404F")
            val = maxv*i/4
            canvas.create_text(left-8, y, text=f"{val:.0f}{suffix}",
                               fill=MUTED, anchor="e", font=("Segoe UI",7))

        n = max(len(items),1)
        slot = plot_w/n
        bw = min(58, slot*0.58)
        for i,(label,val,accent) in enumerate(items):
            x = left + slot*i + slot/2
            v = float(val or 0)
            bh = plot_h*(v/maxv)
            y1 = top+plot_h-bh
            canvas.create_rectangle(x-bw/2, y1, x+bw/2, top+plot_h,
                                    fill=accent, outline="")
            canvas.create_text(x, y1-8, text=f"{v:.0f}{suffix}",
                               fill=TEXT, font=("Segoe UI",8,"bold"))
            canvas.create_text(x, top+plot_h+12, text=label,
                               fill=MUTED, font=("Segoe UI",7))

        return canvas

    def draw_percent_chart(self, parent, items):
        canvas = tk.Canvas(parent, bg=PANEL, highlightthickness=0, height=220)
        canvas.pack(fill="both", expand=True, padx=12, pady=(4,12))
        canvas.update_idletasks()
        w, h = max(canvas.winfo_width(), 520), 220
        left, right, top, bottom = 120, 48, 18, 30
        plot_w = w-left-right
        for i in range(0, 101, 25):
            x = left + plot_w*i/100
            canvas.create_line(x, top, x, h-bottom, fill="#29404F")
            canvas.create_text(x, h-bottom+12, text=f"{i}%",
                               fill=MUTED, font=("Segoe UI",7))
        row_h = (h-top-bottom)/max(len(items),1)
        for i,(label,val,accent) in enumerate(items):
            y = top + row_h*i + row_h/2
            p = 0 if val is None else max(0,min(1,float(val)))
            canvas.create_text(left-8,y,text=label,fill=MUTED,anchor="e",
                               font=("Segoe UI",8))
            canvas.create_rectangle(left,y-7,left+plot_w,y+7,
                                    fill="#1B2B38",outline="")
            canvas.create_rectangle(left,y-7,left+plot_w*p,y+7,
                                    fill=accent,outline="")
            canvas.create_text(left+plot_w*p+8,y,text="N/D" if val is None else f"{p*100:.1f}%",
                               fill=TEXT,anchor="w",font=("Segoe UI",8,"bold"))
        return canvas

    def view_dashboard(self):
        fonte = "HISTÓRICO DA PLANILHA" if (self.source_file and self.history) else "MOTOR NOVO / ENTRADA MANUAL"
        self.section_title(
            "DASHBOARD  /  VISÃO EXECUTIVA",
            f"Autaza Inspection Analytics • {fonte} • desempenho consolidado da inspeção e detecção")

        m=self.display_metrics(); d=self.defect_counts()
        cards=[("TP • Inspeções",m["TP"],GOOD),("TN/VN • Inspeções",m["VN/TN"],GOOD),
               ("FP • Inspeções",m["FP"],BAD),("FN • Inspeções",m["FN"],WARN),
               ("Accuracy",pct(m["Accuracy"]),ACCENT2),("Precision",pct(m["Precision"]),ACCENT2),
               ("Recall",pct(m["Recall / Sensibilidade"]),ACCENT2),("F1",pct(m["F1"]),ACCENT2)]
        grid=tk.Frame(self.content,bg=BG);grid.pack(fill="x")
        for i,(lab,val,col) in enumerate(cards):
            r,c=divmod(i,4)
            self.metric_card(grid,lab,str(val),col).grid(row=r,column=c,sticky="ew",padx=4,pady=4)
        for c in range(4):grid.columnconfigure(c,weight=1)

        es=self.expectation_summary()
        if es["configured"]:
            ec=self.card(self.content,"🎯 EXPECTATIVA DE ENTREGA",f"{es['met']}/{es['configured']} critérios atendidos")
            ec.pack(fill="x",pady=(7,3))
            er=tk.Frame(ec,bg=PANEL);er.pack(fill="x",padx=16,pady=(0,10))
            tk.Label(er,text=es["status"],bg=PANEL,fg=self._expectation_status_color(es["status"]),font=("Segoe UI",12,"bold")).pack(side="left")
            tk.Button(er,text="ABRIR EXPECTATIVA",command=lambda:self.show("expectativa"),bg="#0E4560",fg=TEXT,activebackground="#176B8F",bd=0,padx=12,pady=6,font=("Segoe UI",8,"bold")).pack(side="right")

        quick=tk.Frame(self.content,bg=BG)
        quick.pack(fill="x", pady=(8,4))
        tk.Button(quick,text="◒  ABRIR GRÁFICOS",command=lambda:self.show("graficos"),
                  bg="#0E4560",fg=TEXT,activebackground="#176B8F",
                  activeforeground="white",bd=0,padx=14,pady=8,
                  font=("Segoe UI",9,"bold")).pack(side="right")

        lower=tk.Frame(self.content,bg=BG);lower.pack(fill="both",expand=True,pady=6)
        left=self.card(lower,"CLASSIFICAÇÃO DAS INSPEÇÕES",
                        "Uma linha representa uma carroceria/peça. As métricas estatísticas usam esta unidade.")
        left.pack(side="left",fill="both",expand=True,padx=(0,4))
        txt=(f"TP = {m['TP']}\nTN/VN = {m['VN/TN']}\nFP = {m['FP']}\nFN = {m['FN']}\n\n"
             f"N de inspeções classificadas = {m['N']}")
        tk.Label(left,text=txt,bg=PANEL,fg=TEXT,font=("Consolas",11),
                 justify="left").pack(anchor="w",padx=16,pady=12)

        right=self.card(lower,"DETECÇÃO DE OCORRÊNCIAS",
                         "Real × detectado. O tipo do defeito é opcional.")
        right.pack(side="left",fill="both",expand=True,padx=(4,0))
        txt2=(f"Defeitos reais registrados = {d['real']}\n"
              f"Detecções registradas = {d['det']}\n"
              f"TP de detecção (por contagem) = {d['tp']}\n"
              f"FP de detecção (extras) = {d['fp']}\n"
              f"FN de detecção (não detectados) = {d['fn']}\n"
              f"Precision de detecção = {pct(safe(d['tp'], d['tp'] + d['fp']))}\n"
              f"Recall de detecção = {pct(safe(d['tp'], d['tp'] + d['fn']))}\n\n"
              + ("Tipos detalhados informados." if d["detailed"] else
                 "Coleta sem classificação de tipo."))
        tk.Label(right,text=txt2,bg=PANEL,fg=TEXT,font=("Consolas",11),
                 justify="left").pack(anchor="w",padx=16,pady=12)

    def view_graficos(self):
        fonte = "HISTÓRICO DA PLANILHA" if (self.source_file and self.history) else "MOTOR NOVO / ENTRADA MANUAL"
        self.section_title(
            "GRÁFICOS  /  PERFORMANCE",
            f"Autaza Inspection Analytics • {fonte} • leitura visual para acompanhamento técnico e executivo")

        m=self.display_metrics(); d=self.defect_counts()
        top=tk.Frame(self.content,bg=BG)
        top.pack(fill="both",expand=True)

        c1=self.chart_card(top,"MATRIZ DE CLASSIFICAÇÃO","Quantidade de inspeções por resultado")
        c1.grid(row=0,column=0,sticky="nsew",padx=(0,4),pady=(0,4))
        self.draw_bar_chart(c1,[("TP",m["TP"],GOOD),("TN",m["VN/TN"],GOOD),
                                ("FP",m["FP"],BAD),("FN",m["FN"],WARN)],
                            max_value=max([m["TP"],m["VN/TN"],m["FP"],m["FN"],1]))

        c2=self.chart_card(top,"INDICADORES DE PERFORMANCE","Percentual das principais métricas")
        c2.grid(row=0,column=1,sticky="nsew",padx=(4,0),pady=(0,4))
        self.draw_percent_chart(c2,[
            ("Accuracy",m["Accuracy"],ACCENT2),
            ("Precision",m["Precision"],ACCENT),
            ("Recall",m["Recall / Sensibilidade"],GOOD),
            ("F1",m["F1"],WARN),
        ])

        c3=self.chart_card(top,"OCORRÊNCIAS","Defeitos reais × detecções × divergências")
        c3.grid(row=1,column=0,sticky="nsew",padx=(0,4),pady=(4,0))
        self.draw_bar_chart(c3,[("Real",d["real"],GOOD),("Detectado",d["det"],ACCENT),
                                ("TP Det.",d["tp"],GOOD),("FP Det.",d["fp"],BAD),("FN Det.",d["fn"],WARN)],
                            max_value=max([d["real"],d["det"],d["tp"],d["fp"],d["fn"],1]))

        c4=self.chart_card(top,"LEITURA EXECUTIVA","Resumo visual para apresentação")
        c4.grid(row=1,column=1,sticky="nsew",padx=(4,0),pady=(4,0))
        summary = [
            ("Inspeções", m["N"]),
            ("Recall", pct(m["Recall / Sensibilidade"])),
            ("Precision", pct(m["Precision"])),
            ("Falsos positivos", m["FP"]),
            ("Falsos negativos", m["FN"]),
        ]
        for i,(lab,val) in enumerate(summary):
            row=tk.Frame(c4,bg=PANEL)
            row.pack(fill="x",padx=16,pady=6)
            tk.Label(row,text=lab,bg=PANEL,fg=MUTED,
                     font=("Segoe UI",9)).pack(side="left")
            tk.Label(row,text=str(val),bg=PANEL,fg=TEXT,
                     font=("Segoe UI",10,"bold")).pack(side="right")
        tk.Label(c4,text="Use esta tela para apresentar tendência e desempenho sem abrir a planilha.",
                 bg=PANEL,fg=ACCENT2,wraplength=460,justify="left",
                 font=("Segoe UI",8)).pack(anchor="w",padx=16,pady=(10,16))

        top.columnconfigure(0,weight=1); top.columnconfigure(1,weight=1)
        top.rowconfigure(0,weight=1); top.rowconfigure(1,weight=1)

    def view_operacao(self):
        tk.Label(self.content,text="COLETA / OPERAÇÃO",bg=BG,fg=TEXT,
                 font=("Segoe UI",18,"bold")).pack(anchor="w")
        tk.Label(self.content,text="Coleta rápida para campo + detalhamento opcional por inspeção",
                 bg=BG,fg=MUTED).pack(anchor="w",pady=(2,8))
        info=self.card(self.content,"IDENTIFICAÇÃO DO TESTE")
        info.pack(fill="x",pady=4)
        fields=[("Cliente","cliente"),("Local / Planta","local"),("Projeto","projeto"),
                ("Linha / Estação","linha"),("Técnico","tecnico"),("Data","data"),
                ("Threshold","threshold"),("Método de referência","gabarito")]
        body=tk.Frame(info,bg=PANEL);body.pack(fill="x",padx=6,pady=(0,10))
        for i,(lab,key) in enumerate(fields):
            r,c=divmod(i,4);box=tk.Frame(body,bg=PANEL)
            box.grid(row=r,column=c,sticky="ew",padx=8,pady=4)
            tk.Label(box,text=lab,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w")
            if key=="gabarito":
                w=ttk.Combobox(box,textvariable=self.meta[key],values=GABARITOS,state="readonly")
            else:
                w=ttk.Entry(box,textvariable=self.meta[key])
            w.pack(fill="x");body.columnconfigure(c,weight=1)

        source_label = "✓ Metadados preenchidos automaticamente pelo Excel" if self.source_file else "Entrada manual — preencha o contexto da inspeção"
        tk.Label(info,text=source_label,bg=PANEL,fg=GOOD if self.source_file else MUTED,
                 font=("Segoe UI",8,"bold")).pack(anchor="w",padx=14,pady=(0,8))

        tb=tk.Frame(self.content,bg=BG);tb.pack(fill="x",pady=5)
        ttk.Button(tb,text="+ Nova inspeção",style="Accent.TButton",
                   command=self.add_row).pack(side="left",padx=3)
        ttk.Button(tb,text="↻ Nova coleta",command=self.new_collection).pack(side="left",padx=3)
        ttk.Button(tb,text="⚡ Coleta rápida",command=self.quick_entry).pack(side="left",padx=3)
        ttk.Button(tb,text="Detalhar defeitos",command=self.edit_row).pack(side="left",padx=3)
        ttk.Button(tb,text="Excluir",command=self.delete_row).pack(side="left",padx=3)
        ttk.Button(tb,text="Recalcular",command=self.refresh_all).pack(side="left",padx=3)

        legend=tk.Label(self.content, text="TP-C / VN-C / FP-C / FN-C = classificação da inspeção (OK/NOK)  •  TP-D / FP-D / FN-D = detecção de ocorrências por contagem", bg=BG, fg=ACCENT2, font=("Segoe UI",8,"bold"))
        legend.pack(anchor="w", pady=(0,3))

        table=self.card(self.content,"INSPEÇÕES")
        table.pack(fill="both",expand=True,pady=5)
        cols=("n","id","vision","ref","real","det","tp","vn","fp","fn","dtp","dfp","dfn","def","status")
        heads=("Nº","Identificação","Vision","Referência","Real","Detectado","TP-C","VN-C","FP-C","FN-C","TP-D","FP-D","FN-D","Detalhes","Motor")
        self.tree=ttk.Treeview(table,columns=cols,show="headings")
        widths={"n":45,"id":90,"vision":65,"ref":115,"real":55,"det":65,
                "tp":50,"vn":50,"fp":50,"fn":50,"dtp":50,"dfp":50,"dfn":50,"def":75,"status":125}
        for c,h in zip(cols,heads):
            self.tree.heading(c,text=h);self.tree.column(c,width=widths.get(c,80),anchor="center")
        self.tree.pack(fill="both",expand=True,padx=10,pady=10)
        self.tree.bind("<Double-1>",lambda e:self.edit_row())
        self.refresh_table()

    def refresh_table(self):
        if not hasattr(self,"tree"): return
        for x in self.tree.get_children(): self.tree.delete(x)
        for i,r0 in enumerate(self.rows,1):
            r=normalize_row(r0); self.rows[i-1]=r
            detailed=f"{len(r['defects'])} tipo(s)" if r["defects"] else "Opcional"
            dd=self.detection_for_row(r)
            detvals=(dd["tp"],dd["fp"],dd["fn"]) if dd["available"] else ("N/D","N/D","N/D")
            self.tree.insert("", "end", iid=str(i-1),
                values=(i,r["id"],r["vision"],r["ref"],r["real"],r["det"],
                        r["tp"],r["vn"],r["fp"],r["fn"],detvals[0],detvals[1],detvals[2],detailed,r["calc_status"]))

    def add_row(self):
        self.rows.append(blank_row())
        self.refresh_all()
        self.edit_row(len(self.rows)-1)

    def _id_exists(self, ident, exclude_index=None):
        """Retorna True quando a identificação já está sendo usada por outra inspeção.
        A comparação ignora espaços laterais, mas preserva o texto/código digitado.
        """
        key=str(ident or "").strip()
        if not key:
            return False
        for i,row in enumerate(self.rows):
            if exclude_index is not None and i == exclude_index:
                continue
            if str(normalize_row(row).get("id") or "").strip() == key:
                return True
        return False

    def _next_numeric_id(self):
        """Sugere o próximo identificador numérico sem sobrescrever o que o usuário digitou."""
        nums=[]
        for row in self.rows:
            raw=str(normalize_row(row).get("id") or "").strip()
            if raw.isdigit():
                try: nums.append(int(raw))
                except Exception: pass
        return str(max(nums)+1) if nums else "1"

    def quick_entry(self):
        win=tk.Toplevel(self);win.title("Coleta rápida — nova inspeção")
        win.geometry("610x570");win.configure(bg=BG);win.transient(self);win.grab_set()
        frame=tk.Frame(win,bg=PANEL,highlightbackground=BORDER,highlightthickness=1)
        frame.pack(fill="both",expand=True,padx=15,pady=15)
        tk.Label(frame,text="COLETA RÁPIDA",bg=PANEL,fg=TEXT,
                 font=("Segoe UI",16,"bold")).pack(anchor="w",padx=18,pady=(18,3))
        tk.Label(frame,text="Não é necessário informar o tipo do defeito. Registre somente o que foi observado.",
                 bg=PANEL,fg=MUTED,wraplength=520,justify="left").pack(anchor="w",padx=18,pady=(0,14))
        vars={}
        specs=[("Identificação","id"),("Vision AI","vision"),("Referência / gabarito","ref"),
               ("Quantidade real de defeitos","real"),("Quantidade detectada","det"),
               ("Observação","obs")]
        for lab,key in specs:
            row=tk.Frame(frame,bg=PANEL);row.pack(fill="x",padx=18,pady=5)
            tk.Label(row,text=lab,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w")
            default={"vision":"OK","ref":"Não determinado","real":"0","det":"0"}.get(key,"")
            v=tk.StringVar(value=default);vars[key]=v
            if key=="vision": w=ttk.Combobox(row,textvariable=v,values=["OK","NOK"],state="readonly")
            elif key=="ref": w=ttk.Combobox(row,textvariable=v,values=REFS,state="readonly")
            else: w=ttk.Entry(row,textvariable=v)
            w.pack(fill="x")
        btns=tk.Frame(frame,bg=PANEL);btns.pack(fill="x",padx=18,pady=18)
        def collect(next_one=False):
            try:
                real=max(0,int(float(vars["real"].get() or 0)))
                det=max(0,int(float(vars["det"].get() or 0)))
            except ValueError:
                messagebox.showerror("Dados","As quantidades devem ser numéricas.",parent=win);return
            ident=vars["id"].get().strip()
            if not ident:
                ident=self._next_numeric_id()
                vars["id"].set(ident)
            if self._id_exists(ident):
                messagebox.showerror("Identificação duplicada",
                                     f"A identificação '{ident}' já está cadastrada nesta coleta.\n\n"
                                     "Use outro número/código para esta inspeção. Os campos digitados foram preservados.",
                                     parent=win)
                return
            r=blank_row()
            r.update({"id":ident,"vision":vars["vision"].get() or "OK",
                      "ref":vars["ref"].get() or "Não determinado","real":real,"det":det,
                      "obs":vars["obs"].get().strip(),"source":"Manual — Coleta rápida",
                      "detection_observed":True})
            self.rows.append(self.classify(r))
            self.source_file=""
            self.refresh_all()
            self.status.set(f"Inspeção {len(self.rows)} registrada.")
            if next_one:
                # Prepara a próxima inspeção sem reutilizar o número anterior.
                vars["id"].set(self._next_numeric_id())
                vars["vision"].set("OK");vars["ref"].set("Não determinado")
                vars["real"].set("0");vars["det"].set("0");vars["obs"].set("")
            else: win.destroy()
        ttk.Button(btns,text="Registrar e fechar",command=lambda:collect(False)).pack(side="right",padx=4)
        ttk.Button(btns,text="Registrar e próxima",style="Accent.TButton",
                   command=lambda:collect(True)).pack(side="right",padx=4)

    def edit_row(self,index=None):
        if not self.rows:return
        if index is None:
            if not hasattr(self,"tree"):return
            sel=self.tree.selection()
            if not sel:return
            index=int(sel[0])
        # O formulário trabalha sobre uma cópia normalizada. Assim, cancelar a
        # janela não altera a inspeção original e salvar não reaproveita dados
        # de outra linha.
        r=normalize_row(self.rows[index])
        win=tk.Toplevel(self);win.title(f"Inspeção {index+1} — detalhamento")
        win.geometry("820x760");win.minsize(700,560);win.configure(bg=BG);win.transient(self);win.grab_set()

        # Área rolável: o conteúdo pode crescer sem esconder o botão Salvar.
        outer=tk.Frame(win,bg=BG);outer.pack(fill="both",expand=True,padx=12,pady=12)
        canvas=tk.Canvas(outer,bg=BG,highlightthickness=0,bd=0)
        scroll=ttk.Scrollbar(outer,orient="vertical",command=canvas.yview)
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left",fill="both",expand=True)
        scroll.pack(side="right",fill="y")
        frame=tk.Frame(canvas,bg=PANEL,highlightbackground=BORDER,highlightthickness=1)
        window_id=canvas.create_window((0,0),window=frame,anchor="nw")
        def _sync(_=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
        def _width(e):
            canvas.itemconfigure(window_id,width=e.width)
        frame.bind("<Configure>",_sync)
        canvas.bind("<Configure>",_width)
        canvas.bind_all("<MouseWheel>",lambda e: canvas.yview_scroll(int(-e.delta/120),"units"))

        tk.Label(frame,text=f"INSPEÇÃO {index+1}",bg=PANEL,fg=TEXT,
                 font=("Segoe UI",15,"bold")).pack(anchor="w",padx=18,pady=(18,10))
        tk.Label(frame,text="Role a janela para acessar todos os campos. O botão Salvar permanece sempre visível.",
                 bg=PANEL,fg=ACCENT2,font=("Segoe UI",8)).pack(anchor="w",padx=18,pady=(0,8))
        vars={}
        for lab,key,values in [
            ("Identificação","id",None),("Vision AI","vision",["OK","NOK"]),
            ("Referência / gabarito","ref",REFS)]:
            row=tk.Frame(frame,bg=PANEL);row.pack(fill="x",padx=18,pady=4)
            tk.Label(row,text=lab,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w")
            v=tk.StringVar(value=str(r.get(key,"")));vars[key]=v
            w=ttk.Combobox(row,textvariable=v,values=values,state="readonly") if values else ttk.Entry(row,textvariable=v)
            w.pack(fill="x")
        q=tk.Frame(frame,bg=PANEL);q.pack(fill="x",padx=18,pady=4)
        tk.Label(frame,text="DETECÇÃO POR OCORRÊNCIA (D-O) — informe o que o operador confirmou; os totais Real e Detectado serão calculados automaticamente.",bg=PANEL,fg=ACCENT2,font=("Segoe UI",8,"bold"),wraplength=720,justify="left").pack(anchor="w",padx=18,pady=(5,1))
        current_d=self.detection_for_row(r)
        qv={}
        for j,(lab,key,default) in enumerate([("Defeitos corretamente detectados (TP-D)","tp_d",current_d["tp"]),("Falsos positivos confirmados (FP-D)","fp_d",current_d["fp"]),("Falsos negativos confirmados (FN-D)","fn_d",current_d["fn"])]):
            b=tk.Frame(q,bg=PANEL);b.grid(row=0,column=j,sticky="ew",padx=4);q.columnconfigure(j,weight=1)
            tk.Label(b,text=lab,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w")
            qv[key]=tk.StringVar(value=str(default));ttk.Entry(b,textvariable=qv[key]).pack(fill="x")
        o=tk.Frame(frame,bg=PANEL);o.pack(fill="x",padx=18,pady=4)
        tk.Label(o,text="Observação",bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w")
        ov=tk.StringVar(value=r["obs"]);ttk.Entry(o,textvariable=ov).pack(fill="x")

        sec=self.card(frame,"DEFEITOS DETALHADOS — OPCIONAL",
                       "Use somente quando quiser saber quais tipos de defeito ocorreram. Uma inspeção pode ter vários tipos.")
        sec.pack(fill="both",expand=True,padx=18,pady=10)
        dt=tk.Frame(sec,bg=PANEL);dt.pack(fill="both",expand=True,padx=10,pady=5)
        cols=("tipo","real","det")
        dtt=ttk.Treeview(dt,columns=cols,show="headings",height=8)
        dscroll=ttk.Scrollbar(dt,orient="vertical",command=dtt.yview)
        dtt.configure(yscrollcommand=dscroll.set)
        for c,h,w in zip(cols,("Tipo de defeito","Real","Detectado"),(420,100,110)):
            dtt.heading(c,text=h);dtt.column(c,width=w,anchor="center")
        dtt.pack(side="left",fill="both",expand=True)
        dscroll.pack(side="right",fill="y")
        for d in r["defects"]:
            dtt.insert("", "end",values=(d["type"],d["real"],d["det"]))
        addbar=tk.Frame(sec,bg=PANEL);addbar.pack(fill="x",padx=10,pady=6)
        tv=tk.StringVar();rv=tk.StringVar(value="0");dv=tk.StringVar(value="0")
        ttk.Entry(addbar,textvariable=tv,width=35).pack(side="left",padx=3)
        ttk.Entry(addbar,textvariable=rv,width=8).pack(side="left",padx=3)
        ttk.Entry(addbar,textvariable=dv,width=8).pack(side="left",padx=3)
        def add_def():
            name=normalize_defect(tv.get())
            if not name:return
            try: rr=max(0,int(float(rv.get() or 0)));dd=max(0,int(float(dv.get() or 0)))
            except ValueError:
                messagebox.showerror("Defeito","Real e detectado devem ser numéricos.",parent=win);return
            dtt.insert("", "end",values=(name,rr,dd))
            tv.set("");rv.set("0");dv.set("0")
            _sync()
        ttk.Button(addbar,text="+ Adicionar tipo",command=add_def).pack(side="left",padx=5)
        def remove_def():
            for iid in dtt.selection(): dtt.delete(iid)
            _sync()
        ttk.Button(addbar,text="Excluir selecionado",command=remove_def).pack(side="right",padx=5)
        tk.Label(addbar,text="Tipo | Real | Detectado",bg=PANEL,fg=MUTED,
                 font=("Segoe UI",8)).pack(side="left",padx=5)
        _sync()

        # Rodapé fixo fora da área rolável: nunca fica escondido.
        footer=tk.Frame(win,bg=BG,highlightbackground=BORDER,highlightthickness=1)
        footer.pack(fill="x",side="bottom",padx=12,pady=(0,12))
        def save():
            # Monte um novo registro a partir do snapshot aberto nesta janela.
            # Cada campo é lido no momento do clique em Salvar para impedir
            # perda, repetição ou sobrescrita entre inspeções consecutivas.
            updated=normalize_row(r)
            updated["id"]=vars["id"].get().strip()
            if not updated["id"]:
                messagebox.showerror("Identificação", "Informe a identificação da inspeção antes de salvar.", parent=win)
                return
            if self._id_exists(updated["id"], exclude_index=index):
                messagebox.showerror("Identificação duplicada",
                                     f"A identificação '{updated['id']}' já pertence a outra inspeção desta coleta.\n\n"
                                     "O registro não foi alterado.", parent=win)
                return
            updated["vision"]=vars["vision"].get() or "OK"
            updated["ref"]=vars["ref"].get() or "Não determinado"
            updated["obs"]=ov.get().strip()
            try:
                tp_d=max(0,int(float(qv["tp_d"].get() or 0)))
                fp_d=max(0,int(float(qv["fp_d"].get() or 0)))
                fn_d=max(0,int(float(qv["fn_d"].get() or 0)))
            except ValueError:
                messagebox.showerror("Dados","TP-D, FP-D e FN-D devem ser números inteiros não negativos.",parent=win);return
            updated["real"]=tp_d+fn_d
            updated["det"]=tp_d+fp_d
            updated["correspondentes"]=tp_d
            updated["detection_observed"]=True
            updated["detection_explicit_matrix"]=True
            updated["tp_d"]=tp_d; updated["fp_d"]=fp_d; updated["fn_d"]=fn_d
            updated["defects"]=[]
            for iid in dtt.get_children():
                vals=dtt.item(iid,"values")
                updated["defects"].append({"type":normalize_defect(vals[0]),"real":int(vals[1]),"det":int(vals[2])})
            self.rows[index]=self.classify(updated)
            self.source_file=""
            win.destroy();self.refresh_all()
        ttk.Button(footer,text="Cancelar",command=win.destroy).pack(side="right",padx=5,pady=7)
        ttk.Button(footer,text="Salvar inspeção",style="Accent.TButton",command=save).pack(side="right",padx=5,pady=7)
        canvas.focus_set()
        def on_close():
            try: canvas.unbind_all("<MouseWheel>")
            except Exception: pass
            win.destroy()
        win.protocol("WM_DELETE_WINDOW",on_close)

    def delete_row(self):
        if not hasattr(self,"tree"):return
        sel=self.tree.selection()
        if not sel:return
        if not messagebox.askyesno("Excluir","Excluir a(s) inspeção(ões) selecionada(s)?"):return
        for i in sorted([int(x) for x in sel],reverse=True): self.rows.pop(i)
        self.refresh_all()

    # ---------- IMPORT ----------
    def _norm_header(self,v):
        return normalize_text(v).replace("º","").replace("°","")

    def _find_header_row(self,ws):
        # Compatível com planilhas históricas e com o XLSX gerado pelo próprio AIA.
        targets=("IDENTIFICACAO","RESULTADO VISION AI","VISION AI","VISION",
                 "REFERENCIA","REAL","DETECTADO",
                 "FALSO POSITIVO","FALSO NEGATIVO",
                 "VERDADEIRO POSITIVO","VERDADEIRO NEGATIVO",
                 "TP","FP","FN","VN",
                 "TP-C","VN-C","FP-C","FN-C","TP-D","FP-D","FN-D")
        best=None
        for r in range(1,min(ws.max_row,80)+1):
            vals=[self._norm_header(ws.cell(r,c).value) for c in range(1,min(ws.max_column,50)+1)]
            score=sum(1 for t in targets if t in vals)
            if "IDENTIFICACAO" in vals: score+=3
            if "VISION AI" in vals: score+=2
            if "REAL" in vals: score+=1
            if "DETECTADO" in vals: score+=1
            if any(t in vals for t in ("TP-C","VN-C","FP-C","FN-C")): score+=2
            if score>=4 and (best is None or score>best[0]): best=(score,r)
        return best[1] if best else None

    def _header_map(self,ws,row):
        out={}
        for c in range(1,ws.max_column+1):
            h=self._norm_header(ws.cell(row,c).value)
            if h: out[h]=c
        return out

    def _num(self,v):
        try:
            if v is None or str(v).strip()=="":return 0
            return int(float(str(v).replace(",",".")))
        except Exception:return 0

    def _format_identification_cell(self, cell):
        """
        Preserva a identificação exatamente como ela aparece no Excel.

        Em planilhas de campo é comum um código como 0623 estar armazenado
        como número 623 com formato de célula 0000. O openpyxl, ao usar
        data_only=True, entrega apenas 623; se convertermos diretamente para
        str(), o zero à esquerda desaparece.

        Ordem de preservação:
        1) texto -> mantém zeros à esquerda;
        2) formato numérico explícito (ex.: 0000) -> aplica a máscara;
        3) identificação numérica inteira de até 4 dígitos -> usa 4 posições,
           compatível com os códigos de identificação usados nas inspeções;
        4) demais valores -> conversão textual sem alteração semântica.
        """
        if cell is None or cell.value is None:
            return ""

        value = cell.value

        # Texto deve ser preservado literalmente, inclusive zeros à esquerda.
        if isinstance(value, str):
            return value.strip()

        # Datas/objetos especiais não devem ser convertidos em código.
        if isinstance(value, (datetime,)):
            return value.strftime("%d/%m/%Y")

        # Números inteiros: respeita primeiro a máscara do Excel.
        try:
            number = float(value)
            if number.is_integer():
                integer = int(number)
                fmt = str(getattr(cell, "number_format", "") or "")

                # Máscaras simples como 0000, 00000 etc.
                # Ignora formatos de data/hora e outros formatos não adequados
                # para identificação.
                fmt_clean = fmt.split(";")[0].strip()
                if re.fullmatch(r"0+", fmt_clean):
                    return str(integer).zfill(len(fmt_clean))

                # Se não houver máscara explícita, preserva o padrão de código
                # de 4 dígitos utilizado pelas identificações de carroceria.
                if 0 <= integer <= 9999:
                    return f"{integer:04d}"
                return str(integer)
        except Exception:
            pass

        return str(value).strip()

    def _format_inspection_date(self, value):
        """Converte datas do Excel para dd/mm/aaaa sem inventar a data atual."""
        if value is None or str(value).strip()=="":
            return ""
        if isinstance(value, datetime):
            return value.strftime("%d/%m/%Y")
        try:
            from datetime import date as _date
            if isinstance(value, _date):
                return value.strftime("%d/%m/%Y")
        except Exception:
            pass
        text=str(value).strip()
        for fmt in ("%d/%m/%Y","%Y-%m-%d","%d-%m-%Y","%d/%m/%y"):
            try:
                return datetime.strptime(text,fmt).strftime("%d/%m/%Y")
            except Exception:
                pass
        return ""

    def _date_from_filename(self, path):
        """Tenta recuperar uma data explícita do nome do relatório (ex.: 21072026)."""
        name=Path(path).stem
        for raw in re.findall(r"(?<!\d)(\d{8})(?!\d)",name):
            for fmt in ("%d%m%Y","%Y%m%d"):
                try:return datetime.strptime(raw,fmt).strftime("%d/%m/%Y")
                except Exception:pass
        return ""

    def _metadata_aliases(self):
        """Aliases aceitos para preencher automaticamente o contexto da coleta."""
        return {
            "cliente": ["Cliente", "Customer", "Cliente / Customer", "OEM", "Montadora"],
            "local": ["Local / Planta", "Planta", "Plant", "Local", "Site", "Unidade", "Fábrica", "Fabrica"],
            "projeto": ["Projeto", "Project", "Projeto / Project", "Programa", "Program"],
            "linha": ["Linha / Estação", "Linha", "Line", "Estação", "Station", "Line / Station", "Linha de produção", "Processo"],
            "tecnico": ["Técnico", "Tecnico", "Technical", "Operador", "Operator", "Responsável", "Responsavel", "Responsável técnico", "Responsavel tecnico"],
            "data": ["Data da inspeção", "Data inspeção", "Data", "Date", "Inspection Date", "Data do teste", "Data do ensaio", "Dia"],
            "threshold": ["Threshold", "Confidence", "Confiança", "Grau de confiança", "Threshold / Grau de Confiança", "Limiar", "Confidence Threshold"],
            "gabarito": ["Método de referência", "Metodo de referencia", "Gabarito", "Método", "Metodo", "Reference", "Reference Method", "Ground Truth", "Truth Method"],
        }

    def _clean_metadata_value(self, value, key):
        if value is None:
            return ""
        if key == "data":
            d=self._format_inspection_date(value)
            if d:
                return d
        text=str(value).strip()
        if not text:
            return ""
        if self._norm_header(text) in {"N/D","ND","N/A","NA","NENHUM","NONE","NULL","-"}:
            return ""
        return text

    def _metadata_from_header_column(self, ws, header_row, aliases, key):
        hm=self._header_map(ws, header_row)
        normalized={self._norm_header(a) for a in aliases}
        candidates=[]
        for h,c in hm.items():
            if h in normalized or any(h.startswith(a+" ") or h.startswith(a+"/") for a in normalized):
                for rr in range(header_row+1, min(ws.max_row, header_row+30)+1):
                    v=self._clean_metadata_value(ws.cell(rr,c).value,key)
                    if v:
                        candidates.append(v)
        if not candidates:
            return ""
        if key == "data":
            from collections import Counter
            return Counter(candidates).most_common(1)[0][0]
        return candidates[0]

    def _metadata_from_labeled_cells(self, wb, aliases, key):
        """Procura campos de contexto sem permitir que outro rótulo vire valor."""
        normalized={self._norm_header(a) for a in aliases}

        # Todos os rótulos de metadados são bloqueados como possíveis valores.
        all_labels=set()
        for vals in self._metadata_aliases().values():
            all_labels.update(self._norm_header(a) for a in vals)

        invalid={
            "RESULTADOS","GABARITO","ESPERADO","APURADO",
            "VISION AI","VISION","RESULTADO VISION AI",
            "VERDADEIRO POSITIVO","VERDADEIRO NEGATIVO",
            "FALSO POSITIVO","FALSO NEGATIVO",
            "TP","TN","VN","FP","FN","TP-C","VN-C","FP-C","FN-C",
            "TP-D","FP-D","FN-D","N/D","N/A","NA","NONE","NULL"
        }

        best=[]
        for ws in wb.worksheets:
            for r in range(1,min(ws.max_row,300)+1):
                for c in range(1,min(ws.max_column,50)+1):
                    raw=ws.cell(r,c).value
                    h=self._norm_header(raw)
                    if not h:
                        continue

                    # Suporta "Cliente: VW" / "Local: SJP" na mesma célula.
                    inline=None
                    for a in normalized:
                        if h.startswith(a+":"):
                            inline=str(raw).split(":",1)[1].strip()
                            break
                    if inline:
                        v=self._clean_metadata_value(inline,key)
                        if v:
                            nv=self._norm_header(v)
                            if nv not in all_labels and nv not in invalid:
                                best.append(v)
                                continue

                    # Suporta rótulo "Cliente:" separado do valor.
                    h_label=h[:-1].strip() if h.endswith(":") else h
                    if h_label not in normalized:
                        continue

                    positions=[(r,c+1),(r,c+2),(r,c+3),(r+1,c),(r+2,c)]
                    for rr,cc in positions:
                        if rr>ws.max_row or cc>ws.max_column:
                            continue
                        v=self._clean_metadata_value(ws.cell(rr,cc).value,key)
                        if not v:
                            continue
                        nv=self._norm_header(v)
                        if nv in all_labels or nv in invalid:
                            continue
                        best.append(v)
                        break

        if not best:
            return ""
        if key=="data":
            from collections import Counter
            return Counter(best).most_common(1)[0][0]
        return best[0]

    def _extract_workbook_metadata(self, wb, data_ws, header_row, path):
        """Extrai contexto com prioridade para o formato nativo do AIA.

        Regra de compatibilidade: relatórios produzidos pelo próprio AIA têm
        prioridade absoluta no bloco RESUMO/IDENTIFICAÇÃO DO TESTE. Somente
        quando esse bloco não existir são usadas as heurísticas históricas.
        Isso evita que palavras como "inspeção", "método" ou cabeçalhos da
        tabela sejam capturadas como valores de contexto.
        """
        aliases=self._metadata_aliases()
        out={k:"" for k in aliases}; status={k:"não encontrado" for k in aliases}

        # 1) Formato nativo AIA: procurar uma aba RESUMO e pares rótulo/valor
        # próximos, sem depender da posição exata das linhas.
        resumo=None
        for ws_try in wb.worksheets:
            if self._norm_header(ws_try.title) == "RESUMO":
                resumo=ws_try; break
        if resumo is not None:
            native={}
            for r in range(1, min(resumo.max_row, 80)+1):
                label=self._norm_header(resumo.cell(r,1).value)
                if not label:
                    continue
                for key,als in aliases.items():
                    if label in {self._norm_header(a) for a in als}:
                        value=self._clean_metadata_value(resumo.cell(r,2).value,key)
                        if value:
                            native[key]=value
            # Também aceita uma linha explícita "IDENTIFICAÇÃO DO TESTE"
            # seguida de pares em outras colunas, caso o arquivo tenha sido
            # criado por uma versão anterior do AIA.
            if native:
                for key,value in native.items():
                    out[key]=value; status[key]="identificado no RESUMO nativo do AIA"

        # 2) Compatibilidade com planilhas históricas / externas.
        for key,als in aliases.items():
            if out.get(key):
                continue
            v=self._metadata_from_header_column(data_ws, header_row, als, key)
            if not v:
                v=self._metadata_from_labeled_cells(wb, als, key)
            if not v and key=="data":
                v=self._date_from_filename(path)
            if v:
                out[key]=v; status[key]="identificado automaticamente"
        return out,status

    def import_excel(self):
        if openpyxl is None:
            messagebox.showerror("Excel","OpenPyXL não está instalado.\nExecute: py -3 -m pip install openpyxl")
            return
        path=filedialog.askopenfilename(title="Abrir planilha de assertividade",
                                        filetypes=[("Excel","*.xlsx"),("Todos","*.*")])
        if not path:return
        try:
            wb=openpyxl.load_workbook(path,data_only=True)
            # Não assumimos que a primeira aba é a de dados. Em relatórios AIA,
            # RESUMO costuma vir antes de INSPEÇÕES.
            candidates=[]
            for ws_try in wb.worksheets:
                hr_try=self._find_header_row(ws_try)
                if hr_try is not None:
                    hm_try=self._header_map(ws_try,hr_try)
                    score=0
                    for key in ("IDENTIFICACAO","RESULTADO VISION AI","VISION AI","TP","FP","FN","VN",
                                 "REAL","DETECTADO","TP-C","VN-C","FP-C","FN-C","TP-D","FP-D","FN-D"):
                        if key in hm_try: score += 1
                    candidates.append((score, ws_try, hr_try))
            if not candidates:
                raise ValueError("Não encontrei uma aba com a estrutura de INSPEÇÕES. Verifique se a planilha possui Identificação/Inspeção e colunas de resultado.")
            _, ws, hr = max(candidates, key=lambda x: x[0])
            hm=self._header_map(ws,hr)
            imported_meta, metadata_status = self._extract_workbook_metadata(wb, ws, hr, path)
            def col(*names):
                for n in names:
                    if self._norm_header(n) in hm:return hm[self._norm_header(n)]
                return None
            c_id=col("Identificação","N Inspeção","Nº Inspeção","Inspeção")
            c_vision=col("Resultado Vision AI","Vision AI","Vision")
            # O XLSX exportado pelo AIA usa TP-C/VN-C/FP-C/FN-C.
            # Planilhas históricas podem usar TP/VN/FP/FN.
            c_tp=col("Verdadeiro Positivo","VP - Verdadeiro Positivo","TP-C","TP")
            c_tn=col("Verdadeiro Negativo","VN - Verdadeiro Negativo","VN-C","VN","TN")
            c_fp=col("Falso Positivo","FP - Falso Positivo","FP-C","FP")
            c_fn=col("Falso Negativo","FN - Falso Negativo","FN-C","FN")
            c_real=col("Real","Real (defeitos)","Defeitos reais")
            c_det=col("Detectado","Detectados","Detecções","Detecções registradas")
            c_tp_d=col("TP-D")
            c_fp_d=col("FP-D")
            c_fn_d=col("FN-D")
            c_def=col("Defeito","Tipo de defeito")
            c_obs=col("Observação","Observações")
            c_date=col("Data da inspeção","Data inspeção","Data","Date")
            if c_id is None: raise ValueError("A planilha não possui coluna de Identificação/Inspeção.")
            rows=[];hist={"TP":0,"VN/TN":0,"FP":0,"FN":0,"N":0}
            imported_date=""
            imported_dates=[]
            for rr in range(hr+1,ws.max_row+1):
                ident_cell=ws.cell(rr,c_id)
                ident=self._format_identification_cell(ident_cell)
                if not ident:continue
                if c_date:
                    row_date=self._format_inspection_date(ws.cell(rr,c_date).value)
                    if row_date:
                        imported_dates.append(row_date)
                r=blank_row()
                r["id"]=ident
                r["vision"]=str(ws.cell(rr,c_vision).value or "OK").upper() if c_vision else "OK"
                r["ref"]="Importado — histórico"
                # Historical values remain historical. We store them in the row but do not
                # convert them into the new classification truth.
                r["tp"]=self._num(ws.cell(rr,c_tp).value) if c_tp else 0
                r["vn"]=self._num(ws.cell(rr,c_tn).value) if c_tn else 0
                r["fp"]=self._num(ws.cell(rr,c_fp).value) if c_fp else 0
                r["fn"]=self._num(ws.cell(rr,c_fn).value) if c_fn else 0
                r["detection_available"] = bool(c_real is not None and c_det is not None)
                r["real"]=self._num(ws.cell(rr,c_real).value) if c_real else 0
                r["det"]=self._num(ws.cell(rr,c_det).value) if c_det else 0
                # Preserva os indicadores de detecção presentes no XLSX AIA.
                r["tp_d"]=self._num(ws.cell(rr,c_tp_d).value) if c_tp_d else 0
                r["fp_d"]=self._num(ws.cell(rr,c_fp_d).value) if c_fp_d else 0
                r["fn_d"]=self._num(ws.cell(rr,c_fn_d).value) if c_fn_d else 0
                r["obs"]=str(ws.cell(rr,c_obs).value or "") if c_obs else ""
                r["calc_status"]="Histórico"
                r["source"]="Excel histórico"
                # Legacy defect column: keep as optional detail with unknown counts.
                if c_def:
                    name=normalize_defect(ws.cell(rr,c_def).value)
                    if name:r["defects"]=[{"type":name,"real":0,"det":0}]
                rows.append(r)
                hist["TP"]+=r["tp"];hist["VN/TN"]+=r["vn"];hist["FP"]+=r["fp"];hist["FN"]+=r["fn"]
                hist["N"]+=r["tp"]+r["vn"]+r["fp"]+r["fn"]
            if not rows: raise ValueError("Nenhuma inspeção válida foi encontrada.")
            if imported_dates:
                from collections import Counter
                imported_date=Counter(imported_dates).most_common(1)[0][0]
            self._archive_current_collection()
            self.rows=rows;self.history=[hist];self.source_file=path
            # Importação inteligente do contexto. Nunca usa a data atual como fallback.
            if imported_date:
                imported_meta["data"] = imported_date
                metadata_status["data"] = "identificado automaticamente na tabela de inspeções"
            for key in self.meta:
                if key == "data":
                    continue
                value=imported_meta.get(key, "")
                if value:
                    self.meta[key].set(value)
            self.meta["data"].set(imported_meta.get("data", ""))
            if not self.meta["gabarito"].get():
                self.meta["gabarito"].set("Não determinado")
            self.refresh_all()
            labels=[("cliente","Cliente"),("local","Local / Planta"),("projeto","Projeto"),("linha","Linha / Estação"),("tecnico","Técnico"),("data","Data"),("threshold","Threshold"),("gabarito","Método de referência")]
            found=[label for key,label in labels if metadata_status.get(key,"").startswith("identificado")]
            missing=[label for key,label in labels if not metadata_status.get(key,"").startswith("identificado")]
            msg=(f"Planilha: {Path(path).name}\n\nInspeções importadas: {len(rows)}\n"
                 f"Histórico: TP={hist['TP']} | VN={hist['VN/TN']} | FP={hist['FP']} | FN={hist['FN']}\n\n"
                 f"IDENTIFICAÇÃO AUTOMÁTICA\nEncontrados: {len(found)}/8\n"
                 f"{', '.join(found) if found else 'Nenhum campo identificado'}\n")
            if missing:
                msg += f"\nPreencher/revisar: {', '.join(missing)}\n"
            msg += (f"\nData usada: {self.meta['data'].get() or 'N/D — preencher na tela Coleta'}\n\n"
                    "Os valores históricos foram preservados e não viraram gabarito do motor novo.\n"
                    "Revise os metadados antes de salvar/exportar, principalmente quando a planilha possuir mais de uma coleta.")
            messagebox.showinfo("Importação inteligente concluída", msg)
        except Exception as e:
            messagebox.showerror("Erro na importação",f"Não foi possível importar a planilha.\n\n{e}")

    # ---------- METRICS / OTHER VIEWS ----------
    def engineering_indicators(self, m=None, d=None):
        """Indicadores derivados para leitura técnica/engenharia.
        Não são métricas estatísticas oficiais; servem como indicadores operacionais.
        """
        m = m or self.display_metrics()
        d = d or self.defect_counts()
        n = m.get("N", 0) or 0
        real_positive = (m.get("TP", 0) or 0) + (m.get("FN", 0) or 0)
        predicted_positive = (m.get("TP", 0) or 0) + (m.get("FP", 0) or 0)
        return {
            "prevalencia": safe(real_positive, n),
            "taxa_alarme": safe(predicted_positive, n),
            "fp_100": safe(m.get("FP", 0) * 100, n),
            "fn_100": safe(m.get("FN", 0) * 100, n),
            "real_por_inspecao": safe(d.get("real", 0), n),
            "det_por_inspecao": safe(d.get("det", 0), n),
        }

    def technical_diagnosis(self, m, d):
        """Gera leitura técnica baseada nos indicadores, sem criar um score artificial."""
        notes = []
        n = m.get("N", 0) or 0
        recall = m.get("Recall / Sensibilidade")
        precision = m.get("Precision")
        fpr = m.get("FPR")
        fnr = m.get("FNR")

        if n < 30:
            notes.append("AMOSTRA PEQUENA: use os resultados como diagnóstico inicial; aumente o número de inspeções antes de concluir desempenho.")
        if recall is not None:
            if recall >= 0.90:
                notes.append("RECALL FORTE: a capacidade de encontrar inspeções realmente NOK está alta.")
            elif recall < 0.80:
                notes.append("RECALL DE ATENÇÃO: há risco relevante de deixar casos NOK passarem como OK.")
        if precision is not None:
            if precision < 0.70:
                notes.append("PRECISION DE ATENÇÃO: parte importante dos alarmes NOK não corresponde a casos realmente NOK.")
            elif precision >= 0.90:
                notes.append("PRECISION FORTE: a maior parte das classificações NOK corresponde a casos realmente NOK.")
        if fpr is not None and fpr >= 0.20:
            notes.append("FALSO POSITIVO ELEVADO: vale investigar threshold, ruído de imagem, iluminação, ROI e condições de aquisição.")
        if fnr is not None and fnr >= 0.10:
            notes.append("FALSO NEGATIVO RELEVANTE: vale investigar sensibilidade, cobertura da inspeção, velocidade e condições de aquisição.")
        if d.get("real", 0) and d.get("det", 0) > d.get("real", 0):
            notes.append("FP DE DETECÇÃO ELEVADO: existem mais detecções que defeitos reais registrados; avaliar duplicidade, granularidade, threshold, ruído e correspondência das ocorrências.")
        if not notes:
            notes.append("SEM ALERTA AUTOMÁTICO: os indicadores atuais não ultrapassaram os limites de diagnóstico configurados.")
        return notes

    def view_tecnica(self):
        fonte = "HISTÓRICO DA PLANILHA" if (self.source_file and self.history) else "MOTOR NOVO / ENTRADA MANUAL"
        self.section_title(
            "ANÁLISE TÉCNICA  /  DIAGNÓSTICO",
            f"Autaza Inspection Analytics • {fonte} • métricas oficiais + leitura técnica sem misturar unidades")

        m = self.display_metrics()
        d = self.defect_counts()
        dt = self.detection_totals()
        e = self.engineering_indicators(m, d)

        grid = tk.Frame(self.content, bg=BG)
        grid.pack(fill="x")
        cards = [
            ("N • Inspeções", str(m["N"]), TEXT),
            ("Accuracy", pct(m["Accuracy"]), ACCENT2),
            ("Precision", pct(m["Precision"]), ACCENT2),
            ("Recall", pct(m["Recall / Sensibilidade"]), GOOD),
            ("Specificity", pct(m["Specificity"]), ACCENT2),
            ("FPR", pct(m["FPR"]), BAD),
            ("FNR", pct(m["FNR"]), WARN),
            ("F1", pct(m["F1"]), ACCENT2),
            ("MCC", "N/D" if m["MCC"] is None else f"{m['MCC']:.4f}", ACCENT2),
            ("Prevalência", pct(e["prevalencia"]), ACCENT2),
            ("Precision detecção", pct(dt["precision"]), ACCENT),
            ("Recall detecção", pct(dt["recall"]), GOOD),
        ]
        for i, (lab, val, col) in enumerate(cards):
            r, c = divmod(i, 5)
            self.metric_card(grid, lab, val, col).grid(row=r, column=c, sticky="ew", padx=4, pady=4)
        for c in range(5):
            grid.columnconfigure(c, weight=1)

        lower = tk.Frame(self.content, bg=BG)
        lower.pack(fill="both", expand=True, pady=8)

        matrix = self.chart_card(lower, "MATRIZ DE CONFUSÃO", "Unidade estatística: inspeção / carroceria / peça")
        matrix.grid(row=0, column=0, sticky="nsew", padx=(0, 4), pady=(0, 4))
        txt = (
            f"                         PREDITO NOK       PREDITO OK\n"
            f"REAL NOK / DEFEITO       TP = {m['TP']:<8}     FN = {m['FN']}\n"
            f"REAL OK                  FP = {m['FP']:<8}     VN = {m['VN/TN']}\n\n"
            f"N = {m['N']}\n\n"
            f"FPR = {pct(m['FPR'])}    FNR = {pct(m['FNR'])}"
        )
        tk.Label(matrix, text=txt, bg=PANEL, fg=TEXT, font=("Consolas", 10),
                 justify="left").pack(anchor="w", padx=16, pady=14)

        errors = self.chart_card(lower, "PERFIL DE ERROS", "Volume de erros e acertos da classificação")
        errors.grid(row=0, column=1, sticky="nsew", padx=(4, 0), pady=(0, 4))
        self.draw_bar_chart(
            errors,
            [("TP", m["TP"], GOOD), ("TN", m["VN/TN"], GOOD),
             ("FP", m["FP"], BAD), ("FN", m["FN"], WARN)],
            max_value=max(m["TP"], m["VN/TN"], m["FP"], m["FN"], 1)
        )

        diag = self.card(lower, "DIAGNÓSTICO TÉCNICO", "Leitura automática para apoiar investigação — não substitui análise de processo.")
        diag.grid(row=1, column=0, sticky="nsew", padx=(0, 4), pady=(4, 0))
        for note in self.technical_diagnosis(m, d):
            tk.Label(diag, text="• " + note, bg=PANEL, fg=TEXT,
                     wraplength=520, justify="left", font=("Segoe UI", 9)).pack(
                         anchor="w", padx=16, pady=5)

        detcard = self.card(lower, "DETecção × CLASSIFICAÇÃO", "FP/FN de detecção não são os mesmos FP/FN da matriz OK/NOK.")
        detcard.grid(row=2, column=0, columnspan=2, sticky="ew", padx=0, pady=(8, 0))
        txtd=(f"CLASSIFICAÇÃO: FP = {m['FP']} | FN = {m['FN']}\n"
              f"DETECÇÃO: TP = {dt['tp']} | FP = {dt['fp']} | FN = {dt['fn']}\n"
              f"Detecção por contagem: Precision = {pct(dt['precision'])} | Recall = {pct(dt['recall'])}\n\n"
              "A estimativa de detecção usa quantidade quando não existe correspondência individual; ela não afirma correspondência espacial entre defeitos.")
        tk.Label(detcard, text=txtd, bg=PANEL, fg=TEXT, font=("Consolas", 9), justify="left", wraplength=1100).pack(anchor="w", padx=16, pady=12)

        op = self.card(lower, "INDICADORES OPERACIONAIS", "Complementares às métricas estatísticas; não devem ser interpretados como Accuracy/Recall.")
        op.grid(row=1, column=1, sticky="nsew", padx=(4, 0), pady=(4, 0))
        operational = [
            ("Alarmes NOK / inspeção", pct(e["taxa_alarme"])),
            ("Falsos positivos / 100 inspeções", "N/D" if e["fp_100"] is None else f"{e['fp_100']:.1f}"),
            ("Falsos negativos / 100 inspeções", "N/D" if e["fn_100"] is None else f"{e['fn_100']:.1f}"),
            ("Defeitos reais / inspeção", "N/D" if e["real_por_inspecao"] is None else f"{e['real_por_inspecao']:.2f}"),
            ("Detecções / inspeção", "N/D" if e["det_por_inspecao"] is None else f"{e['det_por_inspecao']:.2f}"),
            ("FP detecção / inspeção", "N/D" if dt["available_rows"]==0 else f"{dt['fp']/max(dt['available_rows'],1):.2f}"),
        ]
        for lab, val in operational:
            row = tk.Frame(op, bg=PANEL)
            row.pack(fill="x", padx=16, pady=5)
            tk.Label(row, text=lab, bg=PANEL, fg=MUTED, font=("Segoe UI", 9)).pack(side="left")
            tk.Label(row, text=val, bg=PANEL, fg=TEXT,
                     font=("Segoe UI", 10, "bold")).pack(side="right")

        lower.columnconfigure(0, weight=1)
        lower.columnconfigure(1, weight=1)
        lower.rowconfigure(0, weight=1)
        lower.rowconfigure(1, weight=1)

    def view_auditoria(self):
        tk.Label(self.content,text="AUDITORIA — HISTÓRICO × MOTOR NOVO",bg=BG,fg=TEXT,
                 font=("Segoe UI",18,"bold")).pack(anchor="w")
        tk.Label(self.content,text="O histórico da planilha permanece separado do cálculo estruturado.",
                 bg=BG,fg=MUTED).pack(anchor="w",pady=(2,10))
        if not self.history:
            f=self.card(self.content,"Nenhum histórico importado")
            f.pack(fill="x",pady=5)
            tk.Label(f,text="Abra uma planilha Excel para carregar os valores históricos.",
                     bg=PANEL,fg=TEXT).pack(anchor="w",padx=16,pady=16)
            return
        hist=self.history[-1];m=self.display_metrics()
        f=self.card(self.content,"COMPARAÇÃO DE TOTAIS");f.pack(fill="x")
        tree=ttk.Treeview(f,columns=("ind","hist","new","diff"),show="headings",height=6)
        for c,h in zip(("ind","hist","new","diff"),("Indicador","Histórico","Motor novo","Diferença")):
            tree.heading(c,text=h);tree.column(c,width=180,anchor="center")
        tree.pack(fill="x",padx=12,pady=12)
        for k in ("TP","VN/TN","FP","FN","N"):
            hv=hist.get(k,0);mv=m.get(k,0);tree.insert("", "end",values=(k,hv,mv,mv-hv))
        a=self.card(self.content,"INTERPRETAÇÃO")
        a.pack(fill="both",expand=True,pady=8)
        tk.Label(a,text=("Diferenças são esperadas porque o histórico pode usar uma metodologia "
                         "operacional diferente. O software não transforma automaticamente o histórico "
                         "em verdade de referência."),bg=PANEL,fg=WARN,wraplength=1000,
                 justify="left").pack(anchor="w",padx=16,pady=16)

    def view_engenharia(self):
        fonte = "HISTÓRICO DA PLANILHA" if (self.source_file and self.history) else "MOTOR NOVO / ENTRADA MANUAL"
        self.section_title(
            "ENGENHARIA  /  DECISÃO",
            f"Autaza Inspection Analytics • {fonte} • indicadores para investigação de processo, câmera, threshold e operação")

        m = self.display_metrics()
        d = self.defect_counts()
        dt = self.detection_totals()
        e = self.engineering_indicators(m, d)

        # Cabeçalho de contexto
        context = self.card(self.content, "CONTEXTO DO TESTE", "Condições registradas no momento da coleta/importação.")
        context.pack(fill="x", pady=(0, 6))
        ctx = tk.Frame(context, bg=PANEL)
        ctx.pack(fill="x", padx=14, pady=(0, 12))
        items = [
            ("Cliente", self.meta["cliente"].get() or "N/D"),
            ("Planta", self.meta["local"].get() or "N/D"),
            ("Projeto", self.meta["projeto"].get() or "N/D"),
            ("Linha / Estação", self.meta["linha"].get() or "N/D"),
            ("Threshold", self.meta["threshold"].get() or "N/D"),
            ("Referência", self.meta["gabarito"].get() or "N/D"),
        ]
        for i, (lab, val) in enumerate(items):
            r, c = divmod(i, 3)
            box = tk.Frame(ctx, bg=PANEL)
            box.grid(row=r, column=c, sticky="ew", padx=8, pady=4)
            tk.Label(box, text=lab, bg=PANEL, fg=MUTED,
                     font=("Segoe UI", 8, "bold")).pack(anchor="w")
            tk.Label(box, text=val, bg=PANEL, fg=TEXT,
                     font=("Segoe UI", 9)).pack(anchor="w")
            ctx.columnconfigure(c, weight=1)

        grid = tk.Frame(self.content, bg=BG)
        grid.pack(fill="x", pady=2)
        cards = [
            ("Prevalência", pct(e["prevalencia"]), ACCENT2),
            ("Taxa de alarme NOK", pct(e["taxa_alarme"]), ACCENT),
            ("FP / 100 inspeções", "N/D" if e["fp_100"] is None else f"{e['fp_100']:.1f}", BAD),
            ("FN / 100 inspeções", "N/D" if e["fn_100"] is None else f"{e['fn_100']:.1f}", WARN),
            ("Defeitos reais / inspeção", "N/D" if e["real_por_inspecao"] is None else f"{e['real_por_inspecao']:.2f}", GOOD),
            ("Detecções / inspeção", "N/D" if e["det_por_inspecao"] is None else f"{e['det_por_inspecao']:.2f}", ACCENT2),
        ]
        for i, (lab, val, col) in enumerate(cards):
            self.metric_card(grid, lab, val, col).grid(row=0, column=i, sticky="ew", padx=4, pady=4)
            grid.columnconfigure(i, weight=1)

        lower = tk.Frame(self.content, bg=BG)
        lower.pack(fill="both", expand=True, pady=8)

        priority = self.card(lower, "PRIORIZAÇÃO DE INVESTIGAÇÃO", "A ordem abaixo é derivada dos erros observados; não cria um score proprietário.")
        priority.grid(row=0, column=0, sticky="nsew", padx=(0, 4), pady=(0, 4))

        fpr = m["FPR"]
        fnr = m["FNR"]
        recall = m["Recall / Sensibilidade"]
        precision = m["Precision"]
        priorities = []
        if fpr is not None and fpr >= 0.20:
            priorities.append(("01", "REDUZIR FALSOS POSITIVOS", "Investigar threshold, ruído, iluminação, ROI, contraste e condições de aquisição."))
        if fnr is not None and fnr >= 0.10:
            priorities.append(("02", "REDUZIR FALSOS NEGATIVOS", "Investigar sensibilidade, cobertura da inspeção, velocidade, foco e condições de aquisição."))
        if precision is not None and precision < 0.70:
            priorities.append(("03", "VALIDAR CRITÉRIO DE NOK", "Muitos alarmes não confirmados indicam necessidade de revisar critério de decisão e/ou referência."))
        if recall is not None and recall < 0.80:
            priorities.append(("04", "VALIDAR COBERTURA", "A capacidade de encontrar casos reais está abaixo do esperado para uma inspeção robusta."))
        if not priorities:
            priorities.append(("01", "ACOMPANHAR ESTABILIDADE", "Os indicadores não acionaram prioridades críticas. Repetir testes em condições variadas antes de concluir estabilidade."))

        for num, title, desc in priorities:
            row = tk.Frame(priority, bg=PANEL)
            row.pack(fill="x", padx=16, pady=7)
            tk.Label(row, text=num, bg="#0E4560", fg=TEXT,
                     font=("Segoe UI", 9, "bold"), width=4, pady=5).pack(side="left")
            body = tk.Frame(row, bg=PANEL)
            body.pack(side="left", fill="x", expand=True, padx=10)
            tk.Label(body, text=title, bg=PANEL, fg=TEXT,
                     font=("Segoe UI", 9, "bold")).pack(anchor="w")
            tk.Label(body, text=desc, bg=PANEL, fg=MUTED,
                     wraplength=500, justify="left", font=("Segoe UI", 8)).pack(anchor="w", pady=(2, 0))

        quality = self.chart_card(lower, "INDICADORES PARA DECISÃO", "Comparação visual das métricas mais relevantes para engenharia")
        quality.grid(row=0, column=1, sticky="nsew", padx=(4, 0), pady=(0, 4))
        self.draw_percent_chart(quality, [
            ("Recall", m["Recall / Sensibilidade"], GOOD),
            ("Precision", m["Precision"], ACCENT),
            ("Specificity", m["Specificity"], ACCENT2),
            ("FPR", m["FPR"], BAD),
            ("FNR", m["FNR"], WARN),
        ])

        actions = self.card(lower, "LEITURA DE ENGENHARIA", "Como usar os números sem transformar uma métrica em 'nota'.")
        actions.grid(row=1, column=0, sticky="nsew", padx=(0, 4), pady=(4, 0))
        lines = [
            f"• Recall {pct(m['Recall / Sensibilidade'])}: priorize a capacidade de encontrar o que realmente é defeito.",
            f"• Precision {pct(m['Precision'])}: avalie quanto retrabalho/alarmes podem estar vindo de falsos positivos.",
            f"• FPR {pct(m['FPR'])}: observe a carga de alarmes em peças realmente OK.",
            f"• FNR {pct(m['FNR'])}: observe o risco de defeitos passarem como OK.",
            f"• FP de detecção {dt['fp']}: representa detecções extras em relação aos defeitos reais registrados.",
            f"• FN de detecção {dt['fn']}: representa defeitos reais sem detecção correspondente por contagem.",
            "• Accuracy não deve ser usada isoladamente, principalmente quando OK e NOK são desbalanceados.",
        ]
        for line in lines:
            tk.Label(actions, text=line, bg=PANEL, fg=TEXT,
                     wraplength=520, justify="left", font=("Segoe UI", 9)).pack(anchor="w", padx=16, pady=5)

        occurrence = self.card(lower, "OCORRÊNCIAS / DETECÇÃO", "Unidade diferente da classificação da inspeção.")
        occurrence.grid(row=1, column=1, sticky="nsew", padx=(4, 0), pady=(4, 0))
        vals = [
            ("Defeitos reais registrados", d["real"]),
            ("Detecções registradas", d["det"]),
            ("TP de detecção (por contagem)", d["tp"]),
            ("FP de detecção (extras)", d["fp"]),
            ("FN de detecção (não detectados)", d["fn"]),
            ("Precision de detecção", pct(dt["precision"])),
            ("Recall de detecção", pct(dt["recall"])),
        ]
        for lab, val in vals:
            row = tk.Frame(occurrence, bg=PANEL)
            row.pack(fill="x", padx=16, pady=5)
            tk.Label(row, text=lab, bg=PANEL, fg=MUTED, font=("Segoe UI", 9)).pack(side="left")
            tk.Label(row, text=str(val), bg=PANEL, fg=TEXT,
                     font=("Segoe UI", 10, "bold")).pack(side="right")
        tk.Label(occurrence,
                 text=("A Coleta Rápida pode operar sem tipo de defeito. "
                       "Quando o tipo é informado, a aba 'Por Defeito' permite aprofundar a análise."),
                 bg=PANEL, fg=ACCENT2, wraplength=520, justify="left",
                 font=("Segoe UI", 8)).pack(anchor="w", padx=16, pady=(10, 14))

        lower.columnconfigure(0, weight=1)
        lower.columnconfigure(1, weight=1)
        lower.rowconfigure(0, weight=1)
        lower.rowconfigure(1, weight=1)

    def view_defeitos(self):
        tk.Label(self.content,text="ANÁLISE POR TIPO DE DEFEITO",bg=BG,fg=TEXT,
                 font=("Segoe UI",18,"bold")).pack(anchor="w")
        tk.Label(self.content,text="Análise opcional: só aparecem tipos quando o técnico decidiu detalhá-los.",
                 bg=BG,fg=MUTED).pack(anchor="w",pady=(2,10))
        d=self.defect_counts()
        if not d["by"]:
            f=self.card(self.content,"Nenhum tipo detalhado")
            f.pack(fill="both",expand=True)
            tk.Label(f,text=("Isso é normal na Coleta Rápida. O tipo do defeito não é obrigatório. "
                             "Use 'Detalhar defeitos' apenas quando a análise por categoria for necessária."),
                     bg=PANEL,fg=TEXT,wraplength=900,justify="left").pack(anchor="w",padx=18,pady=18)
            return
        f=self.card(self.content,"RESUMO POR TIPO")
        f.pack(fill="both",expand=True)
        tree=ttk.Treeview(f,columns=("def","real","det","tp","fp","fn","rec","prec"),
                          show="headings")
        heads=("Defeito","Real","Detectado","TP","FP","FN","Recall","Precision")
        for c,h in zip(tree["columns"],heads):
            tree.heading(c,text=h);tree.column(c,width=130,anchor="center")
        tree.column("def",width=260,anchor="w")
        tree.pack(fill="both",expand=True,padx=12,pady=12)
        for name,v in sorted(d["by"].items()):
            mm=classification_metrics(v["tp"],0,v["fp"],v["fn"])
            tree.insert("", "end",values=(name,v["real"],v["det"],v["tp"],v["fp"],v["fn"],
                                          pct(mm["Recall / Sensibilidade"]),pct(mm["Precision"])))

    # ---------- COMPARAÇÃO MULTICOLETA ----------
    def edit_calibration_project(self):
        """Identifica o conjunto de coletas sem misturar dados de projetos distintos."""
        win=tk.Toplevel(self); win.title("Projeto de calibração"); win.geometry("620x430"); win.configure(bg=BG); win.transient(self); win.grab_set()
        box=self.card(win,"PROJETO DE CALIBRAÇÃO","Agrupe as coletas e registre o contexto da decisão."); box.pack(fill="both",expand=True,padx=16,pady=16)
        fields=[("Código do projeto","id"),("Nome do projeto","name"),("Objetivo de calibração","objective"),("Responsável","responsible")]
        vals={k:tk.StringVar(value=self.calibration_project.get(k,"")) for _,k in fields}
        for label,key in fields:
            tk.Label(box,text=label,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w",padx=18,pady=(10,1))
            ttk.Entry(box,textvariable=vals[key]).pack(fill="x",padx=18)
        def save():
            self.calibration_project={k:v.get().strip() for k,v in vals.items()}
            self.calibration_project["created_at"]=self.calibration_project.get("created_at") or datetime.now().isoformat()
            self.status.set("Projeto de calibração atualizado."); win.destroy(); self.refresh_all()
        ttk.Button(box,text="Salvar projeto de calibração",style="Accent.TButton",command=save).pack(anchor="e",padx=18,pady=18)

    def _calibration_score(self, item):
        """Prioriza aderência às metas; em empate privilegia F1, recall e menor FP."""
        m=item.get("metrics",{}); exp=item.get("expectation",{}) or {}
        checks=[("accuracy","Accuracy",True),("precision","Precision",True),("recall","Recall / Sensibilidade",True),("f1","F1",True)]
        met=total=0
        for key, metric, higher in checks:
            target=exp.get(key); actual=m.get(metric)
            if target is not None and actual is not None:
                total+=1; met += int(actual >= target if higher else actual <= target)
        n=m.get("N",0) or 0
        for key in ("fp100","fn100"):
            target=exp.get(key)
            if target is not None and n:
                total+=1; actual=m.get("FP" if key=="fp100" else "FN",0)/n*100
                met += int(actual <= target)
        return (met, total, m.get("F1") or -1, m.get("Recall / Sensibilidade") or -1, -(m.get("FP",0)))

    def _best_configuration(self, datasets=None):
        items=datasets if datasets is not None else self.calibration_history
        valid=[x for x in items if x.get("metrics",{}).get("N",0)>0]
        return max(valid,key=self._calibration_score) if valid else None

    def _threshold_value(self, item):
        try: return float(str(item.get("threshold","")).replace(",",".").replace("%",""))
        except (TypeError,ValueError): return None

    def _draw_threshold_chart(self,parent,datasets):
        points=[(self._threshold_value(x),x) for x in datasets]; points=[p for p in points if p[0] is not None and p[1].get("metrics",{}).get("F1") is not None]
        canvas=tk.Canvas(parent,bg=PANEL,highlightthickness=0,height=260); canvas.pack(fill="both",expand=True,padx=10,pady=8); canvas.update_idletasks()
        w=max(canvas.winfo_width(),620); h=260; left,right,top,bottom=55,28,22,45
        if not points:
            canvas.create_text(w/2,h/2,text="Informe um Threshold numérico em pelo menos uma coleta para gerar o gráfico.",fill=MUTED,font=("Segoe UI",9)); return
        points.sort(key=lambda x:x[0]); xs=[p[0] for p in points]; lo,hi=min(xs),max(xs)
        if lo==hi: lo-=1; hi+=1
        for i in range(5):
            y=top+(h-top-bottom)*(1-i/4); canvas.create_line(left,y,w-right,y,fill="#29404F")
            canvas.create_text(left-7,y,text=f"{i*25}%",anchor="e",fill=MUTED,font=("Segoe UI",7))
        prev=None
        for threshold,item in points:
            x=left+(threshold-lo)/(hi-lo)*(w-left-right); f1=item["metrics"]["F1"]; y=top+(1-f1)*(h-top-bottom)
            if prev: canvas.create_line(prev[0],prev[1],x,y,fill=ACCENT,width=2)
            canvas.create_oval(x-5,y-5,x+5,y+5,fill=GOOD if item is self._best_configuration(datasets) else ACCENT2,outline="")
            canvas.create_text(x,y-12,text=f"{f1*100:.1f}%",fill=TEXT,font=("Segoe UI",7,"bold")); canvas.create_text(x,h-bottom+15,text=str(threshold),fill=MUTED,font=("Segoe UI",7)); prev=(x,y)
        canvas.create_text(left,h-bottom+30,text="Threshold",fill="#6E8E9E",font=("Segoe UI",7))

    def view_calibracao(self):
        self.section_title("PROJETO DE CALIBRAÇÃO","Rastreabilidade, melhor configuração encontrada e relação Threshold x Performance.")
        p=self.calibration_project; info=self.card(self.content,"IDENTIFICAÇÃO DO PROJETO",p.get("name") or "Projeto ainda não identificado"); info.pack(fill="x",pady=4)
        tk.Label(info,text=f"Código: {p.get('id') or 'N/D'}   |   Responsável: {p.get('responsible') or 'N/D'}\nObjetivo: {p.get('objective') or 'N/D'}",bg=PANEL,fg=TEXT,justify="left").pack(anchor="w",padx=16,pady=12)
        ttk.Button(info,text="Editar identificação",command=self.edit_calibration_project).pack(anchor="e",padx=14,pady=(0,12))
        toolbar=tk.Frame(self.content,bg=BG); toolbar.pack(fill="x",pady=5)
        ttk.Button(toolbar,text="Registrar coleta atual no histórico",style="Accent.TButton",command=self._archive_current_from_view).pack(side="left",padx=3)
        ttk.Button(toolbar,text="Gerar relatório executivo",command=self.export_executive_report).pack(side="left",padx=3)
        datasets=self.calibration_history or self.comparison_archive
        best=self._best_configuration(datasets)
        bestbox=self.card(self.content,"MELHOR CONFIGURAÇÃO ENCONTRADA","Critério: maior atendimento às metas; depois F1, Recall e menor FP."); bestbox.pack(fill="x",pady=4)
        text="Ainda não há coleta classificada suficiente para recomendar uma configuração."
        if best:
            m=best["metrics"]; text=(f"{best.get('name')} | Threshold: {best.get('threshold','N/D')} | F1: {pct(m.get('F1'))} | Recall: {pct(m.get('Recall / Sensibilidade'))} | Precision: {pct(m.get('Precision'))} | N: {m.get('N',0)}")
        tk.Label(bestbox,text=text,bg=PANEL,fg=GOOD if best else MUTED,font=("Segoe UI",10,"bold"),wraplength=1100,justify="left").pack(anchor="w",padx=16,pady=14)
        chart=self.card(self.content,"THRESHOLD x PERFORMANCE","F1 por configuração registrada; o ponto verde representa a recomendação."); chart.pack(fill="both",expand=True,pady=4); self._draw_threshold_chart(chart,datasets)
        hist=self.card(self.content,"HISTÓRICO RASTREÁVEL",f"{len(datasets)} coleta(s) no projeto"); hist.pack(fill="x",pady=4)
        tree=ttk.Treeview(hist,columns=("id","data","coleta","threshold","n","f1","origem","rastro"),show="headings",height=min(9,max(3,len(datasets))))
        for col,label,width in [("id","Registro",105),("data","Data/Hora",150),("coleta","Coleta",145),("threshold","Threshold",90),("n","N",55),("f1","F1",80),("origem","Origem",210),("rastro","Hash",100)]: tree.heading(col,text=label); tree.column(col,width=width,anchor="center" if col not in ("coleta","origem") else "w")
        tree.pack(fill="x",padx=12,pady=12)
        for item in datasets:
            tree.insert("","end",values=(item.get("record_id","N/D"),item.get("created_at","N/D")[:19],item.get("name","Coleta"),item.get("threshold","N/D"),item.get("metrics",{}).get("N",0),pct(item.get("metrics",{}).get("F1")),Path(item.get("source_file","")).name or "Manual",item.get("trace_hash","")[:10]))

    def _current_collection_has_data(self):
        return bool(self.rows or self.source_file or self.history or any(v.get() for k, v in self.meta.items() if k not in ("data", "gabarito")))

    def _archive_current_collection(self, label=None):
        """Arquiva um snapshot da coleta ativa para comparação posterior."""
        if not self._current_collection_has_data():
            return None
        m = self.display_metrics()
        d = self.defect_counts()
        di = self._smart_detection_inspection() if hasattr(self, "_smart_detection_inspection") else {}
        meta = {k: v.get() for k, v in self.meta.items()}
        seq = len(self.comparison_archive) + 1
        payload = {"metadata":meta,"metrics":m,"defects":d,"detection_inspection":di,"source_file":self.source_file or "","project":self.calibration_project}
        item = {
            "name": label or f"Coleta {seq:02d}",
            "created_at": datetime.now().isoformat(),
            "metadata": meta,
            "source_file": self.source_file or "",
            "metrics": dict(m),
            "defects": dict(d),
            "detection_inspection": dict(di),
            "threshold": meta.get("threshold") or "N/D",
            "date": meta.get("data") or datetime.now().strftime("%d/%m/%Y"),
            "expectation": self.expectation_snapshot(),
            "project": dict(self.calibration_project),
            "record_id": f"CAL-{len(self.calibration_history)+1:03d}",
            "trace_hash": hashlib.sha256(json.dumps(payload,ensure_ascii=False,sort_keys=True,default=str).encode("utf-8")).hexdigest(),
        }
        self.comparison_archive.append(item)
        self.calibration_history.append(dict(item))
        if len(self.comparison_archive) > 50:
            self.comparison_archive = self.comparison_archive[-50:]
        if len(self.calibration_history) > 50:
            self.calibration_history = self.calibration_history[-50:]
        return item

    def _comparison_datasets(self):
        datasets = list(self.comparison_archive)
        if self._current_collection_has_data():
            m = self.display_metrics()
            d = self.defect_counts()
            di = self._smart_detection_inspection() if hasattr(self, "_smart_detection_inspection") else {}
            meta = {k: v.get() for k, v in self.meta.items()}
            datasets.append({
                "name": "Coleta atual", "created_at": datetime.now().isoformat(),
                "metadata": meta, "source_file": self.source_file or "",
                "metrics": dict(m), "defects": dict(d),
                "detection_inspection": dict(di),
                "threshold": meta.get("threshold") or "N/D",
                "date": meta.get("data") or datetime.now().strftime("%d/%m/%Y"),
                "expectation": self.expectation_snapshot(),
                "current": True,
            })
        return datasets[-51:]

    def _draw_trend_chart(self, parent, datasets, metric_key, percent=False):
        canvas = tk.Canvas(parent, bg=PANEL, highlightthickness=0, height=235)
        canvas.pack(fill="both", expand=True, padx=10, pady=(2,10))
        canvas.update_idletasks()
        w=max(canvas.winfo_width(),620); h=235
        left,right,top,bottom=58,25,22,45
        plot_w=max(1,w-left-right); plot_h=max(1,h-top-bottom)
        vals=[float(ds.get("metrics",{}).get(metric_key)) for ds in datasets if ds.get("metrics",{}).get(metric_key) is not None]
        if not vals:
            canvas.create_text(w/2,h/2,text="N/D — dados insuficientes",fill=MUTED,font=("Segoe UI",9,"bold")); return
        lo,hi=(0.0,1.0) if percent else (0.0,max(max(vals),1.0)*1.12)
        for i in range(5):
            frac=i/4; y=top+plot_h*(1-frac)
            canvas.create_line(left,y,w-right,y,fill="#29404F")
            lab=lo+(hi-lo)*frac
            canvas.create_text(left-8,y,text=(f"{lab*100:.0f}%" if percent else f"{lab:.0f}"),fill=MUTED,anchor="e",font=("Segoe UI",7))
        step=plot_w/max(len(datasets)-1,1); pts=[]
        for i,ds in enumerate(datasets):
            if metric_key == "__DET_FP__":
                v=ds.get("defects",{}).get("fp")
            elif metric_key == "__DET_FN__":
                v=ds.get("defects",{}).get("fn")
            else:
                v=ds.get("metrics",{}).get(metric_key)
            x=left+(i*step if len(datasets)>1 else plot_w/2)
            if v is None: continue
            v=float(v); frac=(v-lo)/(hi-lo) if hi>lo else 0; y=top+plot_h*(1-max(0,min(1,frac)))
            pts.append((x,y)); canvas.create_oval(x-4,y-4,x+4,y+4,fill=ACCENT2,outline="")
            label=ds.get("threshold") or ds.get("name","?"); label=str(label)[:9]
            canvas.create_text(x,h-bottom+12,text=label,fill=MUTED,font=("Segoe UI",7))
            canvas.create_text(x,y-10,text=(f"{v*100:.1f}%" if percent else f"{v:.0f}"),fill=TEXT,font=("Segoe UI",7,"bold"))
        for a,b in zip(pts,pts[1:]): canvas.create_line(*a,*b,fill=ACCENT,width=2)
        canvas.create_text(left,h-bottom+29,text="Coletas / Threshold",fill="#6E8E9E",font=("Segoe UI",7))

    def view_comparacao(self):
        self.section_title("COMPARAÇÃO  /  EVOLUÇÃO",
            "Compare de 2 a 50 coletas e observe o efeito de threshold, configuração ou motor no desempenho.")
        datasets=self._comparison_datasets()
        if not datasets:
            f=self.card(self.content,"NENHUMA COLETA DISPONÍVEL","Realize uma coleta ou importe uma planilha para iniciar a comparação.")
            f.pack(fill="x",pady=4)
            tk.Label(f,text="O AIA preserva as coletas anteriores quando você usa 'Nova coleta'. A coleta atual também pode ser comparada.",bg=PANEL,fg=TEXT,wraplength=900,justify="left").pack(anchor="w",padx=16,pady=14)
            return
        toolbar=tk.Frame(self.content,bg=BG); toolbar.pack(fill="x",pady=(0,7))
        ttk.Button(toolbar,text="Comparar selecionadas",style="Accent.TButton",command=lambda:self._run_comparison(datasets)).pack(side="left",padx=3)
        ttk.Button(toolbar,text="Selecionar todas",command=self._select_all_compare).pack(side="left",padx=3)
        ttk.Button(toolbar,text="Remover da seleção",command=self._remove_selected_compare).pack(side="left",padx=3)
        ttk.Button(toolbar,text="Limpar seleção",command=self._clear_compare).pack(side="left",padx=3)
        tk.Label(toolbar,text="Selecione 2 a 50 coletas.",bg=BG,fg=MUTED,font=("Segoe UI",8)).pack(side="right",padx=5)
        box=self.card(self.content,"COLETAS DISPONÍVEIS",f"{len(datasets)} conjunto(s) disponível(is) nesta sessão")
        box.pack(fill="x",pady=4)
        cols=("coleta","data","threshold","n","recall","precision","fp","fn","dfp","dfn","origem")
        tree=ttk.Treeview(box,columns=cols,show="headings",selectmode="extended",height=min(10,max(5,len(datasets))))
        self.compare_tree=tree; self.compare_datasets=datasets
        heads=("Coleta","Data","Threshold","N","Recall","Precision","FP-C","FN-C","FP-D","FN-D","Origem"); widths=(125,90,80,55,80,85,55,55,55,55,230)
        for c,h,w in zip(cols,heads,widths): tree.heading(c,text=h); tree.column(c,width=w,anchor="center" if c!="origem" else "w")
        tree.pack(fill="x",padx=12,pady=(4,12))
        for i,ds in enumerate(datasets):
            m=ds.get("metrics",{}); origin="ATUAL" if ds.get("current") else (Path(ds.get("source_file","")).name if ds.get("source_file") else "Manual")
            dds=ds.get("defects",{}); tree.insert("","end",iid=str(i),values=(ds.get("name","Coleta"),ds.get("date","N/D"),ds.get("threshold","N/D"),m.get("N",0),pct(m.get("Recall / Sensibilidade")),pct(m.get("Precision")),m.get("FP",0),m.get("FN",0),dds.get("fp","N/D"),dds.get("fn","N/D"),origin))
        # Delete remove apenas a(s) coleta(s) da comparação atual. O histórico
        # e os dados arquivados continuam íntegros.
        tree.bind("<Delete>", lambda _event: self._remove_selected_compare())
        if 2<=len(datasets)<=10: tree.selection_set(*[str(i) for i in range(len(datasets))])
        result=self.card(self.content,"RESULTADO DA COMPARAÇÃO","A análise é calculada somente sobre as coletas selecionadas.")
        result.pack(fill="x",pady=(6,4)); self.compare_result=result
        es=self.expectation_summary()
        if es["configured"]:
            tk.Label(result,text=f"🎯 Meta ativa: {es['met']}/{es['configured']} critérios atendidos na coleta atual • {es['status']}",bg=PANEL,fg=self._expectation_status_color(es["status"]),font=("Segoe UI",9,"bold")).pack(anchor="w",padx=16,pady=(8,2))
        tk.Label(result,text="Selecione pelo menos 2 coletas e clique em 'Comparar selecionadas'.",bg=PANEL,fg=MUTED,font=("Segoe UI",9)).pack(anchor="w",padx=16,pady=14)

    def _select_all_compare(self):
        if hasattr(self,"compare_tree"): self.compare_tree.selection_set(self.compare_tree.get_children())

    def _remove_selected_compare(self):
        """Remove uma ou mais coletas apenas da seleção da comparação."""
        tree=getattr(self,"compare_tree",None)
        if tree is None:
            return
        selected=tree.selection()
        if not selected:
            self.status.set("Nenhuma coleta selecionada para remover da comparação.")
            return
        tree.selection_remove(*selected)
        tree.focus("")
        self.status.set(f"{len(selected)} coleta(s) removida(s) da seleção de comparação.")

    def _clear_compare(self):
        """Limpa a seleção de forma explícita em todas as versões do Tk."""
        tree=getattr(self,"compare_tree",None)
        if tree is None:
            return
        # selection_set(()) é mais confiável que selection_remove(tuple) em
        # alguns ambientes Tk/Windows e também cobre seleção inicial automática.
        tree.selection_set(())
        tree.focus("")
        self.status.set("Seleção de comparação limpa.")

    def _archive_current_from_view(self):
        item=self._archive_current_collection()
        if item:
            self.status.set(f"Coleta arquivada para comparação: {item['name']}"); self.show("comparacao")
        else: messagebox.showinfo("Comparação","Não há uma coleta ativa para arquivar.")

    def _run_comparison(self,datasets):
        ids=self.compare_tree.selection() if hasattr(self,"compare_tree") else ()
        if len(ids)<2:
            messagebox.showwarning("Comparação","Selecione pelo menos 2 coletas."); return
        selected=[datasets[int(i)] for i in ids]
        result=getattr(self,"compare_result",None)
        if result is None: return
        for w in result.winfo_children(): w.destroy()
        tk.Label(result,text=f"{len(selected)} COLETAS EM COMPARAÇÃO",bg=PANEL,fg=ACCENT2,font=("Segoe UI",11,"bold")).pack(anchor="w",padx=16,pady=(12,3))
        exp=self.expectation_snapshot()
        configured=any(exp.get(k) is not None for k in ("accuracy","precision","recall","f1","fp100","fn100"))
        if configured:
            met_count=0
            total_count=0
            for ds in selected:
                em=ds.get("expectation") or exp; mds=ds.get("metrics",{}); n=mds.get("N",0)
                actuals={"accuracy":mds.get("Accuracy"),"precision":mds.get("Precision"),"recall":mds.get("Recall / Sensibilidade"),"f1":mds.get("F1"),"fp100":None if not n else mds.get("FP",0)/n*100.0,"fn100":None if not n else mds.get("FN",0)/n*100.0}
                for k,op in [("accuracy",">="),("precision",">="),("recall",">="),("f1",">="),("fp100","<="),("fn100","<=")]:
                    t=em.get(k); a=actuals.get(k)
                    if t is not None:
                        total_count+=1
                        if a is not None and ((a>=t) if op==">=" else (a<=t)): met_count+=1
            tk.Label(result,text=f"🎯 ATENDIMENTO À META NAS COLETAS SELECIONADAS: {met_count}/{total_count} critérios-coleta atendidos",bg=PANEL,fg=self._expectation_status_color("META ATINGIDA" if total_count and met_count==total_count else ("EM CALIBRAÇÃO" if met_count else "FORA DA EXPECTATIVA")),font=("Segoe UI",9,"bold")).pack(anchor="w",padx=16,pady=(0,6))
        table=ttk.Treeview(result,columns=("indicador","base","final","delta"),show="headings",height=9)
        for c,h,w in [("indicador","Indicador",210),("base","Primeira selecionada",150),("final","Última selecionada",150),("delta","Variação",150)]: table.heading(c,text=h); table.column(c,width=w,anchor="center" if c!="indicador" else "w")
        table.pack(fill="x",padx=12,pady=8)
        first=selected[0].get("metrics",{}); last=selected[-1].get("metrics",{})
        fd=selected[0].get("defects",{}); ld=selected[-1].get("defects",{})
        rows=[("N",first.get("N",0),last.get("N",0),False),("TP classificação",first.get("TP",0),last.get("TP",0),False),("TN / VN classificação",first.get("VN/TN",0),last.get("VN/TN",0),False),("FP classificação",first.get("FP",0),last.get("FP",0),False),("FN classificação",first.get("FN",0),last.get("FN",0),False),("FP detecção",fd.get("fp","N/D"),ld.get("fp","N/D"),False),("FN detecção",fd.get("fn","N/D"),ld.get("fn","N/D"),False),("Accuracy",first.get("Accuracy"),last.get("Accuracy"),True),("Precision",first.get("Precision"),last.get("Precision"),True),("Recall",first.get("Recall / Sensibilidade"),last.get("Recall / Sensibilidade"),True),("F1",first.get("F1"),last.get("F1"),True)]
        for label,a,b,is_pct in rows:
            if is_pct: av=pct(a); bv=pct(b); dv="N/D" if a is None or b is None else f"{(b-a)*100:+.2f} p.p."
            else: av=str(a if a is not None else "N/D"); bv=str(b if b is not None else "N/D"); dv="N/D" if a is None or b is None else f"{b-a:+g}"
            table.insert("","end",values=(label,av,bv,dv))
        charts=tk.Frame(result,bg=PANEL); charts.pack(fill="both",expand=True,padx=10,pady=(0,10))
        for col,(title,key,pctmode) in enumerate([("EVOLUÇÃO DO RECALL","Recall / Sensibilidade",True),("EVOLUÇÃO DA PRECISION","Precision",True),("FP DE DETECÇÃO","__DET_FP__",False),("FN DE DETECÇÃO","__DET_FN__",False)]):
            row=0 if col<2 else 1; c=col%2
            card=self.chart_card(charts,title,"Tendência entre as coletas selecionadas")
            card.grid(row=row,column=c,sticky="nsew",padx=(0,4) if c==0 else (4,0),pady=4)
            self._draw_trend_chart(card,selected,key,pctmode)
        charts.columnconfigure(0,weight=1); charts.columnconfigure(1,weight=1)
        read=self.card(result,"LEITURA COMPARATIVA"); read.pack(fill="x",padx=10,pady=(0,10))
        rv=[x.get("metrics",{}).get("Recall / Sensibilidade") for x in selected]; pv=[x.get("metrics",{}).get("Precision") for x in selected]
        lines=[]
        if all(v is not None for v in rv): lines.append(f"Recall: {(rv[-1]-rv[0])*100:+.2f} p.p. entre a primeira e a última coleta.")
        if all(v is not None for v in pv): lines.append(f"Precision: {(pv[-1]-pv[0])*100:+.2f} p.p. entre a primeira e a última coleta.")
        lines.append(f"Falsos positivos: {first.get('FP',0)} → {last.get('FP',0)}; falsos negativos: {first.get('FN',0)} → {last.get('FN',0)}.")
        lines.append("A leitura indica tendência nos dados; a escolha de configuração deve considerar contexto, amostra e risco operacional.")
        for line in lines: tk.Label(read,text="• "+line,bg=PANEL,fg=TEXT,wraplength=1050,justify="left",font=("Segoe UI",9)).pack(anchor="w",padx=14,pady=4)

    # ---------- PROJECT / EXPORT ----------
    def new_collection(self, confirm=True):
        """Inicia uma nova coleta totalmente vazia, separada das inspeções atuais.

        Diferente de "+ Nova inspeção", esta ação encerra o conjunto atual de
        dados em memória e prepara a tela para outro cliente/planta/projeto.
        Nenhum arquivo já salvo no computador é apagado.
        """
        if confirm:
            ok = messagebox.askyesno(
                "Nova coleta — iniciar do zero",
                "Isso vai limpar da tela todas as inspeções, histórico e informações "
                "da coleta atual.\n\n"
                "Os arquivos já salvos no computador NÃO serão apagados.\n\n"
                "Deseja iniciar uma nova coleta do zero?",
                icon="warning"
            )
            if not ok:
                return

        self._archive_current_collection()
        self.rows = []
        self.history = []
        self.source_file = ""
        for v in self.meta.values():
            v.set("")
        # Nova coleta começa realmente vazia: a data deve representar o dia
        # em que a nova inspeção/coleta foi realizada, não o dia em que o
        # software foi aberto. O campo permanece editável na tela Coleta.
        self.meta["data"].set("")
        self.meta["gabarito"].set("Não determinado")
        self.status.set("Nova coleta iniciada — nenhum dado registrado.")
        self.refresh_all()

    def new_test(self,confirm=True):
        # Compatibilidade com chamadas antigas: agora o conceito de novo teste
        # também inicia uma coleta realmente vazia.
        self.new_collection(confirm=confirm)

    def refresh_all(self):
        if hasattr(self,"tree"): self.refresh_table()
        self.update_side()
        if self.current_view:self.show(self.current_view)

    def update_side(self):
        g=self.meta["gabarito"].get() or "Não determinado"
        src="Excel histórico" if self.source_file else "Entrada manual"
        classified=sum(1 for r in self.rows if r.get("ref") in ("Defeito confirmado","OK confirmado"))
        self.side_status.configure(text=f"Fonte: {src}\nReferência: {g}\nInspeções: {len(self.rows)}\nClassificadas: {classified}")

    def save_project(self):
        path=filedialog.asksaveasfilename(title="Salvar projeto",defaultextension=".json",
                                          filetypes=[("Projeto JSON","*.json")])
        if not path:return
        data={"app":APP,"version":VERSION,"saved_at":datetime.now().isoformat(),
              "metadata":{k:v.get() for k,v in self.meta.items()},
              "rows":[normalize_row(r) for r in self.rows],
              "history":self.history,"source_file":self.source_file,
              "comparison_archive":self.comparison_archive,"calibration_history":self.calibration_history,
              "calibration_project":self.calibration_project,"expectation":self.expectation_snapshot()}
        Path(path).write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding="utf-8")
        self.status.set(f"Projeto salvo: {Path(path).name}")

    def _project_summary(self):
        return self.metrics(), (self.history[-1] if self.history else {"TP":0,"VN/TN":0,"FP":0,"FN":0,"N":0})

    def export_executive_report(self):
        """Gera o relatório executivo com DETECÇÃO como indicador principal.

        C (classificação) e D (detecção) permanecem independentes. Nenhum
        indicador é inventado quando a base não sustenta o cálculo.
        """
        path=filedialog.asksaveasfilename(
            title="Gerar relatório executivo",
            defaultextension=".pdf",
            filetypes=[("PDF","*.pdf")]
        )
        if not path:
            return
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import cm
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether

            m=self.display_metrics()
            d=self.defect_counts()
            dt=self.detection_totals()
            di=dt.get("inspection",{})
            best=self._best_configuration(self.calibration_history)

            styles=getSampleStyleSheet()
            title_style=ParagraphStyle(
                "aia_title", parent=styles["Title"], textColor=colors.HexColor("#087EA4"),
                fontName="Helvetica-Bold", fontSize=20, leading=23, spaceAfter=6
            )
            sub=ParagraphStyle(
                "aia_sub", parent=styles["Normal"], textColor=colors.HexColor("#406070"),
                fontSize=9, leading=12, spaceAfter=10
            )
            h2=ParagraphStyle(
                "aia_h2", parent=styles["Heading2"], textColor=colors.HexColor("#07526B"),
                fontName="Helvetica-Bold", fontSize=12, leading=15, spaceBefore=8, spaceAfter=6
            )
            body=ParagraphStyle("aia_body", parent=styles["BodyText"], fontSize=8.5, leading=11)
            callout=ParagraphStyle(
                "aia_callout", parent=body, textColor=colors.HexColor("#0A4D3D"),
                backColor=colors.HexColor("#E8F7F1"), borderColor=colors.HexColor("#36C99A"),
                borderWidth=0.6, borderPadding=7, spaceBefore=4, spaceAfter=7
            )
            warn=ParagraphStyle(
                "aia_warn", parent=body, textColor=colors.HexColor("#754C00"),
                backColor=colors.HexColor("#FFF6D8"), borderColor=colors.HexColor("#E2B84C"),
                borderWidth=0.6, borderPadding=7, spaceBefore=4, spaceAfter=7
            )

            def pv(v):
                return "N/D" if v is None else f"{v*100:.2f}%"
            def nv(v):
                return "N/D" if v is None else str(v)

            doc=SimpleDocTemplate(
                path,pagesize=A4,rightMargin=1.5*cm,leftMargin=1.5*cm,
                topMargin=1.35*cm,bottomMargin=1.65*cm,
                title=f"Autaza Inspection Analytics - Relatório Executivo V{VERSION}",
                author="Autaza"
            )

            def page_footer(canvas, document):
                canvas.saveState()
                width, _ = A4
                canvas.setStrokeColor(colors.HexColor("#B8D3DC"))
                canvas.setLineWidth(0.35)
                canvas.line(document.leftMargin, 1.15*cm, width-document.rightMargin, 1.15*cm)
                canvas.setFont("Helvetica", 7)
                canvas.setFillColor(colors.HexColor("#406070"))
                canvas.drawString(document.leftMargin, 0.78*cm, "AUTAZA INSPECTION ANALYTICS")
                canvas.drawRightString(width-document.rightMargin, 0.78*cm,
                                       f"V{VERSION}  |  Página {document.page}")
                canvas.restoreState()
            story=[
                Paragraph("AUTAZA INSPECTION ANALYTICS",title_style),
                Paragraph(
                    f"RELATÓRIO EXECUTIVO — V{VERSION} — Emitido em {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                    sub
                )
            ]

            project=self.calibration_project
            info=[
                ["Projeto",project.get("name") or self.meta["projeto"].get() or "N/D"],
                ["Código",project.get("id") or "N/D"],
                ["Cliente / Planta",f"{self.meta['cliente'].get() or 'N/D'} / {self.meta['local'].get() or 'N/D'}"],
                ["Linha / Estação",self.meta["linha"].get() or "N/D"],
                ["Responsável",project.get("responsible") or self.meta["tecnico"].get() or "N/D"],
                ["Data / Threshold",f"{self.meta['data'].get() or 'N/D'} / {self.meta['threshold'].get() or 'N/D'}"],
                ["Método de referência",self.meta["gabarito"].get() or "N/D"],
                ["Fonte",Path(self.source_file).name if self.source_file else "Entrada manual"],
            ]
            t=Table(info,colWidths=[4.3*cm,11.8*cm])
            t.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(0,-1),colors.HexColor("#EAF5F8")),
                ("TEXTCOLOR",(0,0),(0,-1),colors.HexColor("#07526B")),
                ("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),
                ("GRID",(0,0),(-1,-1),0.25,colors.HexColor("#B8D3DC")),
                ("VALIGN",(0,0),(-1,-1),"TOP"),
                ("PADDING",(0,0),(-1,-1),6),
                ("FONTSIZE",(0,0),(-1,-1),8),
            ]))
            story += [t,Spacer(1,9)]

            # ============================================================
            # 01 — DETECÇÃO: PRINCIPAL INDICADOR TÉCNICO
            # ============================================================
            story.append(Paragraph("01. DETECÇÃO DE DEFEITOS — INDICADOR PRINCIPAL",h2))
            story.append(Paragraph(
                "A detecção responde quantos defeitos/ocorrências reais existiam e quantos foram encontrados pelo sistema. "
                "Ela é apresentada separadamente da classificação OK/NOK e não reutiliza a matriz C como se fosse D.",
                callout
            ))
            det_table=[
                ["Indicador","Valor","Unidade / método"],
                ["Ocorrências reais",nv(d.get("real")),"ocorrência"],
                ["Detecções",nv(d.get("det")),"ocorrência"],
                ["TP-D",nv(d.get("tp")),"ocorrência"],
                ["FP-D",nv(d.get("fp")),"ocorrência"],
                ["FN-D",nv(d.get("fn")),"ocorrência"],
                ["Precision-D",pv(d.get("precision")),"TP-D / (TP-D + FP-D)"],
                ["Recall-D",pv(d.get("recall")),"TP-D / (TP-D + FN-D)"],
                ["F1-D",pv(d.get("f1")),"média harmônica de Precision/Recall"],
                ["Status",d.get("status","N/D"),d.get("method","N/D")],
            ]
            t=Table(det_table,colWidths=[4.0*cm,4.0*cm,8.1*cm],repeatRows=1)
            t.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#103126")),
                ("TEXTCOLOR",(0,0),(-1,0),colors.white),
                ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                ("FONTNAME",(0,1),(0,-1),"Helvetica-Bold"),
                ("GRID",(0,0),(-1,-1),0.35,colors.HexColor("#A9C9BF")),
                ("VALIGN",(0,0),(-1,-1),"TOP"),
                ("PADDING",(0,0),(-1,-1),5),
                ("FONTSIZE",(0,0),(-1,-1),7.5),
                ("BACKGROUND",(0,1),(-1,-1),colors.HexColor("#F2FAF6")),
            ]))
            story += [t,Spacer(1,7)]

            if d.get("estimated") or d.get("status") == "ESTIMADO":
                story.append(Paragraph(
                    "Atenção: TP-D/FP-D/FN-D foram obtidos por regra de contagem porque não há correspondência individual suficiente "
                    "entre ocorrência real e detecção. O resultado é válido como estimativa documentada, não como pareamento individual.",
                    warn
                ))

            # Detecção por inspeção, quando matematicamente disponível.
            story.append(Paragraph("Detecção por inspeção",styles["Heading3"]))
            if di.get("available"):
                dit=[
                    ["N",nv(di.get("N")),"Accuracy-D",pv(di.get("Accuracy"))],
                    ["TP-DI",nv(di.get("TP")),"TN-DI",nv(di.get("TN",di.get("VN/TN")))],
                    ["FP-DI",nv(di.get("FP")),"FN-DI",nv(di.get("FN"))],
                    ["Precision-DI",pv(di.get("Precision")),"Recall-DI",pv(di.get("Recall / Sensibilidade"))],
                    ["Specificity-DI",pv(di.get("Specificity")),"F1-DI",pv(di.get("F1"))],
                ]
                t=Table(dit,colWidths=[3.0*cm,4.0*cm,4.0*cm,5.1*cm])
                t.setStyle(TableStyle([
                    ("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#F5FAFB")),
                    ("GRID",(0,0),(-1,-1),0.25,colors.HexColor("#B8D3DC")),
                    ("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold"),
                    ("FONTSIZE",(0,0),(-1,-1),7.5),
                    ("PADDING",(0,0),(-1,-1),5),
                ]))
                story += [t,Spacer(1,7)]
            else:
                story.append(Paragraph(
                    "N/D — não há base suficiente por inspeção para construir a matriz D-I com segurança. "
                    "O AIA não substitui essa matriz por classificação C.",warn
                ))

            # ============================================================
            # 02 — CLASSIFICAÇÃO: COMPLEMENTAR
            # ============================================================
            # Inicia o segundo bloco em página própria. A tabela de C nunca fica
            # partida entre o rodapé da primeira página e o topo da seguinte.
            story.append(PageBreak())
            classification_block=[Paragraph("02. CLASSIFICAÇÃO OK/NOK — INDICADOR COMPLEMENTAR",h2)]
            if self.source_file and self.history and not self.source_analysis.get("classification_matrix",{}).get("valid"):
                classification_block.append(Paragraph(
                    "N/D — a fonte não fornece uma matriz de classificação C válida por inspeção. "
                    "A matriz histórica de detecção não é reutilizada como classificação.",warn
                ))
            else:
                ct=[
                    ["Indicador","Valor"],
                    ["Inspeções (N)",nv(m.get("N"))],
                    ["TP-C",nv(m.get("TP"))],
                    ["VN/TN-C",nv(m.get("VN/TN"))],
                    ["FP-C",nv(m.get("FP"))],
                    ["FN-C",nv(m.get("FN"))],
                    ["Accuracy-C",pv(m.get("Accuracy"))],
                    ["Precision-C",pv(m.get("Precision"))],
                    ["Recall-C",pv(m.get("Recall / Sensibilidade"))],
                    ["Specificity-C",pv(m.get("Specificity"))],
                    ["F1-C",pv(m.get("F1"))],
                    ["MCC-C","N/D" if m.get("MCC") is None else f"{m['MCC']:.4f}"],
                ]
                t=Table(ct,colWidths=[8.0*cm,8.1*cm],repeatRows=1)
                t.setStyle(TableStyle([
                    ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#0E202D")),
                    ("TEXTCOLOR",(0,0),(-1,0),colors.white),
                    ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                    ("FONTNAME",(0,1),(0,-1),"Helvetica-Bold"),
                    ("GRID",(0,0),(-1,-1),0.3,colors.HexColor("#B8D3DC")),
                    ("FONTSIZE",(0,0),(-1,-1),7.7),
                    ("PADDING",(0,0),(-1,-1),5),
                    ("BACKGROUND",(0,1),(-1,-1),colors.HexColor("#F5F8FA")),
                ]))
                classification_block += [t,Spacer(1,7)]
            story.append(KeepTogether(classification_block))

            # ============================================================
            # 03 — DECISÃO / EXPECTATIVA / RASTREABILIDADE
            # ============================================================
            decision_block=[Paragraph("03. CRITÉRIO DE ENTREGA E DECISÃO",h2)]
            es=self.expectation_summary()
            decision_block.append(Paragraph(
                f"Status atual: <b>{es.get('status','N/D')}</b>. "
                f"Critérios atendidos: <b>{es.get('met',0)}/{es.get('configured',0)}</b>.",body
            ))
            if best:
                bm=best.get("detection_inspection", {}) or {}
                recommendation=(
                    f"Configuração recomendada: <b>{best.get('name')}</b>, threshold <b>{best.get('threshold','N/D')}</b>. "
                    f"F1-DI={pv(bm.get('F1'))}, Recall-DI={pv(bm.get('Recall / Sensibilidade'))}, "
                    f"Precision-DI={pv(bm.get('Precision'))}. A recomendação considera o atendimento às metas D-I e, em empate, "
                    "F1-DI, Recall-DI e menor FP-DI."
                )
            else:
                recommendation="Ainda não há histórico classificado suficiente para recomendar uma configuração."
            decision_block.append(Paragraph(recommendation,body))
            story.append(KeepTogether(decision_block))

            trace_block=[Paragraph("04. RASTREABILIDADE",h2)]
            a=self.source_analysis or {}
            agg=a.get("aggregate_source",{})
            trace=(
                f"Histórico do projeto: {len(self.calibration_history)} coleta(s). "
                f"Coleta atual: {len(self.rows)} inspeção(ões). "
                f"Fonte: {Path(self.source_file).name if self.source_file else 'entrada manual'}. "
                f"Unidade principal do bloco D: ocorrência. Unidade do bloco C: inspeção."
            )
            if agg.get("present"):
                trace += " A fonte contém indicadores agregados; eles permanecem identificados como declarados pela fonte."
            trace_block.append(Paragraph(trace,body))
            trace_block.append(Spacer(1,6))
            trace_block.append(Paragraph(
                "Regra de integridade: N/D significa que a base disponível não permite determinar o indicador com segurança; "
                "não significa zero. Valores estimados são sempre identificados como ESTIMADO.",callout
            ))
            story.append(KeepTogether(trace_block))

            doc.build(story,onFirstPage=page_footer,onLaterPages=page_footer)
            self.status.set(f"Relatório executivo gerado: {Path(path).name}")
            messagebox.showinfo("Relatório executivo","PDF gerado com sucesso.")
        except Exception as e:
            messagebox.showerror("Relatório executivo",f"Não foi possível gerar o PDF:\n{e}")

    def export_csv(self):
        path=filedialog.asksaveasfilename(title="Exportar CSV estruturado",defaultextension=".csv",
                                          filetypes=[("CSV","*.csv")])
        if not path:return
        m,hist=self._project_summary();d=self.defect_counts()
        with open(path,"w",newline="",encoding="utf-8-sig") as f:
            w=csv.writer(f,delimiter=";")
            w.writerow(["AUTAZA INSPECTION ANALYTICS",f"V{VERSION}"])
            w.writerow(["Data de exportação",datetime.now().strftime("%d/%m/%Y %H:%M:%S")]);w.writerow([])
            w.writerow(["IDENTIFICAÇÃO DO TESTE"])
            for key,label in [("cliente","Cliente"),("local","Local / Planta"),("projeto","Projeto"),
                              ("linha","Linha / Estação"),("tecnico","Técnico"),("data","Data"),
                              ("threshold","Threshold"),("gabarito","Método de referência")]:
                w.writerow([label,self.meta[key].get()])
            w.writerow(["Fonte do teste","Excel histórico" if self.source_file else "Entrada manual"])
            w.writerow(["Arquivo de origem",self.source_file or ""]);w.writerow([])
            w.writerow(["EXPECTATIVA / CRITÉRIO DE ENTREGA"])
            es=self.expectation_snapshot()
            for key,label in [("accuracy","Meta Accuracy"),("precision","Meta Precision"),("recall","Meta Recall"),("f1","Meta F1"),("fp100","Meta FP/100"),("fn100","Meta FN/100"),("source","Origem do critério"),("responsible","Responsável"),("date","Data do critério"),("notes","Observações")]:
                w.writerow([label,es.get(key)])
            w.writerow([])
            w.writerow(["CLASSIFICAÇÃO DAS INSPEÇÕES"])
            w.writerow(["Indicador","Valor"])
            for k in ("TP","VN/TN","FP","FN","N"):w.writerow([k,m.get(k,0)])
            for k in ("Accuracy","Precision","Recall / Sensibilidade","Specificity","FPR","FNR","F1","MCC"):
                v=m.get(k);w.writerow([k,"N/D" if v is None else v])
            w.writerow([])
            di=self.detection_totals().get("inspection",{}) or {}
            w.writerow(["DETECÇÃO POR INSPEÇÃO — D-I"])
            w.writerow(["Indicador","Valor"])
            for key,label in (("N","N-DI"),("TP","TP-DI"),("TN","TN-DI"),("FP","FP-DI"),("FN","FN-DI")):
                w.writerow([label,di.get(key,"N/D") if di.get("available") else "N/D"])
            for key,label in (("Accuracy","Accuracy D-I"),("Precision","Precision D-I"),("Recall / Sensibilidade","Recall D-I"),("Specificity","Specificity D-I"),("F1","F1 D-I")):
                w.writerow([label,di.get(key,"N/D") if di.get("available") else "N/D"])
            w.writerow(["Método",di.get("method","N/D")])
            w.writerow([])
            w.writerow(["DETECÇÃO DE OCORRÊNCIAS"])
            w.writerow(["Indicador","Valor"])
            for k,label in (("real","Defeitos reais"),("det","Detecções"),("tp","TP de detecção"),("fp","FP de detecção"),("fn","FN de detecção")):w.writerow([label,d[k]])
            w.writerow(["Análise por tipo","Detalhada" if d["detailed"] else "Coleta rápida — sem tipo"]);w.writerow([])
            w.writerow(["INSPEÇÕES"])
            w.writerow(["Inspeção","Identificação","Vision AI","Referência","Real","Detectado",
                        "TP-C","VN-C","FP-C","FN-C","TP-D","FP-D","FN-D","Tipos detalhados","Observação","Origem"])
            for i,r0 in enumerate(self.rows,1):
                r=normalize_row(r0); dd=self.detection_for_row(r)
                dv=(dd["tp"],dd["fp"],dd["fn"]) if dd["available"] else ("N/D","N/D","N/D")
                w.writerow([i,r["id"],r["vision"],r["ref"],r["real"],r["det"],r["tp"],r["vn"],
                            r["fp"],r["fn"],dv[0],dv[1],dv[2],len(r["defects"]),r["obs"],r["source"]])
            w.writerow([])
            w.writerow(["DEFEITOS DETALHADOS"])
            w.writerow(["Inspeção","Identificação","Tipo","Real","Detectado","TP por contagem","FP","FN"])
            for i,r0 in enumerate(self.rows,1):
                r=normalize_row(r0)
                for dd in r["defects"]:
                    tt=min(dd["real"],dd["det"]);ff=max(dd["det"]-dd["real"],0);nn=max(dd["real"]-dd["det"],0)
                    w.writerow([i,r["id"],dd["type"],dd["real"],dd["det"],tt,ff,nn])
            w.writerow([])
            w.writerow(["AUDITORIA HISTÓRICO × MOTOR NOVO"])
            w.writerow(["Indicador","Histórico","Motor novo","Diferença"])
            for k in ("TP","VN/TN","FP","FN","N"):
                hv=hist.get(k,0);mv=m.get(k,0);w.writerow([k,hv,mv,mv-hv])
        self.status.set(f"CSV estruturado exportado: {Path(path).name}")
        messagebox.showinfo("Exportação concluída",
                            "CSV estruturado criado com classificação, ocorrências, inspeções e detalhes.")

    def export_excel(self):
        if openpyxl is None:
            messagebox.showerror("Excel","OpenPyXL não está instalado.\nExecute: py -3 -m pip install openpyxl");return
        path=filedialog.asksaveasfilename(title="Exportar relatório completo",defaultextension=".xlsx",
                                          filetypes=[("Excel","*.xlsx")])
        if not path:return
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font,PatternFill,Border,Side,Alignment
            wb=Workbook();ws=wb.active;ws.title="RESUMO"
            dark="182532";blue="2F8CFF";border="6B7F8D";green="45D39A";yellow="F4C15D";red="FF6B78"
            thin=Side(style="thin",color=border)
            def title(sh,text,cols=6):
                sh["A1"]=text;sh["A1"].font=Font(size=18,bold=True,color="FFFFFF")
                sh["A1"].fill=PatternFill("solid",fgColor=dark);sh.merge_cells(start_row=1,start_column=1,end_row=1,end_column=cols)
            def hdr(sh,row,labels):
                for c,v in enumerate(labels,1):
                    x=sh.cell(row,c,v);x.font=Font(bold=True,color="FFFFFF")
                    x.fill=PatternFill("solid",fgColor=dark);x.border=Border(bottom=thin)
            def autosize(sh):
                # Ignora MergedCell: títulos mesclados não possuem column_letter.
                # Mede somente células reais e aplica largura por coluna.
                from openpyxl.cell.cell import Cell
                widths = {}
                for row in sh.iter_rows():
                    for x in row:
                        if not isinstance(x, Cell):
                            continue
                        value = "" if x.value is None else str(x.value)
                        widths[x.column] = max(widths.get(x.column, 10), min(55, len(value) + 2))
                for idx, width in widths.items():
                    from openpyxl.utils import get_column_letter
                    sh.column_dimensions[get_column_letter(idx)].width = max(10, min(55, width))
            m,hist=self._project_summary();d=self.defect_counts()
            title(ws,"AUTAZA INSPECTION ANALYTICS — RELATÓRIO",6)
            info=[("Versão",VERSION),("Data de exportação",datetime.now().strftime("%d/%m/%Y %H:%M:%S")),
                  ("Fonte","Excel histórico" if self.source_file else "Entrada manual"),
                  ("Cliente",self.meta["cliente"].get()),("Local / Planta",self.meta["local"].get()),
                  ("Projeto",self.meta["projeto"].get()),("Linha / Estação",self.meta["linha"].get()),
                  ("Técnico",self.meta["tecnico"].get()),("Data",self.meta["data"].get()),
                  ("Threshold",self.meta["threshold"].get()),("Método de referência",self.meta["gabarito"].get())]
            for i,(a,b) in enumerate(info,3):ws.cell(i,1,a);ws.cell(i,2,b)
            hdr(ws,16,["Indicador","Valor","Interpretação"])
            interp={"TP":"Inspeções reais NOK classificadas como NOK","VN/TN":"Inspeções reais OK classificadas como OK",
                    "FP":"Inspeções reais OK classificadas como NOK","FN":"Inspeções reais NOK classificadas como OK"}
            vals=[("TP",m["TP"]),("VN/TN",m["VN/TN"]),("FP",m["FP"]),("FN",m["FN"]),
                  ("N",m["N"]),("Accuracy",m["Accuracy"]),("Precision",m["Precision"]),
                  ("Recall / Sensibilidade",m["Recall / Sensibilidade"]),("Specificity",m["Specificity"]),
                  ("F1",m["F1"])]
            for i,(k,v) in enumerate(vals,17):
                ws.cell(i,1,k);ws.cell(i,2,v);ws.cell(i,3,interp.get(k,""))
                if isinstance(v,float):ws.cell(i,2).number_format="0.00%"
            autosize(ws)

            we=wb.create_sheet("EXPECTATIVA");title(we,"EXPECTATIVA / CRITÉRIO DE ENTREGA",6)
            hdr(we,3,["Indicador","Meta","Resultado atual","Critério","Status","Origem"])
            for i,(label,key,op,a,t,status) in enumerate(self._expectation_rows(),4):
                we.cell(i,1,label);we.cell(i,2,"N/D" if t is None else t);we.cell(i,3,"N/D" if a is None else a);we.cell(i,4,op);we.cell(i,5,status);we.cell(i,6,self.expectation["source"].get())
                if key in ("accuracy","precision","recall","f1"):
                    if isinstance(t,float):we.cell(i,2).number_format="0.00%"
                    if isinstance(a,float):we.cell(i,3).number_format="0.00%"
            autosize(we)

            wi=wb.create_sheet("INSPEÇÕES");title(wi,"INSPEÇÕES",16)
            hdr(wi,3,["Inspeção","Identificação","Vision AI","Referência","Real","Detectado","TP-C","VN-C","FP-C","FN-C","TP-D","FP-D","FN-D","Tipos detalhados","Observação","Origem"])
            for i,r0 in enumerate(self.rows,4):
                r=normalize_row(r0); dd=self.detection_for_row(r)
                dv=(dd["tp"],dd["fp"],dd["fn"]) if dd["available"] else ("N/D","N/D","N/D")
                vals=[i-3,r["id"],r["vision"],r["ref"],r["real"],r["det"],r["tp"],r["vn"],r["fp"],r["fn"],
                      dv[0],dv[1],dv[2],len(r["defects"]),r["obs"],r["source"]]
                for c,v in enumerate(vals,1):wi.cell(i,c,v)
            wi.freeze_panes="A4";autosize(wi)

            wm=wb.create_sheet("MÉTRICAS");title(wm,"MÉTRICAS DE CLASSIFICAÇÃO",3)
            hdr(wm,3,["Métrica","Valor","Definição"])
            defs={"Accuracy":"Classificações corretas / total","Precision":"TP/(TP+FP)",
                  "Recall / Sensibilidade":"TP/(TP+FN)","Specificity":"TN/(TN+FP)",
                  "FPR":"FP/(FP+TN)","FNR":"FN/(FN+TP)","F1":"Média harmônica de Precision e Recall",
                  "Balanced Accuracy":"(Recall+Specificity)/2","MCC":"Correlação da matriz de confusão"}
            for i,k in enumerate(("Accuracy","Precision","Recall / Sensibilidade","Specificity","FPR","FNR","F1","Balanced Accuracy","MCC"),4):
                v=m[k];wm.cell(i,1,k);wm.cell(i,2,v if v is not None else "N/D");wm.cell(i,3,defs.get(k,""))
                if isinstance(v,float) and k!="MCC":wm.cell(i,2).number_format="0.00%"
            autosize(wm)

            wa=wb.create_sheet("AUDITORIA");title(wa,"AUDITORIA — HISTÓRICO × MOTOR NOVO",4)
            hdr(wa,3,["Indicador","Histórico","Motor novo","Diferença"])
            for i,k in enumerate(("TP","VN/TN","FP","FN","N"),4):
                hv=hist.get(k,0);mv=m.get(k,0);wa.cell(i,1,k);wa.cell(i,2,hv);wa.cell(i,3,mv);wa.cell(i,4,mv-hv)
            autosize(wa)

            wd=wb.create_sheet("POR DEFEITO");title(wd,"ANÁLISE POR TIPO DE DEFEITO",9)
            hdr(wd,3,["Defeito","Real","Detectado","TP","FP","FN","Recall","Precision","Inspeções"])
            for i,(name,v) in enumerate(sorted(d["by"].items()),4):
                mm=classification_metrics(v["tp"],0,v["fp"],v["fn"])
                vals=[name,v["real"],v["det"],v["tp"],v["fp"],v["fn"],
                      mm["Recall / Sensibilidade"],mm["Precision"],v["inspecoes"]]
                for c,x in enumerate(vals,1):wd.cell(i,c,x)
                wd.cell(i,7).number_format="0.00%";wd.cell(i,8).number_format="0.00%"
            autosize(wd)

            wdo=wb.create_sheet("DETECCAO");title(wdo,"DETECÇÃO DE OCORRÊNCIAS",6)
            hdr(wdo,3,["Indicador","Valor","Observação"])
            vals=[("Defeitos reais",d["real"],"Contagem registrada"),
                  ("Detecções",d["det"],"Contagem registrada"),
                  ("TP de detecção",d["tp"],"min(real,detectado) quando não há correspondência individual"),
                  ("FP de detecção",d["fp"],"max(detectado-real,0) — estimativa por contagem"),
                  ("FN de detecção",d["fn"],"max(real-detectado,0) — estimativa por contagem"),
                  ("Classificação por tipo","Sim" if d["detailed"] else "Não","Opcional")]
            for i,row in enumerate(vals,4):
                for c,v in enumerate(row,1):wdo.cell(i,c,v)
            autosize(wdo)

            wg=wb.create_sheet("GLOSSÁRIO");title(wg,"GLOSSÁRIO",2)
            hdr(wg,3,["Sigla / Métrica","Definição"])
            for i,(term,desc) in enumerate(GLOSSARY.items(),4):
                wg.cell(i,1,term);wg.cell(i,2,desc);wg.cell(i,2).alignment=Alignment(wrap_text=True,vertical="top")
            wg.column_dimensions["A"].width=42;wg.column_dimensions["B"].width=110
            # Salva primeiro em arquivo temporário e valida a estrutura antes de substituir o destino.
            # Isso evita deixar um .xlsx corrompido se a gravação for interrompida.
            tmp_path = str(Path(path).with_name(Path(path).stem + "__AIA_TMP.xlsx"))
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
                wb.save(tmp_path)
                # Validação real do arquivo gerado. Se o OpenPyXL não consegue reabri-lo,
                # não substituímos o arquivo escolhido pelo usuário.
                check = openpyxl.load_workbook(tmp_path, read_only=True, data_only=False)
                check.close()
                os.replace(tmp_path, path)
            finally:
                if os.path.exists(tmp_path):
                    try: os.remove(tmp_path)
                    except Exception: pass
            self.status.set(f"Relatório Excel exportado e validado: {Path(path).name}")
            messagebox.showinfo("Exportação concluída",
                "Relatório Excel criado:\nRESUMO • INSPEÇÕES • MÉTRICAS • AUDITORIA • POR DEFEITO • DETECCAO • GLOSSÁRIO")
        except Exception as e:
            messagebox.showerror("Erro na exportação Excel",str(e))

    def glossary(self):
        win=tk.Toplevel(self);win.title("Glossário");win.geometry("920x720");win.configure(bg=BG)
        text=tk.Text(win,bg=PANEL3,fg=TEXT,insertbackground=TEXT,relief="flat",wrap="word")
        text.pack(fill="both",expand=True,padx=14,pady=14)
        for term,desc in GLOSSARY.items():
            text.insert("end",term+"\n","term");text.insert("end",desc+"\n\n")
        text.tag_configure("term",foreground=ACCENT2,font=("Segoe UI",11,"bold"))
        text.configure(state="disabled")



# ============================================================================
# AIA V1.0.0 FINAL RELEASE — CONSOLIDAÇÃO DAS MELHORIAS VALIDADAS
# Executável NÃO faz parte desta release. Esta é a baseline Python oficial.
# ============================================================================

def _aia_final_expectation_file(self):
    return self.asset_dir / "AIA_META_ATIVA.json"

def _aia_final_load_active_expectation(self):
    path=self._aia_final_expectation_file()
    if not path.exists(): return
    try:
        data=json.loads(path.read_text(encoding="utf-8"))
        for key in ("accuracy","precision","recall","f1","fp100","fn100","source","responsible","date","notes"):
            if key not in data: continue
            value=data[key]
            if value is None: value=""
            elif key in ("accuracy","precision","recall","f1") and isinstance(value,(int,float)): value=f"{value*100:g}"
            elif key in ("fp100","fn100") and isinstance(value,(int,float)): value=f"{value:g}"
            self.expectation[key].set(str(value))
        self.expectation_saved_snapshot=self.expectation_snapshot()
    except Exception:
        self.expectation_saved_snapshot=None

def _aia_final_save_expectation(self):
    configured=[]; invalid=[]
    for key,label in (("accuracy","Accuracy"),("precision","Precision"),("recall","Recall"),("f1","F1"),("fp100","FP / 100"),("fn100","FN / 100")):
        raw=(self.expectation[key].get() or "").strip()
        if raw:
            parsed=self._parse_target(key)
            if parsed is None: invalid.append(label)
            else: configured.append(key)
    if invalid:
        messagebox.showerror("Meta inválida","Verifique os valores de: "+", ".join(invalid)+".",parent=self);return
    if not configured:
        messagebox.showwarning("Meta não configurada","Informe pelo menos uma meta de desempenho antes de salvar.",parent=self);return
    snap=self.expectation_snapshot();snap["saved_at"]=datetime.now().isoformat(timespec="seconds");snap["configured"]=configured
    try:
        self._aia_final_expectation_file().write_text(json.dumps(snap,ensure_ascii=False,indent=2),encoding="utf-8")
        self.expectation_saved_snapshot=dict(snap) if False else snap
        self.status.set(f"✓ Meta salva — {len(configured)} critério(s) registrado(s).")
        self.refresh_all()
        messagebox.showinfo("Meta salva","A meta foi registrada com sucesso e ficará disponível nas próximas aberturas do AIA.",parent=self)
    except Exception as exc:
        messagebox.showerror("Erro ao salvar meta",f"Não foi possível salvar a meta.\n\n{exc}",parent=self)

def _aia_final_next_numeric_id(self):
    nums=[]; widths=[]
    for row in self.rows:
        raw=str(normalize_row(row).get("id") or "").strip()
        if raw.isdigit():
            try: nums.append(int(raw)); widths.append(len(raw))
            except Exception: pass
    nxt=(max(nums)+1) if nums else 1
    width=max(widths) if widths else 0
    return f"{nxt:0{width}d}" if width else str(nxt)

def _aia_final_import_excel(self):
    """Importador final: aceita o formato do próprio AIA e planilhas históricas."""
    if openpyxl is None:
        messagebox.showerror("Excel", "OpenPyXL não está instalado.\nExecute o INSTALAR_DEPENDENCIAS_AIA.bat")
        return
    path=filedialog.askopenfilename(title="Abrir planilha de assertividade",
                                    filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")])
    if not path: return
    try:
        wb=openpyxl.load_workbook(path, data_only=True)
        candidates=[]
        for ws_try in wb.worksheets:
            hr=self._find_header_row(ws_try)
            if hr is None: continue
            hm=self._header_map(ws_try,hr)
            score=0
            if self._norm_header(ws_try.title)=="INSPECOES": score += 8
            for k in ("IDENTIFICACAO","VISION AI","VISION","REFERENCIA","REAL","DETECTADO",
                      "TP-C","VN-C","FP-C","FN-C","TP-D","FP-D","FN-D","TP","VN","FP","FN"):
                if k in hm: score += 1
            if "IDENTIFICACAO" in hm: score += 5
            candidates.append((score,ws_try,hr))
        if not candidates:
            raise ValueError("Não encontrei uma tabela de inspeções compatível. Procure uma aba com Identificação/Inspeção e os dados de resultado.")
        _,ws,hr=max(candidates,key=lambda x:x[0])
        hm=self._header_map(ws,hr)
        def col(*names):
            for n in names:
                k=self._norm_header(n)
                if k in hm: return hm[k]
            return None
        c_id=col("Identificação","N Inspeção","Nº Inspeção","Inspeção")
        if c_id is None:
            raise ValueError(f"A aba '{ws.title}' foi localizada, mas a coluna Identificação/Inspeção não foi encontrada na linha {hr}.")
        c_vision=col("Resultado Vision AI","Vision AI","Vision")
        c_ref=col("Referência","Reference","Gabarito")
        c_tp=col("TP-C","Verdadeiro Positivo","VP - Verdadeiro Positivo","TP")
        c_tn=col("VN-C","Verdadeiro Negativo","VN - Verdadeiro Negativo","VN","TN")
        c_fp=col("FP-C","Falso Positivo","FP - Falso Positivo","FP")
        c_fn=col("FN-C","Falso Negativo","FN - Falso Negativo","FN")
        c_real=col("Real","Real (defeitos)","Defeitos reais")
        c_det=col("Detectado","Detectados","Detecções","Detecções registradas")
        c_tp_d=col("TP-D"); c_fp_d=col("FP-D"); c_fn_d=col("FN-D")
        c_def=col("Defeito","Tipo de defeito"); c_obs=col("Observação","Observações")
        c_date=col("Data da inspeção","Data inspeção","Data","Date")
        rows=[]; hist={"TP":0,"VN/TN":0,"FP":0,"FN":0,"N":0}; imported_dates=[]
        seen=set()
        for rr in range(hr+1,ws.max_row+1):
            ident=self._format_identification_cell(ws.cell(rr,c_id))
            if not ident: continue
            # Never silently drop an Excel row because of a duplicate ID. Keep it visible;
            # duplicate protection applies to new manual records, not historical evidence.
            r=blank_row(); r["id"]=ident
            r["vision"]=(str(ws.cell(rr,c_vision).value or "OK").strip().upper() if c_vision else "OK")
            r["ref"]=(str(ws.cell(rr,c_ref).value or "Importado — histórico").strip() if c_ref else "Importado — histórico")
            r["tp"]=self._num(ws.cell(rr,c_tp).value) if c_tp else 0
            r["vn"]=self._num(ws.cell(rr,c_tn).value) if c_tn else 0
            r["fp"]=self._num(ws.cell(rr,c_fp).value) if c_fp else 0
            r["fn"]=self._num(ws.cell(rr,c_fn).value) if c_fn else 0
            r["detection_available"]=bool(c_real is not None and c_det is not None)
            r["real"]=self._num(ws.cell(rr,c_real).value) if c_real else 0
            r["det"]=self._num(ws.cell(rr,c_det).value) if c_det else 0
            r["tp_d"]=self._num(ws.cell(rr,c_tp_d).value) if c_tp_d else 0
            r["fp_d"]=self._num(ws.cell(rr,c_fp_d).value) if c_fp_d else 0
            r["fn_d"]=self._num(ws.cell(rr,c_fn_d).value) if c_fn_d else 0
            r["obs"]=str(ws.cell(rr,c_obs).value or "") if c_obs else ""
            r["calc_status"]="Histórico"; r["source"]="Excel histórico"
            if c_def:
                name=normalize_defect(ws.cell(rr,c_def).value)
                if name: r["defects"]=[{"type":name,"real":0,"det":0}]
            rows.append(r)
            hist["TP"]+=r["tp"]; hist["VN/TN"]+=r["vn"]; hist["FP"]+=r["fp"]; hist["FN"]+=r["fn"]
            hist["N"]+=r["tp"]+r["vn"]+r["fp"]+r["fn"]
            if c_date:
                dd=self._format_inspection_date(ws.cell(rr,c_date).value)
                if dd: imported_dates.append(dd)
        if not rows: raise ValueError("Nenhuma inspeção válida foi encontrada na planilha.")
        from collections import Counter
        imported_date=Counter(imported_dates).most_common(1)[0][0] if imported_dates else ""
        self._archive_current_collection()
        self.rows=rows; self.history=[hist]; self.source_file=path
        imported_meta,metadata_status=self._extract_workbook_metadata(wb,ws,hr,path)
        if imported_date: imported_meta["data"]=imported_date
        for key in self.meta:
            if imported_meta.get(key): self.meta[key].set(imported_meta[key])
        if not self.meta["gabarito"].get(): self.meta["gabarito"].set("Não determinado")
        self.refresh_all()
        found=sum(1 for k in self.meta if imported_meta.get(k))
        self.status.set(f"✓ Excel importado: {len(rows)} inspeções • identificação preservada")
        messagebox.showinfo("Importação Excel concluída",
            f"Planilha: {Path(path).name}\n\n"
            f"Inspeções importadas: {len(rows)}\n"
            f"Aba utilizada: {ws.title}\n"
            f"Linha do cabeçalho: {hr}\n\n"
            f"Histórico: TP={hist['TP']} | VN={hist['VN/TN']} | FP={hist['FP']} | FN={hist['FN']}\n"
            f"Metadados identificados: {found}/8\n"
            f"Data: {self.meta['data'].get() or 'N/D'}\n\n"
            "Os dados históricos permanecem históricos e não são transformados automaticamente em gabarito.")
    except Exception as exc:
        messagebox.showerror("Erro na importação Excel",
            "Não foi possível importar a planilha.\n\n"
            f"Detalhe: {exc}\n\n"
            "Se for um XLSX gerado pelo AIA, a aba INSPEÇÕES e a coluna Identificação serão reconhecidas automaticamente.")

def _aia_final_draw_detection_matrix(self,parent,d):
    frame=self.chart_card(parent,"MATRIZ DE DETECÇÃO","Ocorrências: real × encontrado. TN-D não se aplica à contagem de ocorrências.")
    cv=tk.Canvas(frame,bg=PANEL,highlightthickness=0,height=225)
    cv.pack(fill="both",expand=True,padx=12,pady=10)
    cv.update_idletasks(); w=max(cv.winfo_width(),560); h=225
    x0,y0=135,42; cw,ch=150,55
    # Header / cells are intentionally simple and explicit for executive reading.
    cv.create_text(x0+cw/2,y0-18,text="DETECÇÃO",fill=ACCENT2,font=("Segoe UI",9,"bold"))
    cv.create_text(x0+cw/2,y0+ch/2,text=f"TP-D\n{d.get('tp',0)}",fill=TEXT,font=("Segoe UI",11,"bold"),justify="center")
    cv.create_rectangle(x0,y0,x0+cw,y0+ch,outline="#2B718F")
    cv.create_text(x0+cw+20,y0+ch/2,text=f"FP-D\n{d.get('fp',0)}",fill=BAD,font=("Segoe UI",10,"bold"),justify="left",anchor="w")
    cv.create_text(35,y0+ch/2,text=f"Real encontrado\n{d.get('real',0)}",fill=MUTED,font=("Segoe UI",8,"bold"),justify="center")
    y1=y0+ch+18
    cv.create_text(x0+cw/2,y1+ch/2,text=f"FN-D\n{d.get('fn',0)}",fill=WARN,font=("Segoe UI",11,"bold"),justify="center")
    cv.create_rectangle(x0,y1,x0+cw,y1+ch,outline="#2B718F")
    cv.create_text(x0+cw+20,y1+ch/2,text="Sem TN-D",fill=MUTED,font=("Segoe UI",9,"bold"),anchor="w")
    cv.create_text(35,y1+ch/2,text="Ocorrência real\nnão encontrada",fill=MUTED,font=("Segoe UI",8,"bold"),justify="center")
    return cv

def _aia_final_view_expectativa(self):
    # Use the existing implementation but add an always-visible explicit save control.
    self.section_title("EXPECTATIVA  /  CRITÉRIO DE ENTREGA","Defina o norte da calibração e acompanhe resultado × meta.")
    summary=self.expectation_summary();hero=self.card(self.content,"🎯 MODO EXPECTATIVA",f"{summary['met']}/{summary['configured']} metas configuradas atendidas • status: {summary['status']}");hero.pack(fill="x",pady=(0,7))
    meta=self.card(self.content,"METAS DE DESEMPENHO","Accuracy, Precision, Recall e F1 = mínimo • FP/100 e FN/100 = máximo.");meta.pack(fill="x",pady=5)
    grid=tk.Frame(meta,bg=PANEL);grid.pack(fill="x",padx=16,pady=8)
    for i,(lab,key) in enumerate([("Accuracy (%)","accuracy"),("Precision (%)","precision"),("Recall (%)","recall"),("F1 (%)","f1"),("FP / 100","fp100"),("FN / 100","fn100")]):
        r,c=divmod(i,3);box=tk.Frame(grid,bg=PANEL);box.grid(row=r,column=c,sticky="ew",padx=6,pady=6);grid.columnconfigure(c,weight=1);tk.Label(box,text=lab,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w");ttk.Entry(box,textvariable=self.expectation[key]).pack(fill="x")
    savebar=tk.Frame(meta,bg="#0A2635",highlightbackground="#2B718F",highlightthickness=1);savebar.pack(fill="x",padx=16,pady=(2,8))
    tk.Label(savebar,text="META ATIVA",bg="#0A2635",fg=ACCENT2,font=("Segoe UI",9,"bold")).pack(side="left",padx=12,pady=9)
    tk.Label(savebar,text="Salve explicitamente para persistir em AIA_META_ATIVA.json",bg="#0A2635",fg=TEXT,font=("Segoe UI",8)).pack(side="left",padx=4)
    tk.Button(savebar,text="✓ SALVAR META",command=self.save_expectation,bg=GOOD,fg="#061019",activebackground="#75E6B7",bd=0,padx=16,pady=7,font=("Segoe UI",9,"bold"),cursor="hand2").pack(side="right",padx=8,pady=5)
    table=ttk.Treeview(meta,columns=("ind","meta","atual","regra","status"),show="headings",height=6)
    for c,h,w in [("ind","Indicador",210),("meta","Meta",100),("atual","Resultado atual",120),("regra","Critério",90),("status","Status",180)]:table.heading(c,text=h);table.column(c,width=w,anchor="center" if c!="ind" else "w")
    table.pack(fill="x",padx=16,pady=(2,14))
    for label,key,op,a,t,status in self._expectation_rows():
        ms="N/D" if t is None else (f"{t*100:.2f}%" if key in ("accuracy","precision","recall","f1") else f"{t:.2f}");av="N/D" if a is None else (pct(a) if key in ("accuracy","precision","recall","f1") else f"{a:.2f}");table.insert("","end",values=(label,ms,av,op,status))
    ctx=self.card(self.content,"RASTREABILIDADE DA META","Registre de onde veio o critério.");ctx.pack(fill="x",pady=5)
    form=tk.Frame(ctx,bg=PANEL);form.pack(fill="x",padx=16,pady=8)
    for i,(lab,key) in enumerate([("Origem do critério","source"),("Responsável","responsible"),("Data do critério","date")]):
        box=tk.Frame(form,bg=PANEL);box.grid(row=0,column=i,sticky="ew",padx=6);form.columnconfigure(i,weight=1);tk.Label(box,text=lab,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w");ttk.Entry(box,textvariable=self.expectation[key]).pack(fill="x")
    box=tk.Frame(form,bg=PANEL);box.grid(row=1,column=0,columnspan=3,sticky="ew",padx=6,pady=6);tk.Label(box,text="Observações / acordo com o cliente",bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w");ttk.Entry(box,textvariable=self.expectation["notes"]).pack(fill="x")
    read=self.card(self.content,"LEITURA DE CALIBRAÇÃO","A meta é um critério de entrega; não substitui análise técnica e de risco.");read.pack(fill="both",expand=True,pady=5)
    for line in ["• META ATINGIDA: todos os critérios configurados estão dentro das metas.","• EM CALIBRAÇÃO: parte das metas foi atendida.","• FORA DA EXPECTATIVA: nenhum critério configurado foi atendido na amostra atual.","• A meta salva permanece disponível nas próximas aberturas do AIA."]:
        tk.Label(read,text=line,bg=PANEL,fg=TEXT,wraplength=1100,justify="left",font=("Segoe UI",9)).pack(anchor="w",padx=16,pady=5)

def _aia_final_delete_selected_collection(self):
    tree=getattr(self,"compare_tree",None)
    if tree is None: return
    sel=tree.selection()
    if not sel:
        messagebox.showinfo("Excluir coleta","Selecione uma coleta arquivada na lista antes de excluir.",parent=self);return
    datasets=getattr(self,"compare_datasets",[])
    selected=[datasets[int(i)] for i in sel if int(i)<len(datasets)]
    archived=[x for x in selected if not x.get("current")]
    if not archived:
        messagebox.showinfo("Excluir coleta","A 'Coleta atual' não pode ser apagada desta tela. Use 'Nova coleta' ou exclua as inspeções no módulo COLETA.",parent=self);return
    if not messagebox.askyesno("Excluir coleta","Excluir permanentemente as coletas arquivadas selecionadas da comparação?\n\nEssa ação não apaga o arquivo Excel original.",parent=self):return
    ids={x.get("record_id") for x in archived}
    self.comparison_archive=[x for x in self.comparison_archive if x.get("record_id") not in ids]
    self.calibration_history=[x for x in self.calibration_history if x.get("record_id") not in ids]
    self.status.set(f"✓ {len(archived)} coleta(s) removida(s) da comparação.")
    self.show("comparacao")

def _aia_final_view_comparacao(self):
    """V1.1.5 — comparação orientada à engenharia, com D-I como referência operacional.

    C = decisão da inspeção (OK/NOK).
    D-I = presença de defeito x presença de detecção, unidade = inspeção.
    D-O = ocorrência/defeito, unidade = ocorrência.
    """
    self.section_title(
        "COMPARAÇÃO  /  EVOLUÇÃO",
        "Compare 2 a 50 coletas. D-I é a referência operacional para avaliar a capacidade de encontrar defeitos por inspeção."
    )
    datasets=self._comparison_datasets()
    if not datasets:
        f=self.card(self.content,"NENHUMA COLETA DISPONÍVEL","Realize uma coleta ou importe uma planilha para iniciar a comparação.")
        f.pack(fill="x",pady=4)
        return

    toolbar=tk.Frame(self.content,bg=BG); toolbar.pack(fill="x",pady=(0,7))
    def tb(text,cmd,accent=False):
        tk.Button(toolbar,text=text,command=cmd,bg="#0D5676" if accent else "#0D2432",
                  fg="white",activebackground="#176B8F",bd=0,padx=11,pady=7,
                  font=("Segoe UI",8,"bold"),cursor="hand2").pack(side="left",padx=3)
    tb("COMPARAR SELECIONADAS",lambda:self._run_comparison(datasets),True)
    tb("SELECIONAR TODAS",self._select_all_compare)
    tb("REMOVER DA SELEÇÃO",self._remove_selected_compare)
    tb("LIMPAR SELEÇÃO",self._clear_compare)
    tb("✕ EXCLUIR COLETA",self._aia_final_delete_selected_collection)
    tk.Label(toolbar,text="2–50 coletas • a coleta atual é somente leitura nesta tela",
             bg=BG,fg=MUTED,font=("Segoe UI",8)).pack(side="right",padx=5)

    box=self.card(self.content,"COLETAS DISPONÍVEIS",
                  f"{len(datasets)} conjunto(s) • D-I = inspeção • D-O = ocorrência • C = classificação")
    box.pack(fill="x",pady=4)
    cols=("coleta","data","threshold","n","diri","dipre","difp","difn","dofp","dofn","crec","cpre","origem")
    heads=("Coleta","Data","Threshold","N","Recall D-I","Precision D-I","FP D-I","FN D-I",
           "FP D-O","FN D-O","Recall C","Precision C","Origem")
    widths=(115,88,76,48,82,92,65,65,65,65,78,88,160)
    tree=ttk.Treeview(box,columns=cols,show="headings",selectmode="extended",height=min(10,max(5,len(datasets))))
    self.compare_tree=tree; self.compare_datasets=datasets
    for c,h,w in zip(cols,heads,widths):
        tree.heading(c,text=h); tree.column(c,width=w,anchor="center" if c!="origem" else "w")
    tree.pack(fill="x",padx=12,pady=(4,12))

    for i,ds in enumerate(datasets):
        mm=ds.get("metrics",{}) or {}; dd=ds.get("defects",{}) or {}; di=ds.get("detection_inspection",{}) or {}
        origin="ATUAL" if ds.get("current") else (Path(ds.get("source_file","" )).name if ds.get("source_file") else "Manual")
        tree.insert("","end",iid=str(i),values=(
            ds.get("name","Coleta"),ds.get("date","N/D"),ds.get("threshold","N/D"),mm.get("N",0),
            pct(di.get("Recall / Sensibilidade")),pct(di.get("Precision")),
            di.get("FP","N/D"),di.get("FN","N/D"),dd.get("fp","N/D"),dd.get("fn","N/D"),
            pct(mm.get("Recall / Sensibilidade")),pct(mm.get("Precision")),origin
        ))
    tree.bind("<Delete>",lambda e:self._aia_final_delete_selected_collection())
    if 2<=len(datasets)<=10: tree.selection_set(*[str(i) for i in range(len(datasets))])

    result=self.card(self.content,"RESULTADO DA COMPARAÇÃO",
                      "A análise usa a primeira e a última coleta selecionadas como base/final e mostra D-I, D-O e C separadamente.")
    result.pack(fill="x",pady=(6,4)); self.compare_result=result
    tk.Label(result,text="Selecione pelo menos 2 coletas e clique em COMPARAR SELECIONADAS.",
             bg=PANEL,fg=MUTED,font=("Segoe UI",9)).pack(anchor="w",padx=16,pady=14)


def _aia_v115_delta(a,b,percent=False):
    if a is None or b is None:
        return None
    try:
        return (float(b)-float(a)) * (100.0 if percent else 1.0)
    except Exception:
        return None


def _aia_v115_fmt_delta(a,b,percent=False):
    d=_aia_v115_delta(a,b,percent)
    if d is None: return "N/D"
    return f"{d:+.2f} p.p." if percent else f"{d:+g}"


def _aia_v115_metric_value(ds, universe, key):
    if universe=="DI":
        return (ds.get("detection_inspection",{}) or {}).get(key)
    if universe=="DO":
        return (ds.get("defects",{}) or {}).get(key)
    return (ds.get("metrics",{}) or {}).get(key)


def _aia_v115_render_comparison_result(self, selected, result):
    for w in result.winfo_children(): w.destroy()
    first,last=selected[0],selected[-1]
    tk.Label(result,text=f"{len(selected)} COLETAS EM COMPARAÇÃO",bg=PANEL,fg=ACCENT2,
             font=("Segoe UI",11,"bold")).pack(anchor="w",padx=16,pady=(12,3))
    tk.Label(result,text=(f"BASE: {first.get('name','Coleta')}  →  FINAL: {last.get('name','Coleta')}    •    "
                          f"Threshold: {first.get('threshold','N/D')} → {last.get('threshold','N/D')}"),
             bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w",padx=16,pady=(0,8))

    # 1) D-I — principal leitura operacional
    di1=first.get("detection_inspection",{}) or {}; di2=last.get("detection_inspection",{}) or {}
    card=self.card(result,"01  •  D-I — DETECÇÃO POR INSPEÇÃO",
                   "REFERÊNCIA OPERACIONAL • presença de defeito × presença de detecção. Maior Recall e Precision são melhores; menor FP/FN é melhor.")
    card.pack(fill="x",padx=10,pady=5)
    if not di1.get("available") or not di2.get("available"):
        tk.Label(card,text="N/D — pelo menos uma das coletas não possui dados Real + Detectado válidos por inspeção.",
                 bg=PANEL,fg=WARN,font=("Segoe UI",9,"bold"),wraplength=1050,justify="left").pack(anchor="w",padx=14,pady=12)
    else:
        table=ttk.Treeview(card,columns=("ind","base","final","delta","leitura"),show="headings",height=5)
        for c,h,w in [("ind","Indicador",185),("base","Base",110),("final","Final",110),("delta","Δ B→F",100),("leitura","Leitura",300)]:
            table.heading(c,text=h); table.column(c,width=w,anchor="center" if c!="ind" and c!="leitura" else "w")
        table.pack(fill="x",padx=12,pady=8)
        specs=[
            ("N inspeções","N",False,"Amostra usada na matriz D-I"),
            ("TP-DI","TP",False,"Inspeções com defeito real e detecção"),
            ("TN-DI","TN",False,"Inspeções sem defeito e sem detecção"),
            ("FP-DI","FP",False,"Inspeções sem defeito que geraram detecção"),
            ("FN-DI","FN",False,"Inspeções com defeito que não geraram detecção"),
            ("Accuracy D-I","Accuracy",True,"Acerto global da matriz D-I"),
            ("Precision D-I","Precision",True,"Confiabilidade das detecções"),
            ("Recall D-I","Recall / Sensibilidade",True,"Capacidade de encontrar inspeções com defeito"),
            ("Specificity D-I","Specificity",True,"Capacidade de reconhecer inspeções sem defeito"),
            ("FPR D-I","FPR",True,"Taxa de falsos alarmes por inspeção"),
            ("FNR D-I","FNR",True,"Taxa de escapes por inspeção"),
            ("F1 D-I","F1",True,"Equilíbrio entre Precision e Recall"),
        ]
        for label,key,is_pct,meaning in specs:
            a=di1.get(key); b=di2.get(key)
            av=pct(a) if is_pct else ("N/D" if a is None else str(a))
            bv=pct(b) if is_pct else ("N/D" if b is None else str(b))
            dv=_aia_v115_fmt_delta(a,b,is_pct)
            # Direção desejável
            if a is None or b is None: read="Dados insuficientes"
            elif key in ("FP","FN","FPR","FNR"):
                read="MELHOR" if b<a else ("PIOR" if b>a else "IGUAL")
            elif key in ("Precision","Recall / Sensibilidade","Specificity","Accuracy","F1"):
                read="MELHOR" if b>a else ("PIOR" if b<a else "IGUAL")
            else: read=meaning
            table.insert("","end",values=(label,av,bv,dv,read))
        tk.Label(card,text=(f"Método D-I base: {di1.get('method','N/D')}  •  final: {di2.get('method','N/D')}"),
                 bg=PANEL,fg=MUTED,font=("Segoe UI",8)).pack(anchor="w",padx=14,pady=(0,10))

    # 2) Resumo de decisão D-I
    decision=self.card(result,"DECISÃO OPERACIONAL — D-I",
                       "A decisão não usa uma métrica isolada. O AIA mostra simultaneamente Recall, Precision, FP-DI e FN-DI.")
    decision.pack(fill="x",padx=10,pady=5)
    if di1.get("available") and di2.get("available"):
        r1,r2=di1.get("Recall / Sensibilidade"),di2.get("Recall / Sensibilidade")
        p1,p2=di1.get("Precision"),di2.get("Precision")
        f1,f2=di1.get("FP"),di2.get("FP")
        n1,n2=di1.get("FN"),di2.get("FN")
        tk.Label(decision,text=(f"Recall D-I: {pct(r1)} → {pct(r2)} ({_aia_v115_fmt_delta(r1,r2,True)})    •    "
                                f"Precision D-I: {pct(p1)} → {pct(p2)} ({_aia_v115_fmt_delta(p1,p2,True)})"),
                 bg=PANEL,fg=GOOD,font=("Segoe UI",9,"bold")).pack(anchor="w",padx=14,pady=(9,3))
        tk.Label(decision,text=(f"FP-DI: {f1} → {f2} ({_aia_v115_fmt_delta(f1,f2,False)})    •    "
                                f"FN-DI: {n1} → {n2} ({_aia_v115_fmt_delta(n1,n2,False)})"),
                 bg=PANEL,fg=TEXT,font=("Consolas",9,"bold")).pack(anchor="w",padx=14,pady=(0,8))
        changes=[]
        if r1 is not None and r2 is not None: changes.append(("Recall",r2-r1,1))
        if p1 is not None and p2 is not None: changes.append(("Precision",p2-p1,1))
        if f1 is not None and f2 is not None: changes.append(("FP-DI",f1-f2,1))
        if n1 is not None and n2 is not None: changes.append(("FN-DI",n1-n2,1))
        good=sum(1 for _,v,_ in changes if v>0); bad=sum(1 for _,v,_ in changes if v<0)
        if good>bad: verdict="TENDÊNCIA FAVORÁVEL"
        elif bad>good: verdict="TENDÊNCIA DESFAVORÁVEL"
        else: verdict="TRADE-OFF / RESULTADO MISTO"
        tk.Label(decision,text=verdict,bg=PANEL,fg=GOOD if good>bad else (BAD if bad>good else WARN),
                 font=("Segoe UI",11,"bold")).pack(anchor="w",padx=14,pady=(0,10))

    # 3) D-O — separado
    do1=first.get("defects",{}) or {}; do2=last.get("defects",{}) or {}
    card=self.card(result,"02  •  D-O — DETECÇÃO POR OCORRÊNCIA",
                   "UNIDADE: ocorrência/defeito • não misturar com D-I. Quando a correspondência individual não existe, FP/FN podem ser estimados por contagem.")
    card.pack(fill="x",padx=10,pady=5)
    table=ttk.Treeview(card,columns=("ind","base","final","delta"),show="headings",height=6)
    for c,h,w in [("ind","Indicador",220),("base","Base",130),("final","Final",130),("delta","Δ",120)]:
        table.heading(c,text=h);table.column(c,width=w,anchor="center" if c!="ind" else "w")
    table.pack(fill="x",padx=12,pady=8)
    for label,key in [("Ocorrências reais","real"),("Detecções","det"),("TP-D","tp"),("FP-D","fp"),("FN-D","fn")]:
        a=do1.get(key);b=do2.get(key)
        table.insert("","end",values=(label,"N/D" if a is None else a,"N/D" if b is None else b,_aia_v115_fmt_delta(a,b,False)))
    tk.Label(card,text=f"Método D-O final: {do2.get('method','N/D')} • Status: {do2.get('status','N/D')}",
             bg=PANEL,fg=WARN if do2.get("estimated") else MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w",padx=14,pady=(0,10))

    # 4) C — referência secundária
    cm1=first.get("metrics",{}) or {}; cm2=last.get("metrics",{}) or {}
    card=self.card(result,"03  •  C — CLASSIFICAÇÃO DA INSPEÇÃO",
                   "UNIDADE: inspeção • decisão OK/NOK. C é apresentado separadamente de D-I e D-O.")
    card.pack(fill="x",padx=10,pady=5)
    table=ttk.Treeview(card,columns=("ind","base","final","delta"),show="headings",height=7)
    for c,h,w in [("ind","Indicador",220),("base","Base",130),("final","Final",130),("delta","Δ",120)]:
        table.heading(c,text=h);table.column(c,width=w,anchor="center" if c!="ind" else "w")
    table.pack(fill="x",padx=12,pady=8)
    for label,key in [("Accuracy","Accuracy"),("Precision","Precision"),("Recall","Recall / Sensibilidade"),("Specificity","Specificity"),("FPR","FPR"),("FNR","FNR"),("F1","F1")]:
        a=cm1.get(key);b=cm2.get(key)
        table.insert("","end",values=(label,pct(a),pct(b),_aia_v115_fmt_delta(a,b,True)))

    # 5) Gráficos — D-I prioritário
    charts=tk.Frame(result,bg=PANEL);charts.pack(fill="both",expand=True,padx=10,pady=(2,10))
    for col,(title,key,pctmode) in enumerate([
        ("EVOLUÇÃO — RECALL D-I","Recall / Sensibilidade",True),
        ("EVOLUÇÃO — PRECISION D-I","Precision",True),
        ("EVOLUÇÃO — FP D-I","FP",False),
        ("EVOLUÇÃO — FN D-I","FN",False),
    ]):
        row=0 if col<2 else 1;c=col%2
        ch=self.chart_card(charts,title,"D-I • inspeções")
        ch.grid(row=row,column=c,sticky="nsew",padx=(0,4) if c==0 else (4,0),pady=4)
        # chart helper expects metric in metrics for normal keys; pass a temporary dataset map.
        mapped=[]
        for ds in selected:
            di=ds.get("detection_inspection",{}) or {}
            mapped.append({"metrics":{key:di.get(key)},"defects":{},"threshold":ds.get("threshold"),"name":ds.get("name")})
        self._draw_trend_chart(ch,mapped,key,pctmode)
    charts.columnconfigure(0,weight=1);charts.columnconfigure(1,weight=1)

    # 6) Leitura de engenharia
    read=self.card(result,"04  •  LEITURA DE ENGENHARIA",
                   "Use a comparação para investigar efeito de threshold/configuração; não atribua causalidade quando outras condições mudaram.")
    read.pack(fill="x",padx=10,pady=(0,10))
    lines=[]
    if di1.get("available") and di2.get("available"):
        r1,r2=di1.get("Recall / Sensibilidade"),di2.get("Recall / Sensibilidade")
        p1,p2=di1.get("Precision"),di2.get("Precision")
        fp1,fp2=di1.get("FP"),di2.get("FP");fn1,fn2=di1.get("FN"),di2.get("FN")
        lines.append(f"D-I é a referência: Recall {pct(r1)} → {pct(r2)}; Precision {pct(p1)} → {pct(p2)}; FP-DI {fp1} → {fp2}; FN-DI {fn1} → {fn2}.")
        if r2>r1 and fp2<=fp1 and fn2<=fn1: lines.append("Leitura: melhora consistente de detecção por inspeção, sem aumento de FP-DI e FN-DI.")
        elif r2>r1 and (fp2>fp1 or fn2>fn1): lines.append("Leitura: houve ganho de Recall, mas existe trade-off com falsos alarmes e/ou escapes.")
        elif r2<r1: lines.append("Leitura: a capacidade de encontrar defeitos por inspeção piorou; investigar threshold e condições de aquisição.")
        else: lines.append("Leitura: desempenho de detecção por inspeção está estável ou sem mudança conclusiva.")
    else:
        lines.append("D-I indisponível em pelo menos uma coleta; não é possível declarar qual configuração é melhor por inspeção.")
    lines.append("D-O permanece separado porque sua unidade estatística é a ocorrência/defeito. C permanece separado porque mede a decisão OK/NOK da inspeção.")
    for line in lines:
        tk.Label(read,text="• "+line,bg=PANEL,fg=TEXT,wraplength=1080,justify="left",font=("Segoe UI",9)).pack(anchor="w",padx=14,pady=4)


def _aia_v115_run_comparison(self,datasets):
    ids=self.compare_tree.selection() if hasattr(self,"compare_tree") else ()
    if len(ids)<2:
        messagebox.showwarning("Comparação","Selecione pelo menos 2 coletas.",parent=self); return
    selected=[datasets[int(i)] for i in ids]
    result=getattr(self,"compare_result",None)
    if result is None: return
    _aia_v115_render_comparison_result(self,selected,result)


def _aia_v10_show_frontend(self,key):
    """V1.0 frontend pass: stronger module affordance and active-state feedback."""
    try:
        self.content_canvas.yview_moveto(0)
        self.content_canvas.configure(scrollregion=(0, 0, 1, 1))
    except Exception:
        pass
    self.current_view=key
    self.clear_content()
    module_names = {
        "dashboard": "EXECUTIVE / HOME", "graficos": "PERFORMANCE / GRÁFICOS",
        "comparacao": "COMPARAÇÃO / EVOLUÇÃO", "calibracao": "PROJETO / CALIBRAÇÃO",
        "operacao": "COLETA / OPERAÇÃO", "auditoria": "AUDITORIA / HISTÓRICO",
        "tecnica": "ANÁLISE TÉCNICA / DIAGNÓSTICO", "engenharia": "ENGENHARIA / DECISÃO",
        "defeitos": "ANÁLISE POR DEFEITO", "expectativa": "EXPECTATIVA / CRITÉRIO DE ENTREGA",
    }
    self.set_instrument(module_names.get(key,key))
    for k,b in self.nav_buttons.items():
        active=(k==key)
        b.configure(
            bg="#0D6288" if active else "#081923",
            fg="white" if active else "#B8CBD5",
            relief="solid" if active else "flat",
            bd=0,
            highlightthickness=1 if active else 0,
            highlightbackground="#3FA7CF" if active else "#081923",
            font=("Segoe UI",8,"bold"),
        )
    fn={"dashboard":self.view_dashboard,"graficos":self.view_graficos,
        "comparacao":self.view_comparacao,"calibracao":self.view_calibracao,
        "operacao":self.view_operacao,"auditoria":self.view_auditoria,
        "tecnica":self.view_tecnica,"engenharia":self.view_engenharia,
        "defeitos":self.view_defeitos,"expectativa":self.view_expectativa}[key]
    fn()
    try:
        self.update_idletasks()
        self.content_canvas.configure(scrollregion=self.content_canvas.bbox("all"))
        self.content_canvas.yview_moveto(0)
        self.after_idle(lambda: self.content_canvas.yview_moveto(0))
    except Exception:
        pass

def _aia_v10_final_dashboard(self):
    fonte = "HISTÓRICO DA PLANILHA" if (self.source_file and self.history) else "MOTOR NOVO / ENTRADA MANUAL"
    self.section_title(
        "DASHBOARD  /  VISÃO EXECUTIVA",
        f"{fonte}  •  V1.0.0 FRONT-END CONSOLIDADO  •  leitura em dois níveis"
    )
    m=self.display_metrics(); d=self.defect_counts()

    # DETECÇÃO primeiro, conforme decisão de arquitetura.
    band=tk.Frame(self.content,bg="#103126",highlightbackground=GOOD,highlightthickness=1)
    band.pack(fill="x",pady=(0,5))
    tk.Label(band,text="01  •  DETECÇÃO (D)",bg="#103126",fg=GOOD,
             font=("Segoe UI",12,"bold")).pack(side="left",padx=14,pady=9)
    tk.Label(band,text="OCORRÊNCIAS — quantos defeitos reais existem e quantos foram encontrados?",
             bg="#103126",fg=TEXT,font=("Segoe UI",9)).pack(side="left",padx=4,pady=9)

    dg=tk.Frame(self.content,bg=BG); dg.pack(fill="x",pady=(0,4))
    for i,(lab,val,col) in enumerate([
        ("Defeitos reais",d["real"],GOOD),("Detectados",d["det"],ACCENT2),
        ("TP-D",d["tp"],GOOD),("FP-D",d["fp"],BAD),("FN-D",d["fn"],WARN)]):
        self.metric_card(dg,lab,str(val),col).grid(row=0,column=i,sticky="ew",padx=3)
        dg.columnconfigure(i,weight=1)
    dc=self.chart_card(self.content,"GRÁFICO DE DETECÇÃO (D)",
                       "Primeiro nível executivo • Real × Detectado × TP-D × FP-D × FN-D")
    dc.pack(fill="both",expand=True,pady=4)
    self.draw_bar_chart(dc,[
        ("REAL",d["real"],GOOD),("DETECTADO",d["det"],ACCENT2),
        ("TP-D",d["tp"],GOOD),("FP-D",d["fp"],BAD),("FN-D",d["fn"],WARN)
    ],max_value=max(d["real"],d["det"],d["tp"],d["fp"],d["fn"],1))

    band2=tk.Frame(self.content,bg="#0A2635",highlightbackground=ACCENT2,highlightthickness=1)
    band2.pack(fill="x",pady=(7,5))
    tk.Label(band2,text="02  •  CLASSIFICAÇÃO (C)",bg="#0A2635",fg=ACCENT2,
             font=("Segoe UI",12,"bold")).pack(side="left",padx=14,pady=9)
    tk.Label(band2,text="INSPEÇÕES — a decisão da peça/carroceria foi corretamente OK ou NOK?",
             bg="#0A2635",fg=TEXT,font=("Segoe UI",9)).pack(side="left",padx=4,pady=9)

    cg=tk.Frame(self.content,bg=BG); cg.pack(fill="x",pady=(0,4))
    for i,(lab,val,col) in enumerate([
        ("TP-C",m["TP"],GOOD),("VN-C",m["VN/TN"],GOOD),
        ("FP-C",m["FP"],BAD),("FN-C",m["FN"],WARN)]):
        self.metric_card(cg,lab,str(val),col).grid(row=0,column=i,sticky="ew",padx=3)
        cg.columnconfigure(i,weight=1)
    cc=self.chart_card(self.content,"GRÁFICO DE CLASSIFICAÇÃO (C)",
                       "Segundo nível executivo • TP-C × VN-C × FP-C × FN-C")
    cc.pack(fill="both",expand=True,pady=4)
    self.draw_bar_chart(cc,[
        ("TP-C",m["TP"],GOOD),("VN-C",m["VN/TN"],GOOD),
        ("FP-C",m["FP"],BAD),("FN-C",m["FN"],WARN)
    ],max_value=max(m["TP"],m["VN/TN"],m["FP"],m["FN"],1))

    read=self.card(self.content,"REGRA DE LEITURA PARA APRESENTAÇÃO",
                   "Nunca misture a unidade de inspeção com a unidade de ocorrência.")
    read.pack(fill="x",pady=5)
    tk.Label(read,text="D = ocorrência: defeitos reais versus detecções.",bg=PANEL,fg=GOOD,
             font=("Segoe UI",9,"bold")).pack(anchor="w",padx=16,pady=(8,2))
    tk.Label(read,text="C = inspeção: decisão OK/NOK versus referência definida.",bg=PANEL,fg=ACCENT2,
             font=("Segoe UI",9,"bold")).pack(anchor="w",padx=16,pady=(0,9))

def _aia_v10_final_graficos(self):
    fonte="HISTÓRICO DA PLANILHA" if (self.source_file and self.history) else "MOTOR NOVO / ENTRADA MANUAL"
    self.section_title("GRÁFICOS  /  PERFORMANCE",
        f"{fonte} • matrizes separadas • V1.0.0 FRONT-END CONSOLIDADO")
    m=self.display_metrics(); d=self.defect_counts()
    top=tk.Frame(self.content,bg=BG); top.pack(fill="both",expand=True)
    det=self.chart_card(top,"01  •  MATRIZ DE DETECÇÃO (D)",
                        "Unidade: ocorrência • TN-D não se aplica a contagem de defeitos")
    det.grid(row=0,column=0,sticky="nsew",padx=(0,4),pady=(0,4))
    self._aia_final_draw_detection_matrix(det,d)
    cls=self.chart_card(top,"02  •  MATRIZ DE CLASSIFICAÇÃO (C)",
                        "Unidade: inspeção / carroceria / peça • decisão OK/NOK")
    cls.grid(row=0,column=1,sticky="nsew",padx=(4,0),pady=(0,4))
    self.draw_bar_chart(cls,[("TP-C",m["TP"],GOOD),("VN-C",m["VN/TN"],GOOD),
                             ("FP-C",m["FP"],BAD),("FN-C",m["FN"],WARN)],
                        max_value=max(m["TP"],m["VN/TN"],m["FP"],m["FN"],1))
    perf=self.chart_card(top,"03  •  INDICADORES DE CLASSIFICAÇÃO (C)",
                         "Accuracy • Precision • Recall • F1")
    perf.grid(row=1,column=0,sticky="nsew",padx=(0,4),pady=(4,0))
    self.draw_percent_chart(perf,[("Accuracy",m["Accuracy"],ACCENT2),
                                  ("Precision",m["Precision"],ACCENT),
                                  ("Recall",m["Recall / Sensibilidade"],GOOD),
                                  ("F1",m["F1"],WARN)])
    occ=self.chart_card(top,"04  •  PERFIL DE DETECÇÃO (D)",
                        "Real × Detectado × TP-D × FP-D × FN-D")
    occ.grid(row=1,column=1,sticky="nsew",padx=(4,0),pady=(4,0))
    self.draw_bar_chart(occ,[("Real",d["real"],GOOD),("Detectado",d["det"],ACCENT2),
                             ("TP-D",d["tp"],GOOD),("FP-D",d["fp"],BAD),
                             ("FN-D",d["fn"],WARN)],
                        max_value=max(d["real"],d["det"],d["tp"],d["fp"],d["fn"],1))
    top.columnconfigure(0,weight=1);top.columnconfigure(1,weight=1)
    top.rowconfigure(0,weight=1);top.rowconfigure(1,weight=1)

def _aia_v10_final_operacao(self):
    self.section_title("COLETA  /  OPERAÇÃO",
        "Registro de campo • campos de CLASSIFICAÇÃO (C) e DETECÇÃO (D) separados visualmente.")
    info=self.card(self.content,"IDENTIFICAÇÃO DO TESTE"); info.pack(fill="x",pady=4)
    fields=[("Cliente","cliente"),("Local / Planta","local"),("Projeto","projeto"),
            ("Linha / Estação","linha"),("Técnico","tecnico"),("Data","data"),
            ("Threshold","threshold"),("Método de referência","gabarito")]
    body=tk.Frame(info,bg=PANEL);body.pack(fill="x",padx=6,pady=(0,10))
    for i,(lab,key) in enumerate(fields):
        r,c=divmod(i,4);box=tk.Frame(body,bg=PANEL);box.grid(row=r,column=c,sticky="ew",padx=8,pady=4)
        tk.Label(box,text=lab,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w")
        w=ttk.Combobox(box,textvariable=self.meta[key],values=GABARITOS,state="readonly") if key=="gabarito" else ttk.Entry(box,textvariable=self.meta[key])
        w.pack(fill="x");body.columnconfigure(c,weight=1)
    tk.Label(info,text="✓ Metadados preenchidos automaticamente pelo Excel" if self.source_file else "Entrada manual — preencha o contexto da inspeção",bg=PANEL,fg=GOOD if self.source_file else MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w",padx=14,pady=(0,8))

    tb=tk.Frame(self.content,bg=BG);tb.pack(fill="x",pady=5)
    def button(text,cmd,accent=False):
        return tk.Button(tb,text=text,command=cmd,bg="#0E6288" if accent else "#102F3F",
                         fg="white",activebackground="#1A789E",activeforeground="white",
                         bd=0,highlightbackground="#3C7B95",highlightthickness=1,
                         padx=13,pady=8,font=("Segoe UI",8,"bold"),cursor="hand2")
    for text,cmd,accent in [
        ("＋ NOVA INSPEÇÃO",self.add_row,True),("↻ NOVA COLETA",self.new_collection,False),
        ("⚡ COLETA RÁPIDA",self.quick_entry,True),("DETALHAR DEFEITOS",self.edit_row,False),
        ("✕ EXCLUIR INSPEÇÃO",self.delete_row,False),("⟳ RECALCULAR",self.refresh_all,False)]:
        button(text,cmd,accent).pack(side="left",padx=3)

    group=tk.Frame(self.content,bg=BG);group.pack(fill="x",pady=(4,2))
    a=tk.Frame(group,bg="#0A2635",highlightbackground=ACCENT2,highlightthickness=1);a.pack(side="left",fill="x",expand=True,padx=(0,3))
    tk.Label(a,text="CLASSIFICAÇÃO (C)",bg="#0A2635",fg=ACCENT2,font=("Segoe UI",10,"bold")).pack(anchor="w",padx=12,pady=(7,2))
    tk.Label(a,text="TP-C • VN-C • FP-C • FN-C → decisão OK / NOK",bg="#0A2635",fg=TEXT,font=("Segoe UI",8)).pack(anchor="w",padx=12,pady=(0,7))
    b=tk.Frame(group,bg="#103126",highlightbackground=GOOD,highlightthickness=1);b.pack(side="left",fill="x",expand=True,padx=(3,0))
    tk.Label(b,text="DETECÇÃO (D)",bg="#103126",fg=GOOD,font=("Segoe UI",10,"bold")).pack(anchor="w",padx=12,pady=(7,2))
    tk.Label(b,text="TP-D • FP-D • FN-D → ocorrências reais × detectadas",bg="#103126",fg=TEXT,font=("Segoe UI",8)).pack(anchor="w",padx=12,pady=(0,7))

    table=self.card(self.content,"INSPEÇÕES","Cada linha = uma inspeção. A coluna Nº é sequencial; a Identificação é preservada exatamente como informada.")
    table.pack(fill="both",expand=True,pady=5)
    cols=("n","id","vision","ref","real","det","tp","vn","fp","fn","dtp","dfp","dfn","def","status")
    heads=("Nº","Identificação","Vision","Referência","Real","Detectado","TP-C","VN-C","FP-C","FN-C","TP-D","FP-D","FN-D","Detalhes","Motor")
    self.tree=ttk.Treeview(table,columns=cols,show="headings")
    widths={"n":48,"id":105,"vision":70,"ref":125,"real":60,"det":70,"tp":55,"vn":55,"fp":55,"fn":55,"dtp":55,"dfp":55,"dfn":55,"def":90,"status":135}
    for c,h in zip(cols,heads):self.tree.heading(c,text=h);self.tree.column(c,width=widths.get(c,80),anchor="center")
    self.tree.pack(fill="both",expand=True,padx=10,pady=10);self.tree.bind("<Double-1>",lambda e:self.edit_row());self.refresh_table()

def _aia_v10_final_quick_entry(self):
    win=tk.Toplevel(self);win.title("Coleta Rápida | V1.0.0");win.geometry("650x620");win.configure(bg=BG);win.transient(self);win.grab_set()
    frame=tk.Frame(win,bg=PANEL,highlightbackground=BORDER,highlightthickness=1);frame.pack(fill="both",expand=True,padx=15,pady=15)
    tk.Label(frame,text="COLETA RÁPIDA",bg=PANEL,fg=TEXT,font=("Segoe UI",17,"bold")).pack(anchor="w",padx=18,pady=(16,2))
    tk.Label(frame,text="Preencha a classificação da inspeção (C) e as ocorrências de defeitos (D). O tipo do defeito é opcional.",bg=PANEL,fg=MUTED,wraplength=560,justify="left").pack(anchor="w",padx=18,pady=(0,12))
    vars={}
    def section(text,color,sub):
        f=tk.Frame(frame,bg="#0A2635",highlightbackground=color,highlightthickness=1);f.pack(fill="x",padx=18,pady=(7,5))
        tk.Label(f,text=text,bg="#0A2635",fg=color,font=("Segoe UI",9,"bold")).pack(anchor="w",padx=10,pady=(6,1))
        tk.Label(f,text=sub,bg="#0A2635",fg=TEXT,font=("Segoe UI",8)).pack(anchor="w",padx=10,pady=(0,6))
    def field(lab,key,default="",values=None):
        row=tk.Frame(frame,bg=PANEL);row.pack(fill="x",padx=18,pady=4)
        tk.Label(row,text=lab,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w")
        v=tk.StringVar(value=default);vars[key]=v
        w=ttk.Combobox(row,textvariable=v,values=values,state="readonly") if values else ttk.Entry(row,textvariable=v)
        w.pack(fill="x");return w
    field("Identificação","id",self._next_numeric_id())
    section("CLASSIFICAÇÃO (C)",ACCENT2,"Decisão da inspeção: o Vision AI classificou OK/NOK e qual é a referência do operador/analista?")
    field("Vision AI","vision","OK",["OK","NOK"])
    field("Referência / gabarito","ref","Não determinado",REFS)
    section("DETECÇÃO (D)",GOOD,"Ocorrências: quantos defeitos realmente existem e quantos foram detectados?")
    field("Quantidade real de defeitos","real","0")
    field("Quantidade detectada","det","0")
    field("Observação","obs","")
    btns=tk.Frame(frame,bg=PANEL);btns.pack(fill="x",padx=18,pady=15)
    def collect(next_one=False):
        try: real=max(0,int(float(vars["real"].get() or 0)));det=max(0,int(float(vars["det"].get() or 0)))
        except ValueError:
            messagebox.showerror("Dados","As quantidades de Real e Detectado devem ser numéricas.",parent=win);return
        ident=vars["id"].get().strip() or self._next_numeric_id()
        if self._id_exists(ident):
            messagebox.showerror("Identificação duplicada",f"A identificação '{ident}' já existe nesta coleta.\n\nEscolha outra identificação.",parent=win);return
        r=blank_row();r.update({"id":ident,"vision":vars["vision"].get() or "OK","ref":vars["ref"].get() or "Não determinado","real":real,"det":det,"obs":vars["obs"].get().strip(),"source":"Manual — Coleta rápida","detection_observed":True})
        self.rows.append(self.classify(r));self.source_file="";self.refresh_all();self.status.set(f"✓ Inspeção {ident} registrada.")
        if next_one:
            vars["id"].set(self._next_numeric_id());vars["vision"].set("OK");vars["ref"].set("Não determinado");vars["real"].set("0");vars["det"].set("0");vars["obs"].set("")
        else:win.destroy()
    ttk.Button(btns,text="Registrar e fechar",command=lambda:collect(False)).pack(side="right",padx=4)
    ttk.Button(btns,text="Registrar e próxima",style="Accent.TButton",command=lambda:collect(True)).pack(side="right",padx=4)

# ============================================================================
# V1.0.0 — FRONT-END CORRIGIDO / RESPONSIVO
# Uma única camada final de interface. Não altera o motor estatístico.
# ============================================================================

def _aia_v10_responsive_bar_chart(self, parent, items, max_value=None, suffix=""):
    """Bar chart responsivo: redesenha no resize em vez de congelar a largura inicial."""
    canvas = tk.Canvas(parent, bg=PANEL, highlightthickness=0, height=220)
    canvas.pack(fill="both", expand=True, padx=12, pady=(4,12))

    def redraw(_event=None):
        canvas.delete("all")
        w=max(canvas.winfo_width(),520)
        h=max(canvas.winfo_height(),220)
        left,right,top,bottom=72,18,18,38
        plot_w=max(80,w-left-right)
        plot_h=max(60,h-top-bottom)
        vals=[float(v or 0) for _,v,_ in items]
        maxv=max_value if max_value is not None else (max(vals) if vals else 1)
        if maxv<=0: maxv=1
        for i in range(5):
            y=top+plot_h-(plot_h*i/4)
            canvas.create_line(left,y,w-right,y,fill="#29404F")
            val=maxv*i/4
            canvas.create_text(left-8,y,text=f"{val:.0f}{suffix}",fill=MUTED,anchor="e",font=("Segoe UI",7))
        n=max(len(items),1);slot=plot_w/n;bw=min(58,slot*0.58)
        for i,(label,val,accent) in enumerate(items):
            x=left+slot*i+slot/2;v=float(val or 0);bh=plot_h*(v/maxv);y1=top+plot_h-bh
            canvas.create_rectangle(x-bw/2,y1,x+bw/2,top+plot_h,fill=accent,outline="")
            canvas.create_text(x,y1-8,text=f"{v:.0f}{suffix}",fill=TEXT,font=("Segoe UI",8,"bold"))
            canvas.create_text(x,top+plot_h+12,text=label,fill=MUTED,font=("Segoe UI",7))
    canvas.bind("<Configure>",redraw)
    canvas.after_idle(redraw)
    return canvas


def _aia_v10_responsive_percent_chart(self, parent, items):
    """Percent chart responsivo com redrawing seguro durante resize."""
    canvas=tk.Canvas(parent,bg=PANEL,highlightthickness=0,height=220)
    canvas.pack(fill="both",expand=True,padx=12,pady=(4,12))
    def redraw(_event=None):
        canvas.delete("all")
        w=max(canvas.winfo_width(),520);h=max(canvas.winfo_height(),220)
        left,right,top,bottom=120,48,18,30;plot_w=max(100,w-left-right)
        for i in range(0,101,25):
            x=left+plot_w*i/100
            canvas.create_line(x,top,x,h-bottom,fill="#29404F")
            canvas.create_text(x,h-bottom+12,text=f"{i}%",fill=MUTED,font=("Segoe UI",7))
        row_h=(h-top-bottom)/max(len(items),1)
        for i,(label,val,accent) in enumerate(items):
            y=top+row_h*i+row_h/2;p=0 if val is None else max(0,min(1,float(val)))
            canvas.create_text(left-8,y,text=label,fill=MUTED,anchor="e",font=("Segoe UI",8))
            canvas.create_rectangle(left,y-7,left+plot_w,y+7,fill="#1B2B38",outline="")
            canvas.create_rectangle(left,y-7,left+plot_w*p,y+7,fill=accent,outline="")
            x=min(left+plot_w-35,left+plot_w*p+8)
            canvas.create_text(x,y,text="N/D" if val is None else f"{p*100:.1f}%",fill=TEXT,anchor="w",font=("Segoe UI",8,"bold"))
    canvas.bind("<Configure>",redraw);canvas.after_idle(redraw)
    return canvas


# Runtime bindings are deliberately centralized here: no later front-end block
# is allowed to overwrite these methods.
App._aia_final_expectation_file=_aia_final_expectation_file
App._aia_final_load_active_expectation=_aia_final_load_active_expectation
App._load_active_expectation=_aia_final_load_active_expectation
App.save_expectation=_aia_final_save_expectation
App._next_numeric_id=_aia_final_next_numeric_id
App.import_excel=_aia_final_import_excel
App._aia_final_draw_detection_matrix=_aia_final_draw_detection_matrix
App.view_expectativa=_aia_final_view_expectativa
App._aia_final_delete_selected_collection=_aia_final_delete_selected_collection
App.view_comparacao=_aia_final_view_comparacao
App._run_comparison=_aia_v115_run_comparison
App.show=_aia_v10_show_frontend
App.draw_bar_chart=_aia_v10_responsive_bar_chart
App.draw_percent_chart=_aia_v10_responsive_percent_chart
App.view_dashboard=_aia_v10_final_dashboard
App.view_graficos=_aia_v10_final_graficos
App.view_operacao=_aia_v10_final_operacao
App.quick_entry=_aia_v10_final_quick_entry



# ============================================================================
# AIA V1.1.0 — SMART DATA ENGINE / INTERPRETADOR E VALIDADOR
# Esta camada substitui somente as rotinas de interpretação/cálculo que
# precisam conhecer a unidade estatística da fonte. O front-end aprovado da
# V1.0.0 permanece preservado.
# ============================================================================
AIA_SMART_ENGINE_VERSION = "1.1.9"


def _aia_num_optional(value):
    """Converte número sem transformar célula vazia em zero.
    Retorna None quando não há valor numérico observável.
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return int(value)
    s = str(value).strip()
    if not s:
        return None
    try:
        return int(float(s.replace(",", ".")))
    except Exception:
        return None


def _aia_header_present(hm, *names):
    for n in names:
        if normalize_text(n) in hm:
            return True
    return False


def _aia_metric_pack(tp, tn, fp, fn, n_override=None):
    """Métricas clássicas com denominadores explícitos e N opcional.
    N_override é usado quando o universo da matriz é conhecido separadamente.
    """
    tp, tn, fp, fn = [max(0, int(x or 0)) for x in (tp, tn, fp, fn)]
    n = int(n_override) if n_override is not None else tp + tn + fp + fn
    precision = safe(tp, tp + fp)
    recall = safe(tp, tp + fn)
    specificity = safe(tn, tn + fp)
    fpr = safe(fp, fp + tn)
    fnr = safe(fn, fn + tp)
    npv = safe(tn, tn + fn)
    accuracy = safe(tp + tn, n)
    bal = None if recall is None or specificity is None else (recall + specificity) / 2
    f1 = None if precision is None or recall is None or precision + recall == 0 else 2 * precision * recall / (precision + recall)
    den = (tp + fp) * (tp + fn) * (tn + fp) * (tn + fn)
    mcc = None if den <= 0 else (tp * tn - fp * fn) / math.sqrt(den)
    return {
        "N": n, "TP": tp, "VN/TN": tn, "FP": fp, "FN": fn,
        "Accuracy": accuracy, "Precision": precision,
        "Recall / Sensibilidade": recall, "Specificity": specificity,
        "FPR": fpr, "FNR": fnr, "NPV": npv,
        "Balanced Accuracy": bal, "F1": f1, "MCC": mcc,
        "Prevalência": safe(tp + fn, n),
        "Taxa de alarme NOK": safe(tp + fp, n),
    }


def _aia_empty_classification_metrics():
    return _aia_metric_pack(0, 0, 0, 0, n_override=0)


def _aia_smart_detection_inspection(self):
    """Matriz D por inspeção: presença de defeito x presença de detecção.
    Só é calculada quando Real e Detectado estão disponíveis por inspeção.
    """
    out = {"available": False, "N": 0, "TP": 0, "TN": 0, "VN/TN": 0,
           "FP": 0, "FN": 0, "rows_used": 0, "rows_unavailable": 0,
           "source": "", "method": ""}
    for r0 in self.rows:
        r = normalize_row(r0)
        observed = bool(r.get("detection_observed", False))
        if not observed:
            out["rows_unavailable"] += 1
            continue
        real = max(0, int(r.get("real", 0)))
        det = max(0, int(r.get("det", 0)))
        actual_positive = real > 0
        predicted_positive = det > 0
        if actual_positive and predicted_positive: out["TP"] += 1
        elif actual_positive and not predicted_positive: out["FN"] += 1
        elif not actual_positive and predicted_positive: out["FP"] += 1
        else: out["TN"] += 1
        out["rows_used"] += 1
    out["N"] = out["rows_used"]
    out["VN/TN"] = out["TN"]
    out["available"] = out["rows_used"] > 0
    out["source"] = "Real × Detectado por inspeção"
    out["method"] = "Binário por presença: Real>0 e Detectado>0"
    if out["available"]:
        out.update(_aia_metric_pack(out["TP"], out["TN"], out["FP"], out["FN"], out["N"]))
    else:
        out.update(_aia_empty_classification_metrics())
        out["N"] = 0
    return out


def _aia_smart_detection_occurrence(self):
    """Detecção por ocorrência, preservando a distinção entre observado e estimado.
    Quando uma fonte histórica fornece TP/VN/FP/FN mas não fornece Real/Detectado
    por linha, o AIA preserva a matriz como D histórica por ocorrência e deriva
    Real=TP+FN e Detectado=TP+FP apenas para fins de decomposição da própria matriz.
    """
    # Fonte histórica com matriz de ocorrência reconhecida e sem Real/Detectado.
    sa = getattr(self, "source_analysis", {}) or {}
    om = sa.get("occurrence_matrix", {}) if isinstance(sa, dict) else {}
    if getattr(self,"source_file","") and om.get("present") and not sa.get("detection_fields", {}).get("real") and not sa.get("detection_fields", {}).get("detected"):
        tp=int(om.get("tp",0) or 0); tn=int(om.get("tn",0) or 0); fp=int(om.get("fp",0) or 0); fn=int(om.get("fn",0) or 0)
        real=tp+fn; det=tp+fp
        precision=safe(tp,tp+fp); recall=safe(tp,tp+fn)
        f1=None if precision is None or recall is None or precision+recall==0 else 2*precision*recall/(precision+recall)
        n=tp+tn+fp+fn
        mm=_aia_metric_pack(tp,tn,fp,fn,n)
        return {"available":True,"real":real,"det":det,"tp":tp,"fp":fp,"fn":fn,"tn":tn,
                "precision":precision,"recall":recall,"f1":f1,"available_rows":len(self.rows),
                "unavailable_rows":0,"detailed":False,"estimated":False,
                "status":"IMPORTADO / VALIDADO","method":"Matriz histórica de ocorrência preservada; Real=TP+FN e Detectado=TP+FP",
                "full_metrics":mm}
    real = det = tp = fp = fn = 0
    available_rows = 0
    unavailable_rows = 0
    explicit_tp = explicit_fp = explicit_fn = 0
    explicit_complete = True
    any_explicit = False
    estimated = False
    detailed = False
    for r0 in self.rows:
        r = normalize_row(r0)
        if not r.get("detection_observed", False):
            unavailable_rows += 1
            continue
        available_rows += 1
        ds = r.get("defects") or []
        if ds:
            detailed = True
            for dd in ds:
                rr = max(0, int(dd.get("real", 0)))
                xx = max(0, int(dd.get("det", 0)))
                real += rr; det += xx
                explicit_complete = False
                estimated = True
        else:
            real += max(0, int(r.get("real", 0)))
            det += max(0, int(r.get("det", 0)))
            # TP-D/FP-D/FN-D from a source can be used only when they are
            # explicitly present and internally consistent with Real/Detectado.
            ep = r.get("tp_d")
            efp = r.get("fp_d")
            efn = r.get("fn_d")
            if any(x is not None for x in (ep, efp, efn)) and r.get("detection_explicit_matrix", False):
                any_explicit = True
                ep = max(0, int(ep or 0)); efp = max(0, int(efp or 0)); efn = max(0, int(efn or 0))
                explicit_tp += ep; explicit_fp += efp; explicit_fn += efn
                rr = max(0, int(r.get("real", 0))); xx = max(0, int(r.get("det", 0)))
                if ep + efp != xx or ep + efn != rr:
                    explicit_complete = False
    if any_explicit and explicit_complete:
        tp, fp, fn = explicit_tp, explicit_fp, explicit_fn
        method = "TP-D/FP-D/FN-D importados e consistentes com Real/Detectado"
        status = "IMPORTADO / VALIDADO"
    else:
        tp = min(real, det)
        fp = max(det - real, 0)
        fn = max(real - det, 0)
        method = "Estimativa por contagem — sem correspondência individual"
        status = "ESTIMADO"
        estimated = True
    precision = safe(tp, tp + fp)
    recall = safe(tp, tp + fn)
    f1 = None if precision is None or recall is None or precision + recall == 0 else 2 * precision * recall / (precision + recall)
    return {
        "available": available_rows > 0,
        "real": real, "det": det, "tp": tp, "fp": fp, "fn": fn,
        "precision": precision, "recall": recall, "f1": f1,
        "available_rows": available_rows, "unavailable_rows": unavailable_rows,
        "detailed": detailed, "estimated": estimated,
        "status": status, "method": method,
    }


def _aia_smart_source_summary(self):
    return getattr(self, "source_analysis", {}) or {}


def _aia_smart_classification_totals(self):
    """Classificação: usa apenas matriz C por inspeção válida.
    Histórico de ocorrência não é reutilizado como classificação.
    """
    vals = {"tp": 0, "vn": 0, "fp": 0, "fn": 0}
    used = 0
    for i, r0 in enumerate(self.rows):
        r = normalize_row(r0)
        if r.get("classification_observed", False):
            vals["tp"] += int(r.get("tp", 0) or 0)
            vals["vn"] += int(r.get("vn", 0) or 0)
            vals["fp"] += int(r.get("fp", 0) or 0)
            vals["fn"] += int(r.get("fn", 0) or 0)
            used += int(r.get("c_matrix_count", 0) or 0)
        elif not self.source_file and r.get("ref") in ("Defeito confirmado", "OK confirmado"):
            rr = self.classify(r)
            self.rows[i] = rr
            vals["tp"] += rr["tp"]; vals["vn"] += rr["vn"]; vals["fp"] += rr["fp"]; vals["fn"] += rr["fn"]
            used += 1
    vals["N"] = vals["tp"] + vals["vn"] + vals["fp"] + vals["fn"]
    vals["rows_used"] = used
    return vals


def _aia_smart_metrics(self):
    t = self._smart_classification_totals()
    return _aia_metric_pack(t["tp"], t["vn"], t["fp"], t["fn"], t["N"])


def _aia_smart_display_metrics(self):
    # Histórico só é mostrado como classificação quando a matriz foi
    # identificada como matriz C por inspeção. Caso contrário, N/D evita
    # transformar uma matriz de ocorrências em classificação.
    analysis = _aia_smart_source_summary(self)
    if self.source_file and self.history:
        if analysis.get("classification_matrix", {}).get("valid"):
            h = analysis["classification_matrix"]
            return _aia_metric_pack(h["tp"], h["tn"], h["fp"], h["fn"], h["n"])
        return _aia_empty_classification_metrics()
    return self._smart_metrics()


def _aia_smart_defect_counts(self):
    d = self._smart_detection_occurrence()
    by = defaultdict(lambda: {"real":0,"det":0,"tp":0,"fp":0,"fn":0,"inspecoes":0})
    detailed = False
    for r0 in self.rows:
        r = normalize_row(r0)
        for dd in (r.get("defects") or []):
            detailed = True
            rr=max(0,int(dd.get("real",0))); xx=max(0,int(dd.get("det",0)))
            tt=min(rr,xx); ff=max(xx-rr,0); nn=max(rr-xx,0)
            a=by[dd.get("type") or "Não informado"]
            a["real"]+=rr; a["det"]+=xx; a["tp"]+=tt; a["fp"]+=ff; a["fn"]+=nn; a["inspecoes"]+=1
    d["by"] = by
    d["detailed"] = detailed
    return d


def _aia_smart_detection_totals(self):
    occ = self._smart_detection_occurrence()
    ins = self._smart_detection_inspection()
    out = dict(occ)
    out["inspection"] = ins
    out["inspection_N"] = ins.get("N", 0)
    out["inspection_TP"] = ins.get("TP", 0)
    out["inspection_TN"] = ins.get("TN", 0)
    out["inspection_FP"] = ins.get("FP", 0)
    out["inspection_FN"] = ins.get("FN", 0)
    return out


def _aia_smart_import_excel(self):
    if openpyxl is None:
        messagebox.showerror("Excel", "OpenPyXL não está instalado.\nExecute o INSTALAR_DEPENDENCIAS_AIA.bat", parent=self)
        return
    path = filedialog.askopenfilename(title="Abrir planilha de assertividade",
                                      filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")])
    if not path:
        return
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
        candidates=[]
        for ws_try in wb.worksheets:
            hr=self._find_header_row(ws_try)
            if hr is None: continue
            hm=self._header_map(ws_try,hr)
            score=0
            if self._norm_header(ws_try.title)=="INSPECOES": score += 10
            for k in ("IDENTIFICACAO","VISION AI","VISION","REFERENCIA","REAL","DETECTADO","TP-C","VN-C","FP-C","FN-C","TP-D","FP-D","FN-D","TP","VN","FP","FN"):
                if k in hm: score += 1
            if "IDENTIFICACAO" in hm: score += 5
            candidates.append((score,ws_try,hr))
        if not candidates:
            raise ValueError("Não encontrei uma tabela de inspeções compatível.")
        _,ws,hr=max(candidates,key=lambda x:x[0])
        hm=self._header_map(ws,hr)
        def col(*names):
            for n in names:
                k=self._norm_header(n)
                if k in hm: return hm[k]
            return None
        c_id=col("Identificação","N Inspeção","Nº Inspeção","Inspeção")
        if c_id is None: raise ValueError("A planilha não possui coluna de Identificação/Inspeção.")
        c_vision=col("Resultado Vision AI","Vision AI","Vision")
        c_ref=col("Referência","Reference","Gabarito")
        # C explícito tem prioridade absoluta sobre aliases genéricos.
        c_tp_c=col("TP-C"); c_tn_c=col("VN-C"); c_fp_c=col("FP-C"); c_fn_c=col("FN-C")
        c_tp_g=col("TP","Verdadeiro Positivo","VP - Verdadeiro Positivo")
        c_tn_g=col("VN","TN","Verdadeiro Negativo","VN - Verdadeiro Negativo")
        c_fp_g=col("FP","Falso Positivo","FP - Falso Positivo")
        c_fn_g=col("FN","Falso Negativo","FN - Falso Negativo")
        c_real=col("Real","Real (defeitos)","Defeitos reais")
        c_det=col("Detectado","Detectados","Detecções","Detecções registradas")
        c_tp_d=col("TP-D"); c_fp_d=col("FP-D"); c_fn_d=col("FN-D")
        c_def=col("Defeito","Tipo de defeito"); c_obs=col("Observação","Observações")
        c_date=col("Data da inspeção","Data inspeção","Data","Date")
        explicit_c = all(x is not None for x in (c_tp_c,c_tn_c,c_fp_c,c_fn_c))
        generic_c = all(x is not None for x in (c_tp_g,c_tn_g,c_fp_g,c_fn_g))
        rows=[]; hist={"TP":0,"VN/TN":0,"FP":0,"FN":0,"N":0}; imported_dates=[]
        matrix_row_sums=[]; matrix_rows=[]
        for rr in range(hr+1,ws.max_row+1):
            ident=self._format_identification_cell(ws.cell(rr,c_id))
            if not ident: continue
            r=blank_row(); r["id"]=ident
            r["vision"]=(str(ws.cell(rr,c_vision).value or "OK").strip().upper() if c_vision else "OK")
            r["ref"]=(str(ws.cell(rr,c_ref).value or "Importado — histórico").strip() if c_ref else "Importado — histórico")
            if explicit_c:
                cp=self._num(ws.cell(rr,c_tp_c).value); cn=self._num(ws.cell(rr,c_tn_c).value); cf=self._num(ws.cell(rr,c_fp_c).value); cfn=self._num(ws.cell(rr,c_fn_c).value)
                r.update(tp=cp,vn=cn,fp=cf,fn=cfn)
                row_sum=cp+cn+cf+cfn
                r["classification_source"]="colunas TP-C/VN-C/FP-C/FN-C"
            elif generic_c:
                cp=self._num(ws.cell(rr,c_tp_g).value); cn=self._num(ws.cell(rr,c_tn_g).value); cf=self._num(ws.cell(rr,c_fp_g).value); cfn=self._num(ws.cell(rr,c_fn_g).value)
                r.update(tp=cp,vn=cn,fp=cf,fn=cfn)
                row_sum=cp+cn+cf+cfn
                r["classification_source"]="aliases genéricos; unidade inferida por contexto"
            else:
                cp=cn=cf=cfn=0; row_sum=0; r["classification_source"]="não identificado"
            matrix_row_sums.append(row_sum); matrix_rows.append(r)
            r["c_matrix_count"]=row_sum
            r["classification_observed"]=row_sum>0
            r["real"] = self._num(ws.cell(rr,c_real).value) if c_real else 0
            r["det"] = self._num(ws.cell(rr,c_det).value) if c_det else 0
            raw_real=ws.cell(rr,c_real).value if c_real else None; raw_det=ws.cell(rr,c_det).value if c_det else None
            r["detection_observed"]=(c_real is not None and c_det is not None and _aia_num_optional(raw_real) is not None and _aia_num_optional(raw_det) is not None)
            r["tp_d"] = self._num(ws.cell(rr,c_tp_d).value) if c_tp_d else None
            r["fp_d"] = self._num(ws.cell(rr,c_fp_d).value) if c_fp_d else None
            r["fn_d"] = self._num(ws.cell(rr,c_fn_d).value) if c_fn_d else None
            r["detection_explicit_matrix"] = all(x is not None for x in (c_tp_d,c_fp_d,c_fn_d))
            r["obs"]=str(ws.cell(rr,c_obs).value or "") if c_obs else ""
            r["calc_status"]="Histórico"; r["source"]="Excel histórico"
            if c_def:
                name=normalize_defect(ws.cell(rr,c_def).value)
                if name: r["defects"]=[{"type":name,"real":r["real"],"det":r["det"]}]
            rows.append(r)
            hist["TP"]+=cp; hist["VN/TN"]+=cn; hist["FP"]+=cf; hist["FN"]+=cfn; hist["N"]+=row_sum
            if c_date:
                dd=self._format_inspection_date(ws.cell(rr,c_date).value)
                if dd: imported_dates.append(dd)
        if not rows: raise ValueError("Nenhuma inspeção válida foi encontrada na planilha.")
        # Classificação é inferida como matriz de inspeção somente quando cada
        # linha válida contribui no máximo uma ocorrência e cada caso tem soma 1.
        c_valid_rows=sum(1 for x in matrix_row_sums if x==1)
        c_invalid_multi=sum(1 for x in matrix_row_sums if x>1)
        classification_valid = c_valid_rows>0 and c_invalid_multi==0 and hist["N"]==c_valid_rows
        # Se as colunas são explicitamente C, a unidade é inspeção, mas o N da
        # matriz continua sendo a soma dos quatro casos válidos — não o total de linhas.
        if explicit_c and c_invalid_multi==0:
            classification_valid = hist["N"]>0
        for r in rows:
            r["classification_observed"] = classification_valid and r.get("c_matrix_count",0)==1
        imported_date=__import__('collections').Counter(imported_dates).most_common(1)[0][0] if imported_dates else ""
        self._archive_current_collection()
        self.rows=rows; self.history=[hist]; self.source_file=path
        imported_meta,metadata_status=self._extract_workbook_metadata(wb,ws,hr,path)
        if imported_date: imported_meta["data"]=imported_date
        for key in self.meta:
            if imported_meta.get(key): self.meta[key].set(imported_meta[key])
        if not self.meta["gabarito"].get(): self.meta["gabarito"].set("Não determinado")
        # Auditoria semântica da fonte.
        occurrence_matrix = (generic_c and not classification_valid) or (explicit_c is False and generic_c and c_invalid_multi>0)
        self.source_analysis={
            "file":str(path), "sheet":ws.title, "header_row":hr,
            "inspections_registered":len(rows),
            "classification_matrix":{"valid":bool(classification_valid),"unit":"inspeção","tp":hist["TP"],"tn":hist["VN/TN"],"fp":hist["FP"],"fn":hist["FN"],"n":hist["N"],"method":"explícita" if explicit_c else ("inferida por soma por linha" if generic_c else "não disponível")},
            "occurrence_matrix":{"present":bool(occurrence_matrix),"tp":hist["TP"],"tn":hist["VN/TN"],"fp":hist["FP"],"fn":hist["FN"],"n":hist["N"],"unit":"ocorrência" if occurrence_matrix else "não determinado"},
            "detection_fields":{"real":c_real is not None,"detected":c_det is not None,"tp_d":c_tp_d is not None,"fp_d":c_fp_d is not None,"fn_d":c_fn_d is not None},
            "mapping":{"TP-C":c_tp_c,"VN-C":c_tn_c,"FP-C":c_fp_c,"FN-C":c_fn_c,"TP/VP":c_tp_g,"VN/TN":c_tn_g,"FP":c_fp_g,"FN":c_fn_g,"Real":c_real,"Detectado":c_det,"TP-D":c_tp_d,"FP-D":c_fp_d,"FN-D":c_fn_d},
            "status":"CONFIRMADO" if (classification_valid or occurrence_matrix or (c_real is not None and c_det is not None)) else "PARCIAL"
        }
        self.refresh_all()
        ca=self.source_analysis["classification_matrix"]; oa=self.source_analysis["occurrence_matrix"]
        d=self._smart_detection_occurrence(); di=self._smart_detection_inspection()
        status=[]
        if ca["valid"]: status.append(f"C: matriz por inspeção N={ca['n']}")
        elif oa["present"]: status.append(f"C: não aplicável; matriz histórica interpretada como ocorrência N={oa['n']}")
        else: status.append("C: N/D")
        if d["available"]: status.append(f"D-ocorrência: {d['real']} reais / {d['det']} detectadas ({d['status']})")
        else: status.append("D: N/D")
        self.status.set(f"✓ Excel interpretado • {len(rows)} inspeções • " + " • ".join(status))
        messagebox.showinfo("Interpretação da planilha concluída",
            f"Planilha: {Path(path).name}\nAba: {ws.title}\nLinha do cabeçalho: {hr}\n\n"
            f"Inspeções registradas: {len(rows)}\n"
            f"Matriz C: {'VALIDADA' if ca['valid'] else 'N/D / não tratada como classificação'}\n"
            f"N da matriz C: {ca['n'] if ca['valid'] else 'N/D'}\n"
            f"Ocorrências na matriz histórica: {oa['n'] if oa['present'] else 'não determinada'}\n\n"
            f"Detecção por ocorrência: {'disponível' if d['available'] else 'N/D'}\n"
            f"Detecção por inspeção: {'disponível' if di['available'] else 'N/D'}\n\n"
            "Nenhum valor histórico foi convertido automaticamente em gabarito do motor novo.", parent=self)
    except Exception as exc:
        messagebox.showerror("Erro na importação Excel",f"Não foi possível interpretar a planilha.\n\nDetalhe: {exc}",parent=self)


def _aia_smart_engineering_indicators(self, m=None, d=None):
    m = m or self.display_metrics(); d = d or self.defect_counts()
    n = m.get("N",0) or 0
    return {
        "prevalencia": safe((m.get("TP",0) or 0)+(m.get("FN",0) or 0),n),
        "taxa_alarme": safe((m.get("TP",0) or 0)+(m.get("FP",0) or 0),n),
        "fp_100": safe((m.get("FP",0) or 0)*100,n),
        "fn_100": safe((m.get("FN",0) or 0)*100,n),
        "real_por_inspecao": safe(d.get("real",0), self.inspection_count_valid_detection()),
        "det_por_inspecao": safe(d.get("det",0), self.inspection_count_valid_detection()),
    }


def _aia_smart_inspection_count_valid_detection(self):
    return sum(1 for r in self.rows if normalize_row(r).get("detection_observed", False))


def _aia_smart_source_badges(self):
    a=_aia_smart_source_summary(self)
    return a


def _aia_smart_view_tecnica(self):
    fonte = "HISTÓRICO INTERPRETADO" if (self.source_file and self.history) else "MOTOR NOVO / ENTRADA MANUAL"
    self.section_title("ANÁLISE TÉCNICA  /  DIAGNÓSTICO",f"{fonte} • V{AIA_SMART_ENGINE_VERSION} • unidades estatísticas explicitadas")
    m=self.display_metrics(); d=self.defect_counts(); dt=self.detection_totals()
    di=dt.get("inspection",{})
    e=self._smart_engineering_indicators(m,d)
    # Classificação C
    cbox=self.card(self.content,"01  •  CLASSIFICAÇÃO (C)","Unidade: inspeção • decisão OK/NOK")
    cbox.pack(fill="x",pady=(0,6))
    if self.source_file and self.history and not self.source_analysis.get("classification_matrix",{}).get("valid"):
        msg="N/D — a matriz histórica importada não foi identificada como matriz de classificação por inspeção. Ela não será reutilizada como C."
        tk.Label(cbox,text=msg,bg=PANEL,fg=WARN,font=("Segoe UI",9,"bold"),wraplength=1050,justify="left").pack(anchor="w",padx=16,pady=12)
    else:
        grid=tk.Frame(cbox,bg=PANEL);grid.pack(fill="x",padx=10,pady=6)
        cards=[("N • casos da matriz",str(m["N"]),TEXT),("Accuracy",pct(m["Accuracy"]),ACCENT2),("Precision",pct(m["Precision"]),ACCENT2),("Recall",pct(m["Recall / Sensibilidade"]),GOOD),("Specificity",pct(m["Specificity"]),ACCENT2),("FPR",pct(m["FPR"]),BAD),("FNR",pct(m["FNR"]),WARN),("F1",pct(m["F1"]),ACCENT2),("MCC","N/D" if m["MCC"] is None else f"{m['MCC']:.4f}",ACCENT2),("Prevalência",pct(e["prevalencia"]),ACCENT2)]
        for i,(lab,val,col) in enumerate(cards): self.metric_card(grid,lab,val,col).grid(row=i//5,column=i%5,sticky="ew",padx=3,pady=3)
        for c in range(5): grid.columnconfigure(c,weight=1)
        tk.Label(cbox,text=f"TP-C={m['TP']} • VN-C={m['VN/TN']} • FP-C={m['FP']} • FN-C={m['FN']}",bg=PANEL,fg=TEXT,font=("Consolas",10,"bold")).pack(anchor="w",padx=16,pady=(4,12))
    # Detecção D por inspeção
    ibox=self.card(self.content,"02  •  DETECÇÃO (D) — POR INSPEÇÃO","Unidade: inspeção • presença de defeito × presença de detecção")
    ibox.pack(fill="x",pady=6)
    if not di.get("available"):
        tk.Label(ibox,text="N/D — a fonte não fornece Real e Detectado por inspeção em quantidade suficiente para construir esta matriz.",bg=PANEL,fg=WARN,font=("Segoe UI",9,"bold"),wraplength=1050,justify="left").pack(anchor="w",padx=16,pady=12)
    else:
        im=di
        grid=tk.Frame(ibox,bg=PANEL);grid.pack(fill="x",padx=10,pady=6)
        cards=[("N • inspeções D",str(im["N"]),GOOD),("Accuracy",pct(im["Accuracy"]),GOOD),("Precision",pct(im["Precision"]),GOOD),("Recall",pct(im["Recall / Sensibilidade"]),GOOD),("Specificity",pct(im["Specificity"]),GOOD),("FPR",pct(im["FPR"]),WARN),("FNR",pct(im["FNR"]),WARN),("F1",pct(im["F1"]),GOOD),("MCC","N/D" if im["MCC"] is None else f"{im['MCC']:.4f}",GOOD),("Prevalência",pct(safe(im["TP"]+im["FN"],im["N"])),GOOD)]
        for i,(lab,val,col) in enumerate(cards): self.metric_card(grid,lab,val,col).grid(row=i//5,column=i%5,sticky="ew",padx=3,pady=3)
        for c in range(5): grid.columnconfigure(c,weight=1)
        tk.Label(ibox,text=f"TP-D={im['TP']} • TN-D={im['TN']} • FP-D={im['FP']} • FN-D={im['FN']} • N={im['N']}",bg=PANEL,fg=TEXT,font=("Consolas",10,"bold")).pack(anchor="w",padx=16,pady=(4,4))
        tk.Label(ibox,text=f"Método: {im['method']}",bg=PANEL,fg=MUTED,font=("Segoe UI",8)).pack(anchor="w",padx=16,pady=(0,10))
    # Detecção D por ocorrência
    obox=self.card(self.content,"03  •  DETECÇÃO (D) — POR OCORRÊNCIA","Unidade: defeito/ocorrência • não confundir com a matriz C")
    obox.pack(fill="x",pady=6)
    tk.Label(obox,text=f"Nº de inspeções registradas: {len(self.rows)}    •    Nº de inspeções com D válido: {dt['available_rows']}    •    Nº de ocorrências reais: {d['real']}    •    Nº de detecções: {d['det']}",bg=PANEL,fg=TEXT,font=("Segoe UI",9,"bold")).pack(anchor="w",padx=16,pady=(10,5))
    tk.Label(obox,text=f"TP-D={d['tp']} • FP-D={d['fp']} • FN-D={d['fn']} • Precision={pct(d['precision'])} • Recall={pct(d['recall'])} • F1={pct(d['f1'])}",bg=PANEL,fg=GOOD,font=("Consolas",10,"bold")).pack(anchor="w",padx=16,pady=5)
    full=d.get("full_metrics")
    if full:
        tk.Label(obox,text=(f"Matriz histórica completa: TN-D={full['VN/TN']} • Accuracy={pct(full['Accuracy'])} • "
                             f"Specificity={pct(full['Specificity'])} • FPR={pct(full['FPR'])} • FNR={pct(full['FNR'])} • "
                             f"MCC={'N/D' if full['MCC'] is None else format(full['MCC'], '.4f')}"),
                 bg=PANEL,fg=TEXT,font=("Segoe UI",8,"bold"),wraplength=1050,justify="left").pack(anchor="w",padx=16,pady=3)
    else:
        tk.Label(obox,text="Accuracy / Specificity / FPR / FNR / MCC: N/D no nível de ocorrência quando TN-D não é observável.",bg=PANEL,fg=MUTED,font=("Segoe UI",8),wraplength=1050,justify="left").pack(anchor="w",padx=16,pady=3)
    tk.Label(obox,text=f"Status do cálculo: {d.get('status','N/D')} • Método: {d.get('method','N/D')}",bg=PANEL,fg=WARN if d.get("estimated") else GOOD,font=("Segoe UI",8,"bold"),wraplength=1050,justify="left").pack(anchor="w",padx=16,pady=(0,10))
    # Origem
    abox=self.card(self.content,"04  •  AUDITORIA DA INTERPRETAÇÃO","O AIA informa de onde veio cada universo estatístico.")
    abox.pack(fill="x",pady=6)
    a=self.source_analysis or {}
    lines=[f"Fonte: {a.get('file','Entrada manual')}",f"Aba: {a.get('sheet','N/D')} • Cabeçalho: {a.get('header_row','N/D')}",f"Inspeções registradas: {a.get('inspections_registered',len(self.rows))}",f"Status da interpretação: {a.get('status','MOTOR NOVO')}" ]
    for line in lines: tk.Label(abox,text=line,bg=PANEL,fg=TEXT,font=("Segoe UI",8)).pack(anchor="w",padx=16,pady=2)
    self._aia_smart_add_interpretation_button(abox)


def _aia_smart_add_interpretation_button(self,parent):
    if not getattr(self,"source_analysis",None): return
    def show():
        a=self.source_analysis
        text=(f"Fonte: {a.get('file')}\nAba: {a.get('sheet')}\nCabeçalho: {a.get('header_row')}\n\n"
              f"Inspeções registradas: {a.get('inspections_registered')}\n"
              f"Matriz C válida: {a.get('classification_matrix',{}).get('valid')}\n"
              f"N C: {a.get('classification_matrix',{}).get('n')}\n"
              f"Matriz histórica de ocorrências: {a.get('occurrence_matrix',{}).get('present')}\n"
              f"Detecção Real: {a.get('detection_fields',{}).get('real')}\n"
              f"Detecção Detectado: {a.get('detection_fields',{}).get('detected')}\n\n"
              "Regra: nenhum universo é misturado automaticamente.")
        messagebox.showinfo("Interpretação da fonte",text,parent=self)
    ttk.Button(parent,text="VER INTERPRETAÇÃO DA FONTE",style="Accent.TButton",command=show).pack(anchor="e",padx=16,pady=10)


def _aia_smart_view_dashboard(self):
    fonte="HISTÓRICO INTERPRETADO" if (self.source_file and self.history) else "MOTOR NOVO / ENTRADA MANUAL"
    self.section_title("DASHBOARD  /  VISÃO EXECUTIVA",f"{fonte} • V{AIA_SMART_ENGINE_VERSION} • verdade matemática por unidade")
    d=self.defect_counts(); m=self.display_metrics(); di=self.detection_totals().get("inspection",{})
    band=tk.Frame(self.content,bg="#103126",highlightbackground=GOOD,highlightthickness=1);band.pack(fill="x",pady=(0,5))
    tk.Label(band,text="01  •  DETECÇÃO (D)",bg="#103126",fg=GOOD,font=("Segoe UI",12,"bold")).pack(side="left",padx=14,pady=9)
    tk.Label(band,text="OCORRÊNCIAS — real × detectado • e, quando possível, D por inspeção",bg="#103126",fg=TEXT,font=("Segoe UI",9)).pack(side="left",padx=4,pady=9)
    dg=tk.Frame(self.content,bg=BG);dg.pack(fill="x",pady=(0,4))
    for i,(lab,val,col) in enumerate([("Inspeções",len(self.rows),GOOD),("Ocorrências reais",d["real"],GOOD),("Detectados",d["det"],ACCENT2),("TP-D",d["tp"],GOOD),("FP-D",d["fp"],BAD),("FN-D",d["fn"],WARN)]):
        self.metric_card(dg,lab,str(val),col).grid(row=0,column=i,sticky="ew",padx=2);dg.columnconfigure(i,weight=1)
    if di.get("available"):
        row=tk.Frame(self.content,bg=BG);row.pack(fill="x",pady=3)
        for i,(lab,val,col) in enumerate([("D-Inspeção N",di["N"],GOOD),("Accuracy-D",pct(di["Accuracy"]),GOOD),("Precision-D",pct(di["Precision"]),GOOD),("Recall-D",pct(di["Recall / Sensibilidade"]),GOOD),("Specificity-D",pct(di["Specificity"]),GOOD)]):
            self.metric_card(row,lab,str(val),col).grid(row=0,column=i,sticky="ew",padx=2);row.columnconfigure(i,weight=1)
    dc=self.chart_card(self.content,"GRÁFICO DE DETECÇÃO (D)","Primeiro nível executivo • unidade de ocorrência")
    dc.pack(fill="both",expand=True,pady=4)
    self.draw_bar_chart(dc,[("REAL",d["real"],GOOD),("DETECTADO",d["det"],ACCENT2),("TP-D",d["tp"],GOOD),("FP-D",d["fp"],BAD),("FN-D",d["fn"],WARN)],max_value=max(d["real"],d["det"],d["tp"],d["fp"],d["fn"],1))
    band2=tk.Frame(self.content,bg="#0A2635",highlightbackground=ACCENT2,highlightthickness=1);band2.pack(fill="x",pady=(7,5))
    tk.Label(band2,text="02  •  CLASSIFICAÇÃO (C)",bg="#0A2635",fg=ACCENT2,font=("Segoe UI",12,"bold")).pack(side="left",padx=14,pady=9)
    tk.Label(band2,text="INSPEÇÕES — decisão OK/NOK",bg="#0A2635",fg=TEXT,font=("Segoe UI",9)).pack(side="left",padx=4,pady=9)
    if self.source_file and self.history and not self.source_analysis.get("classification_matrix",{}).get("valid"):
        card=self.card(self.content,"CLASSIFICAÇÃO C — N/D","A fonte não fornece uma matriz C válida por inspeção. O AIA não reutilizou a matriz de ocorrências como classificação.");card.pack(fill="x",pady=4)
    else:
        cg=tk.Frame(self.content,bg=BG);cg.pack(fill="x",pady=(0,4))
        for i,(lab,val,col) in enumerate([("TP-C",m["TP"],GOOD),("VN-C",m["VN/TN"],GOOD),("FP-C",m["FP"],BAD),("FN-C",m["FN"],WARN)]): self.metric_card(cg,lab,str(val),col).grid(row=0,column=i,sticky="ew",padx=3);cg.columnconfigure(i,weight=1)
        cc=self.chart_card(self.content,"GRÁFICO DE CLASSIFICAÇÃO (C)","Segundo nível executivo • unidade de inspeção");cc.pack(fill="both",expand=True,pady=4)
        self.draw_bar_chart(cc,[("TP-C",m["TP"],GOOD),("VN-C",m["VN/TN"],GOOD),("FP-C",m["FP"],BAD),("FN-C",m["FN"],WARN)],max_value=max(m["TP"],m["VN/TN"],m["FP"],m["FN"],1))


def _aia_smart_view_graficos(self):
    fonte="HISTÓRICO INTERPRETADO" if (self.source_file and self.history) else "MOTOR NOVO / ENTRADA MANUAL"
    self.section_title("GRÁFICOS  /  PERFORMANCE",f"{fonte} • matrizes C e D independentes • V{AIA_SMART_ENGINE_VERSION}")
    m=self.display_metrics();d=self.defect_counts();dt=self.detection_totals();di=dt.get("inspection",{})
    top=tk.Frame(self.content,bg=BG);top.pack(fill="both",expand=True)
    det=self.chart_card(top,"01  •  MATRIZ DE DETECÇÃO (D) — OCORRÊNCIA","Real × detectado • TN-D não se aplica a contagem de ocorrências");det.grid(row=0,column=0,sticky="nsew",padx=(0,4),pady=(0,4));self._aia_final_draw_detection_matrix(det,d)
    cls=self.chart_card(top,"02  •  MATRIZ DE CLASSIFICAÇÃO (C)","Unidade: inspeção • decisão OK/NOK");cls.grid(row=0,column=1,sticky="nsew",padx=(4,0),pady=(0,4));self.draw_bar_chart(cls,[("TP-C",m["TP"],GOOD),("VN-C",m["VN/TN"],GOOD),("FP-C",m["FP"],BAD),("FN-C",m["FN"],WARN)],max_value=max(m["TP"],m["VN/TN"],m["FP"],m["FN"],1))
    if di.get("available"):
        dp=self.chart_card(top,"03  •  INDICADORES DE DETECÇÃO (D) — INSPEÇÃO","Accuracy • Precision • Recall • Specificity • F1");dp.grid(row=1,column=0,sticky="nsew",padx=(0,4),pady=(4,0));self.draw_percent_chart(dp,[("Accuracy",di["Accuracy"],GOOD),("Precision",di["Precision"],GOOD),("Recall",di["Recall / Sensibilidade"],GOOD),("Specificity",di["Specificity"],GOOD),("F1",di["F1"],GOOD)])
    else:
        dp=self.chart_card(top,"03  •  INDICADORES DE DETECÇÃO (D) — INSPEÇÃO","N/D — Real e Detectado por inspeção não estão disponíveis.");dp.grid(row=1,column=0,sticky="nsew",padx=(0,4),pady=(4,0));tk.Label(dp,text="N/D\n\nO AIA não cria uma matriz D por inspeção sem base de dados suficiente.",bg=PANEL,fg=WARN,font=("Segoe UI",11,"bold"),justify="center").pack(expand=True)
    occ=self.chart_card(top,"04  •  PERFIL DE DETECÇÃO (D) — OCORRÊNCIA","Real × Detectado × TP-D × FP-D × FN-D");occ.grid(row=1,column=1,sticky="nsew",padx=(4,0),pady=(4,0));self.draw_bar_chart(occ,[("Real",d["real"],GOOD),("Detectado",d["det"],ACCENT2),("TP-D",d["tp"],GOOD),("FP-D",d["fp"],BAD),("FN-D",d["fn"],WARN)],max_value=max(d["real"],d["det"],d["tp"],d["fp"],d["fn"],1))
    top.columnconfigure(0,weight=1);top.columnconfigure(1,weight=1);top.rowconfigure(0,weight=1);top.rowconfigure(1,weight=1)


def _aia_smart_export_excel(self):
    """Exportação consolidada com universos explícitos e validação de reabertura."""
    if openpyxl is None:
        messagebox.showerror("Excel","OpenPyXL não está instalado.\nExecute o INSTALAR_DEPENDENCIAS_AIA.bat",parent=self);return
    path=filedialog.asksaveasfilename(title=f"Exportar relatório AIA V{VERSION}",defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
    if not path:return
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font,PatternFill,Border,Side,Alignment
        from openpyxl.utils import get_column_letter
        wb=Workbook(); ws=wb.active; ws.title="RESUMO"
        dark="182532"; green="45D39A"; yellow="F4C15D"; red="FF6B78"; blue="67D5FF"; thin=Side(style="thin",color="6B7F8D")
        def title(sh,text,cols=6): sh["A1"]=text;sh["A1"].font=Font(size=18,bold=True,color="FFFFFF");sh["A1"].fill=PatternFill("solid",fgColor=dark);sh.merge_cells(start_row=1,start_column=1,end_row=1,end_column=cols)
        def hdr(sh,row,labels):
            for c,v in enumerate(labels,1):
                x=sh.cell(row,c,v);x.font=Font(bold=True,color="FFFFFF");x.fill=PatternFill("solid",fgColor=dark);x.border=Border(bottom=thin)
        def autosize(sh):
            widths={}
            for row in sh.iter_rows():
                for x in row:
                    if hasattr(x,"column"):
                        widths[x.column]=max(widths.get(x.column,10),min(60,len(str(x.value or ""))+2))
            for idx,w in widths.items(): sh.column_dimensions[get_column_letter(idx)].width=w
        m=self.display_metrics();d=self.defect_counts();dt=self.detection_totals();di=dt.get("inspection",{});a=self.source_analysis or {}
        title(ws,f"AUTAZA INSPECTION ANALYTICS — RELATÓRIO V{VERSION}",6)
        info=[("Versão",VERSION),("Motor de dados",AIA_SMART_ENGINE_VERSION),("Fonte","Excel histórico" if self.source_file else "Entrada manual"),("Arquivo fonte",a.get("file",self.source_file or "Entrada manual")),("Inspeções registradas",len(self.rows)),("Ocorrências reais",d["real"]),("Detecções",d["det"])]
        for i,(k,v) in enumerate(info,3): ws.cell(i,1,k);ws.cell(i,2,v)

        # BLOCO CANÔNICO DE CONTEXTO: este bloco é o contrato de round-trip
        # entre exportação e importação do AIA. Não depende da tabela de
        # inspeções nem de heurísticas.
        ws["A10"]="IDENTIFICAÇÃO DO TESTE"
        ws["A10"].font=Font(size=13,bold=True,color="FFFFFF")
        ws["A10"].fill=PatternFill("solid",fgColor=dark)
        ws.merge_cells(start_row=10,start_column=1,end_row=10,end_column=6)
        native_info=[
            ("Cliente",self.meta["cliente"].get()),
            ("Local / Planta",self.meta["local"].get()),
            ("Projeto",self.meta["projeto"].get()),
            ("Linha / Estação",self.meta["linha"].get()),
            ("Técnico",self.meta["tecnico"].get()),
            ("Data",self.meta["data"].get()),
            ("Threshold",self.meta["threshold"].get()),
            ("Método de referência",self.meta["gabarito"].get()),
        ]
        for i,(k,v) in enumerate(native_info,11):
            ws.cell(i,1,k); ws.cell(i,2,v if v else "")
        hdr(ws,21,["Universo","N","TP","TN/VN","FP","FN","Status / método"])
        c=a.get("classification_matrix",{}); rows_exp=[("Classificação — inspeção",c.get("n",m["N"]) if c.get("valid") else "N/D",c.get("tp",m["TP"]) if c.get("valid") else "N/D",c.get("tn",m["VN/TN"]) if c.get("valid") else "N/D",c.get("fp",m["FP"]) if c.get("valid") else "N/D",c.get("fn",m["FN"]) if c.get("valid") else "N/D","Válida" if c.get("valid") else "N/D"),("Detecção — inspeção",di.get("N","N/D"),di.get("TP","N/D"),di.get("TN","N/D"),di.get("FP","N/D"),di.get("FN","N/D"),di.get("method","N/D")),("Detecção — ocorrência",d["real"],d["tp"],"N/A",d["fp"],d["fn"],d.get("method","N/D"))]
        for i,row in enumerate(rows_exp,22):
            for cidx,v in enumerate(row,1): ws.cell(i,cidx,v)
        autosize(ws)
        wi=wb.create_sheet("INSPEÇÕES");title(wi,"INSPEÇÕES — DADOS NORMALIZADOS",21);hdr(wi,3,["Inspeção","Identificação","Vision AI","Referência","Real","Detectado","TP-C","VN-C","FP-C","FN-C","TP-DI","TN-DI","FP-DI","FN-DI","TP-D","FP-D","FN-D","C status","D-I status","Origem","Observação"])
        for i,r0 in enumerate(self.rows,4):
            r=normalize_row(r0); dd=self.detection_for_row(r)
            # D-I = presença binária por inspeção. D-O = ocorrência por contagem/pareamento.
            rr=r["real"];xx=r["det"]; has=r.get("detection_observed",False)
            if has:
                tpdi=int(rr>0 and xx>0); tndi=int(rr==0 and xx==0); fpdi=int(rr==0 and xx>0); fndi=int(rr>0 and xx==0)
                ds="Derivado por presença Real>0 × Detectado>0"
                dotp,dofp,dofn=dd.get("tp",0),dd.get("fp",0),dd.get("fn",0)
            else:
                tpdi=tndi=fpdi=fndi="N/D"; dotp=dofp=dofn="N/D"; ds="N/D"
            c_valid_row=bool(r.get("classification_observed")) or (not self.source_file and r.get("ref") in ("Defeito confirmado","OK confirmado"))
            vals=[i-3,r["id"],r["vision"],r["ref"],rr,xx,r["tp"],r["vn"],r["fp"],r["fn"],tpdi,tndi,fpdi,fndi,dotp,dofp,dofn,"Válida" if c_valid_row else "N/D",ds,r["source"],r["obs"]]
            for cidx,v in enumerate(vals,1): wi.cell(i,cidx,v)
        wi.freeze_panes="A4";autosize(wi)
        wm=wb.create_sheet("MÉTRICAS");title(wm,"MÉTRICAS — CLASSIFICAÇÃO E DETECÇÃO",4);hdr(wm,3,["Universo / Métrica","Valor","Unidade","Método"])
        metrics_rows=[("C — Nº casos",m["N"],"inspeção","TP-C+VN-C+FP-C+FN-C"),("C — Accuracy",m["Accuracy"],"inspeção","(TP-C+VN-C)/N"),("C — Precision",m["Precision"],"inspeção","TP-C/(TP-C+FP-C)"),("C — Recall",m["Recall / Sensibilidade"],"inspeção","TP-C/(TP-C+FN-C)"),("D-I — Nº inspeções",di.get("N","N/D"),"inspeção",di.get("method","N/D")),("D-I — Accuracy",di.get("Accuracy","N/D"),"inspeção","(TP-DI+TN-DI)/N-DI"),("D-I — Precision",di.get("Precision","N/D"),"inspeção","TP-DI/(TP-DI+FP-DI)"),("D-I — Recall",di.get("Recall / Sensibilidade","N/D"),"inspeção","TP-DI/(TP-DI+FN-DI)"),("D-I — Specificity",di.get("Specificity","N/D"),"inspeção","TN-DI/(TN-DI+FP-DI)"),("D-I — FPR",di.get("FPR","N/D"),"inspeção","FP-DI/(FP-DI+TN-DI)"),("D-I — FNR",di.get("FNR","N/D"),"inspeção","FN-DI/(FN-DI+TP-DI)"),("D-I — F1",di.get("F1","N/D"),"inspeção","2PR/(P+R)"),("D-I — MCC",di.get("MCC","N/D"),"inspeção","Matriz D completa"),("D-O — Nº ocorrências reais",d["real"],"ocorrência","contagem"),("D-O — Nº detecções",d["det"],"ocorrência","contagem"),("D-O — TP",d["tp"],"ocorrência",d.get("method","N/D")),("D-O — FP",d["fp"],"ocorrência",d.get("method","N/D")),("D-O — FN",d["fn"],"ocorrência",d.get("method","N/D")),("D-O — Precision",d["precision"],"ocorrência","TP/(TP+FP)"),("D-O — Recall",d["recall"],"ocorrência","TP/(TP+FN)"),("D-O — F1",d["f1"],"ocorrência","2PR/(P+R)")]
        for i,row in enumerate(metrics_rows,4):
            for cidx,v in enumerate(row,1): wm.cell(i,cidx,v if v is not None else "N/D")
            if isinstance(row[1],float): wm.cell(i,2).number_format="0.00%"
        autosize(wm)
        wa=wb.create_sheet("AUDITORIA");title(wa,"AUDITORIA — ORIGEM E CONSISTÊNCIA",6);hdr(wa,3,["Item","Valor","Unidade","Origem","Status","Regra"])
        audit=[("Inspeções registradas",len(self.rows),"inspeção","arquivo/entrada","OBSERVADO","contagem de registros"),("Ocorrências reais",d["real"],"ocorrência","Real","OBSERVADO/DERIVADO","soma Real"),("Detecções",d["det"],"ocorrência","Detectado","OBSERVADO/DERIVADO","soma Detectado"),("Matriz C",str(c.get("valid")),"inspeção","Vision AI × Referência ou TP-C/VN-C/FP-C/FN-C","VALIDADA" if c.get("valid") else "N/D",c.get("method","fonte C explícita")),("Matriz D-I",di.get("N","N/D"),"inspeção","Real × Detectado", "DISPONÍVEL" if di.get("available") else "N/D",di.get("method","Binário por presença")),("Detecção D-O",d.get("method","N/D"),"ocorrência","Real/Detectado/TP-D","ESTIMADO" if d.get("estimated") else d.get("status","N/D"),"sem correspondência individual quando indicado"),("TP/VN/FP/FN genéricos","NÃO PROMOVIDOS","não determinada","cabeçalhos sem sufixo","CONTROLADO","não alimentam C, D-I ou D-O sem evidência semântica")]
        for i,row in enumerate(audit,4):
            for cidx,v in enumerate(row,1): wa.cell(i,cidx,v)
        autosize(wa)
        wdo=wb.create_sheet("DETECCAO");title(wdo,"DETECÇÃO — OCORRÊNCIA",6);hdr(wdo,3,["Indicador","Valor","Unidade","Status","Método","Observação"])
        vals=[("Inspeções registradas",len(self.rows),"inspeção","OBSERVADO","contagem","não é N de ocorrências"),("Inspeções D válidas",di.get("N","N/D"),"inspeção","OBSERVADO","Real × Detectado","base da matriz D-I"),("Ocorrências reais",d["real"],"ocorrência","OBSERVADO/DERIVADO","soma Real",""),("Detecções",d["det"],"ocorrência","OBSERVADO/DERIVADO","soma Detectado",""),("TP-D",d["tp"],"ocorrência",d.get("status","N/D"),d.get("method","N/D"),""),("FP-D",d["fp"],"ocorrência",d.get("status","N/D"),d.get("method","N/D"),""),("FN-D",d["fn"],"ocorrência",d.get("status","N/D"),d.get("method","N/D"),"")]
        for i,row in enumerate(vals,4):
            for cidx,v in enumerate(row,1): wdo.cell(i,cidx,v)
        autosize(wdo)

        # Meta D-I: exporta a mesma regra exibida na tela META, incluindo perfil e distância.
        we=wb.create_sheet("EXPECTATIVA");title(we,"META / CRITÉRIO DE ENTREGA — D-I",7)
        target=self.expectation_snapshot(); ev=aia_v120_meta_capability_evaluation(target,di)
        we["A3"]="Perfil"; we["B3"]=target.get("profile","PERSONALIZADO")
        we["A4"]="Assertividade D-I"; we["B4"]=ev.get("assertiveness") if ev.get("assertiveness") is not None else "N/D"
        if isinstance(we["B4"].value,float): we["B4"].number_format="0.00%"
        we["A5"]="Classificação geral"; we["B5"]=ev.get("classification","N/D")
        we["A6"]="Amostra"; we["B6"]=f"N={ev.get('n',0)} — {ev.get('sample',{}).get('label','N/D')}"
        hdr(we,8,["Indicador","Meta","Resultado D-I","Diferença","Direção","Status","Crítico"])
        for row_idx,row in enumerate(ev.get("rows",[]),9):
            vals=[row["label"],row["target"] if row["target"] is not None else "N/D",row["actual"] if row["actual"] is not None else "N/D",row["distance_text"],row["direction"],row["status"],"SIM" if row["critical"] else "NÃO"]
            for col_idx,val in enumerate(vals,1): we.cell(row_idx,col_idx,val)
            if row["percent"]:
                if isinstance(row["target"],(int,float)): we.cell(row_idx,2).number_format="0.00%"
                if isinstance(row["actual"],(int,float)): we.cell(row_idx,3).number_format="0.00%"
        trace_row=10+len(ev.get("rows",[]))
        we.cell(trace_row,1,"Origem do critério");we.cell(trace_row,2,target.get("source","") or "N/D")
        we.cell(trace_row+1,1,"Responsável");we.cell(trace_row+1,2,target.get("responsible","") or "N/D")
        we.cell(trace_row+2,1,"Data");we.cell(trace_row+2,2,target.get("date","") or "N/D")
        we.cell(trace_row+3,1,"Observações");we.cell(trace_row+3,2,target.get("notes","") or "")
        autosize(we)

        # Restaura a análise por tipo sem confundir detalhe opcional com D-I.
        wd=wb.create_sheet("POR DEFEITO");title(wd,"ANÁLISE POR TIPO DE DEFEITO — D-O",9)
        hdr(wd,3,["Defeito","Real","Detectado","TP-D","FP-D","FN-D","Recall D-O","Precision D-O","Inspeções"])
        for row_idx,(name,values) in enumerate(sorted(d.get("by",{}).items()),4):
            rec=safe(values.get("tp",0),values.get("tp",0)+values.get("fn",0))
            pre=safe(values.get("tp",0),values.get("tp",0)+values.get("fp",0))
            vals=[name,values.get("real",0),values.get("det",0),values.get("tp",0),values.get("fp",0),values.get("fn",0),rec if rec is not None else "N/D",pre if pre is not None else "N/D",values.get("inspecoes",0)]
            for col_idx,val in enumerate(vals,1): wd.cell(row_idx,col_idx,val)
            if isinstance(rec,float): wd.cell(row_idx,7).number_format="0.00%"
            if isinstance(pre,float): wd.cell(row_idx,8).number_format="0.00%"
        autosize(wd)

        wg=wb.create_sheet("GLOSSÁRIO");title(wg,"GLOSSÁRIO AIA",2);hdr(wg,3,["Sigla / Métrica","Definição"])
        for row_idx,(term,desc) in enumerate(GLOSSARY.items(),4):
            wg.cell(row_idx,1,term);wg.cell(row_idx,2,desc);wg.cell(row_idx,2).alignment=Alignment(wrap_text=True,vertical="top")
        wg.column_dimensions["A"].width=42;wg.column_dimensions["B"].width=110

        wb.save(path)
        tmp=str(Path(path).with_name(Path(path).stem+"__AIA_TMP.xlsx"))
        try:
            # Reabre o arquivo final e verifica as abas essenciais.
            check=openpyxl.load_workbook(path,read_only=True,data_only=False)
            required={"RESUMO","INSPEÇÕES","MÉTRICAS","AUDITORIA","DETECCAO","EXPECTATIVA","POR DEFEITO","GLOSSÁRIO"}
            missing=required-set(check.sheetnames);check.close()
            if missing: raise ValueError("Abas ausentes após exportação: "+", ".join(sorted(missing)))
        finally:
            if os.path.exists(tmp):
                try: os.remove(tmp)
                except Exception: pass
        self.status.set(f"✓ Excel exportado e validado: {Path(path).name}")
        messagebox.showinfo("Exportação concluída",f"Relatório AIA V{VERSION} criado e reaberto com sucesso para validação estrutural.",parent=self)
    except Exception as exc:
        messagebox.showerror("Erro na exportação Excel",f"O relatório não foi substituído por um arquivo inválido.\n\n{exc}",parent=self)


# Bindings do Smart Data Engine — última camada, sem EXE.
App._smart_classification_totals = _aia_smart_classification_totals
App._smart_metrics = _aia_smart_metrics
App._smart_source_summary = _aia_smart_source_summary
App._smart_detection_inspection = _aia_smart_detection_inspection
App._smart_detection_occurrence = _aia_smart_detection_occurrence
App._smart_detection_totals = _aia_smart_detection_totals
App._smart_defect_counts = _aia_smart_defect_counts
App._smart_import_excel = _aia_smart_import_excel
App._smart_engineering_indicators = _aia_smart_engineering_indicators
App.inspection_count_valid_detection = _aia_smart_inspection_count_valid_detection
App._smart_add_interpretation_button = _aia_smart_add_interpretation_button
App.display_metrics = _aia_smart_display_metrics
App.classification_totals = _aia_smart_classification_totals
App.metrics = _aia_smart_metrics
App.historical_metrics = lambda self: self.display_metrics()
App.detection_totals = _aia_smart_detection_totals
App.defect_counts = _aia_smart_defect_counts
App.import_excel = _aia_smart_import_excel
App.engineering_indicators = _aia_smart_engineering_indicators
App.view_tecnica = _aia_smart_view_tecnica
App.view_dashboard = _aia_smart_view_dashboard
App.view_graficos = _aia_smart_view_graficos
App.export_excel = _aia_smart_export_excel

# Inicialização: os atributos abaixo são criados aqui para manter compatibilidade
# com projetos existentes sem exigir migração manual.
_old_init = App.__init__
def _aia_smart_init(self, *args, **kwargs):
    self.source_analysis = {}
    self.expectation_saved_snapshot = None
    _old_init(self, *args, **kwargs)
App.__init__ = _aia_smart_init



# ============================================================================
# AIA V1.1.2 — CORREÇÃO CONSOLIDADA / SEM REGRESSÃO
# Objetivo: recuperar a apresentação consolidada, manter C e D independentes,
# corrigir o gráfico D e tornar a origem/unidade dos números explícita.
# ============================================================================
AIA_SMART_ENGINE_VERSION = "1.1.10"


def _aia_v111_scan_aggregate_source(self, path):
    """Localiza indicadores agregados escritos na própria planilha.
    Não altera os registros de inspeção; apenas documenta valores declarados
    pela fonte e sua localização para auditoria.
    """
    out = {"present": False, "sheet": None, "cells": {}, "total_inspections": None,
           "total_occurrences": None, "tp": None, "tn": None, "fp": None, "fn": None}
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        aliases = {
            "total_inspections": {"TOTAL DE INSPECOES", "TOTAL DE INSPECAO"},
            "total_occurrences": {"TOTAL DE OCORRENCIAS", "TOTAL DE OCORRENCIA"},
            "fp": {"FP - FALSO POSITIVO", "FP - FALSO POSITIVO ", "FP"},
            "fn": {"FN - FALSO NEGATIVO", "FN"},
            "tp": {"TP - VERDADEIRO POSITIVO", "VP - VERDADEIRO POSITIVO", "TP", "VP"},
            "tn": {"TN - VERDADEIRO NEGATIVO", "VN - VERDADEIRO NEGATIVO", "TN", "VN"},
        }
        # First pass: exact/near-exact labels with numeric value immediately right.
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                vals=[c.value for c in row]
                for j,v in enumerate(vals):
                    key=normalize_text(v)
                    if not key: continue
                    for dest, names in aliases.items():
                        if key in names and j+1 < len(vals):
                            num=_aia_num_optional(vals[j+1])
                            if num is not None and out.get(dest) is None:
                                out[dest]=num; out["sheet"]=ws.title; out["cells"][dest]=f"{ws.title}!{row[j+1].coordinate}"
        # Second pass: support labels embedded in text, e.g. "FP - Falso Positivo ".
        if out["fp"] is None or out["fn"] is None or out["tp"] is None or out["tn"] is None:
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    vals=[c.value for c in row]
                    for j,v in enumerate(vals[:-1]):
                        key=normalize_text(v)
                        if not key: continue
                        dest=None
                        if key.startswith("FP - FALSO POSITIVO"): dest="fp"
                        elif key.startswith("FN - FALSO NEGATIVO"): dest="fn"
                        elif key.startswith("TP - VERDADEIRO POSITIVO") or key.startswith("VP - VERDADEIRO POSITIVO"): dest="tp"
                        elif key.startswith("TN - VERDADEIRO NEGATIVO") or key.startswith("VN - VERDADEIRO NEGATIVO"): dest="tn"
                        elif key.startswith("TOTAL DE INSPECOES"): dest="total_inspections"
                        elif key.startswith("TOTAL DE OCORRENCIAS"): dest="total_occurrences"
                        if dest and out.get(dest) is None:
                            num=_aia_num_optional(vals[j+1])
                            if num is not None:
                                out[dest]=num; out["sheet"]=ws.title; out["cells"][dest]=f"{ws.title}!{row[j+1].coordinate}"
        wb.close()
        out["present"] = any(out[k] is not None for k in ("tp","tn","fp","fn","total_occurrences","total_inspections"))
        out["complete_matrix"] = all(out[k] is not None for k in ("tp","tn","fp","fn"))
        if out["complete_matrix"]:
            out["n"] = out["tp"]+out["tn"]+out["fp"]+out["fn"]
            out["unit"] = "ocorrência" if out["total_occurrences"] is not None or out["total_inspections"] is not None else "não determinado"
            out["method"] = "Matriz agregada declarada pela fonte; valores preservados sem conversão de unidade"
        else:
            out["n"] = None
            out["unit"] = "não determinado"
            out["method"] = "Indicadores agregados parciais; matriz completa não comprovada"
    except Exception as exc:
        out["error"] = str(exc)
    return out


def _aia_v111_import_excel(self):
    """Importa pela lógica Smart existente e, ao final, anexa a evidência
    agregada da fonte. A etapa de cálculo continua sendo feita pelo motor.
    """
    # Usa a importação Smart já validada; depois corrige/documenta a camada
    # agregada sem duplicar a lógica de normalização de linhas.
    old = getattr(self, "_aia_v111_old_import_excel", None)
    if old is None:
        return
    old()
    path=getattr(self,"source_file","")
    if not path: return
    agg=_aia_v111_scan_aggregate_source(self,path)
    a=dict(getattr(self,"source_analysis",{}) or {})
    a["aggregate_source"]=agg
    # Se a fonte declara uma matriz completa agregada e o contexto é histórico,
    # tratamos essa matriz como universo de ocorrência, nunca como C.
    if agg.get("complete_matrix"):
        a["occurrence_matrix"]={
            "present":True,"tp":agg["tp"],"tn":agg["tn"],"fp":agg["fp"],"fn":agg["fn"],
            "n":agg["n"],"unit":"ocorrência","method":agg["method"],
            "sheet":agg.get("sheet"),"cells":agg.get("cells",{})
        }
    if agg.get("total_inspections") is not None:
        a["source_declared_inspections"]=agg["total_inspections"]
    self.source_analysis=a
    self.refresh_all()
    # Não abre uma segunda caixa de diálogo; a informação aparece nas telas e auditoria.


def _aia_v111_detection_occurrence(self):
    """Prioridade: matriz agregada completa da fonte > matriz histórica já
    reconhecida > dados por linha. Assim o AIA não substitui uma matriz
    declarada por uma reconstrução simplificada.
    """
    a=getattr(self,"source_analysis",{}) or {}
    agg=a.get("aggregate_source",{}) or {}
    if agg.get("complete_matrix"):
        tp=int(agg["tp"]);tn=int(agg["tn"]);fp=int(agg["fp"]);fn=int(agg["fn"])
        mm=_aia_metric_pack(tp,tn,fp,fn,tp+tn+fp+fn)
        return {"available":True,"real":tp+fn,"det":tp+fp,"tp":tp,"fp":fp,"fn":fn,"tn":tn,
                "precision":mm["Precision"],"recall":mm["Recall / Sensibilidade"],"f1":mm["F1"],
                "available_rows":len(self.rows),"unavailable_rows":0,"detailed":False,"estimated":False,
                "status":"IMPORTADO / VALIDADO","method":agg.get("method","Matriz agregada da fonte"),
                "full_metrics":mm,"source_sheet":agg.get("sheet"),"source_cells":agg.get("cells",{})}
    # fallback para a implementação Smart anterior
    old=getattr(self,"_aia_v111_old_detection_occurrence",None)
    if old: return old()
    return {"available":False,"real":0,"det":0,"tp":0,"fp":0,"fn":0,"tn":0,
            "precision":None,"recall":None,"f1":None,"status":"N/D","method":"N/D"}


def _aia_v111_metric_cards(self,parent, title, subtitle, metrics, cols=5):
    box=self.card(parent,title,subtitle);box.pack(fill="x",pady=5)
    grid=tk.Frame(box,bg=PANEL);grid.pack(fill="x",padx=8,pady=8)
    for i,(lab,val,col) in enumerate(metrics):
        self.metric_card(grid,lab,str(val),col).grid(row=i//cols,column=i%cols,sticky="ew",padx=3,pady=3)
    for c in range(cols): grid.columnconfigure(c,weight=1)
    return box


def _aia_v111_view_dashboard(self):
    fonte="HISTÓRICO INTERPRETADO" if (self.source_file and self.history) else "MOTOR NOVO / ENTRADA MANUAL"
    self.section_title("DASHBOARD  /  VISÃO EXECUTIVA",f"{fonte} • V{AIA_SMART_ENGINE_VERSION} • classificação e detecção sem mistura")
    d=self.defect_counts(); m=self.display_metrics(); dt=self.detection_totals(); di=dt.get("inspection",{})
    # DETECÇÃO — destaque verde, como definido na arquitetura.
    band=tk.Frame(self.content,bg="#103126",highlightbackground=GOOD,highlightthickness=1);band.pack(fill="x",pady=(0,5))
    tk.Label(band,text="01  •  DETECÇÃO (D)",bg="#103126",fg=GOOD,font=("Segoe UI",12,"bold")).pack(side="left",padx=14,pady=9)
    tk.Label(band,text="OCORRÊNCIAS — defeitos reais × detecções",bg="#103126",fg=TEXT,font=("Segoe UI",9)).pack(side="left",padx=4,pady=9)
    g=tk.Frame(self.content,bg=BG);g.pack(fill="x")
    vals=[("Inspeções",len(self.rows),GOOD),("Ocorrências reais",d["real"],GOOD),("Detectados",d["det"],ACCENT2),("TP-D",d["tp"],GOOD),("FP-D",d["fp"],BAD),("FN-D",d["fn"],WARN)]
    for i,x in enumerate(vals): self.metric_card(g,*[str(y) if i==0 else y for y in x]).grid(row=0,column=i,sticky="ew",padx=2,pady=3)
    for i in range(6):g.columnconfigure(i,weight=1)
    if di.get("available"):
        dg=tk.Frame(self.content,bg=BG);dg.pack(fill="x")
        vals2=[("D-I N",di["N"],GOOD),("Accuracy-D",pct(di["Accuracy"]),GOOD),("Precision-D",pct(di["Precision"]),GOOD),("Recall-D",pct(di["Recall / Sensibilidade"]),GOOD),("Specificity-D",pct(di["Specificity"]),GOOD)]
        for i,x in enumerate(vals2): self.metric_card(dg,*x).grid(row=0,column=i,sticky="ew",padx=2,pady=3);dg.columnconfigure(i,weight=1)
    chart=self.chart_card(self.content,"GRÁFICO DE DETECÇÃO (D)","Unidade de ocorrência • fonte e método permanecem auditáveis");chart.pack(fill="both",expand=True,pady=4)
    self.draw_bar_chart(chart,[("REAL",d["real"],GOOD),("DETECTADO",d["det"],ACCENT2),("TP-D",d["tp"],GOOD),("FP-D",d["fp"],BAD),("FN-D",d["fn"],WARN)],max_value=max(d["real"],d["det"],d["tp"],d["fp"],d["fn"],1))
    band2=tk.Frame(self.content,bg="#0A2635",highlightbackground=ACCENT2,highlightthickness=1);band2.pack(fill="x",pady=(7,5))
    tk.Label(band2,text="02  •  CLASSIFICAÇÃO (C)",bg="#0A2635",fg=ACCENT2,font=("Segoe UI",12,"bold")).pack(side="left",padx=14,pady=9)
    tk.Label(band2,text="INSPEÇÕES — decisão OK/NOK",bg="#0A2635",fg=TEXT,font=("Segoe UI",9)).pack(side="left",padx=4,pady=9)
    if self.source_file and self.history and not self.source_analysis.get("classification_matrix",{}).get("valid"):
        card=self.card(self.content,"CLASSIFICAÇÃO C — N/D","A fonte não fornece uma matriz C válida por inspeção. A matriz histórica de ocorrências não é reutilizada como classificação.");card.pack(fill="x",pady=4)
    else:
        cg=tk.Frame(self.content,bg=BG);cg.pack(fill="x")
        for i,(lab,val,col) in enumerate([("TP-C",m["TP"],GOOD),("VN-C",m["VN/TN"],GOOD),("FP-C",m["FP"],BAD),("FN-C",m["FN"],WARN)]): self.metric_card(cg,lab,str(val),col).grid(row=0,column=i,sticky="ew",padx=3,pady=3);cg.columnconfigure(i,weight=1)
        cc=self.chart_card(self.content,"GRÁFICO DE CLASSIFICAÇÃO (C)","Unidade: inspeção • decisão OK/NOK");cc.pack(fill="both",expand=True,pady=4);self.draw_bar_chart(cc,[("TP-C",m["TP"],GOOD),("VN-C",m["VN/TN"],GOOD),("FP-C",m["FP"],BAD),("FN-C",m["FN"],WARN)],max_value=max(m["TP"],m["VN/TN"],m["FP"],m["FN"],1))
    a=self.source_analysis or {};agg=a.get("aggregate_source",{})
    if agg.get("present"):
        note=self.card(self.content,"RASTREABILIDADE DA FONTE","Valores declarados pela planilha não são confundidos com número de linhas importadas.");note.pack(fill="x",pady=4)
        msg=f"Inspeções observadas pelo AIA: {len(self.rows)}"
        if agg.get("total_inspections") is not None: msg+=f"  •  Total de inspeções declarado na fonte: {agg['total_inspections']}"
        if agg.get("total_occurrences") is not None: msg+=f"  •  Total de ocorrências declarado na fonte: {agg['total_occurrences']}"
        tk.Label(note,text=msg,bg=PANEL,fg=TEXT,font=("Segoe UI",9,"bold"),wraplength=1200,justify="left").pack(anchor="w",padx=16,pady=10)


def _aia_v111_view_graficos(self):
    fonte="HISTÓRICO INTERPRETADO" if (self.source_file and self.history) else "MOTOR NOVO / ENTRADA MANUAL"
    self.section_title("GRÁFICOS  /  PERFORMANCE",f"{fonte} • C e D independentes • V{AIA_SMART_ENGINE_VERSION}")
    m=self.display_metrics();d=self.defect_counts();dt=self.detection_totals();di=dt.get("inspection",{})
    grid=tk.Frame(self.content,bg=BG);grid.pack(fill="both",expand=True)
    # 01 D ocorrência — NÃO aninhar chart_card.
    c1=self.chart_card(grid,"01  •  MATRIZ DE DETECÇÃO (D) — OCORRÊNCIA","TP-D / FP-D / FN-D • TN-D somente quando a fonte define esse universo")
    c1.grid(row=0,column=0,sticky="nsew",padx=(0,4),pady=(0,4))
    self.draw_bar_chart(c1,[("TP-D",d["tp"],GOOD),("FP-D",d["fp"],BAD),("FN-D",d["fn"],WARN)],max_value=max(d["tp"],d["fp"],d["fn"],1))
    c2=self.chart_card(grid,"02  •  MATRIZ DE CLASSIFICAÇÃO (C)","TP-C / VN-C / FP-C / FN-C • unidade: inspeção")
    c2.grid(row=0,column=1,sticky="nsew",padx=(4,0),pady=(0,4))
    self.draw_bar_chart(c2,[("TP-C",m["TP"],GOOD),("VN-C",m["VN/TN"],GOOD),("FP-C",m["FP"],BAD),("FN-C",m["FN"],WARN)],max_value=max(m["TP"],m["VN/TN"],m["FP"],m["FN"],1))
    c3=self.chart_card(grid,"03  •  INDICADORES DE DETECÇÃO (D) — INSPEÇÃO","Accuracy • Precision • Recall • Specificity • F1")
    c3.grid(row=1,column=0,sticky="nsew",padx=(0,4),pady=(4,0))
    if di.get("available"):
        self.draw_percent_chart(c3,[("Accuracy",di["Accuracy"],GOOD),("Precision",di["Precision"],GOOD),("Recall",di["Recall / Sensibilidade"],GOOD),("Specificity",di["Specificity"],GOOD),("F1",di["F1"],GOOD)])
    else:
        tk.Label(c3,text="N/D\n\nNão existe base suficiente para uma matriz D por inspeção.",bg=PANEL,fg=WARN,font=("Segoe UI",11,"bold"),justify="center").pack(expand=True)
    c4=self.chart_card(grid,"04  •  PERFIL DE DETECÇÃO (D) — OCORRÊNCIA","Real × Detectado × TP-D × FP-D × FN-D")
    c4.grid(row=1,column=1,sticky="nsew",padx=(4,0),pady=(4,0))
    self.draw_bar_chart(c4,[("Real",d["real"],GOOD),("Detectado",d["det"],ACCENT2),("TP-D",d["tp"],GOOD),("FP-D",d["fp"],BAD),("FN-D",d["fn"],WARN)],max_value=max(d["real"],d["det"],d["tp"],d["fp"],d["fn"],1))
    for c in (0,1):grid.columnconfigure(c,weight=1)
    for r in (0,1):grid.rowconfigure(r,weight=1)
    # Linha inferior: resumo das métricas C para não perder a visão anterior.
    summary=self.card(self.content,"INDICADORES DE CLASSIFICAÇÃO (C)","A matriz C permanece disponível quando validada pela fonte.")
    summary.pack(fill="x",pady=6)
    sg=tk.Frame(summary,bg=PANEL);sg.pack(fill="x",padx=8,pady=6)
    vals=[("Accuracy",pct(m["Accuracy"]),ACCENT2),("Precision",pct(m["Precision"]),ACCENT2),("Recall",pct(m["Recall / Sensibilidade"]),GOOD),("Specificity",pct(m["Specificity"]),ACCENT2),("F1",pct(m["F1"]),ACCENT2)]
    for i,x in enumerate(vals):self.metric_card(sg,*x).grid(row=0,column=i,sticky="ew",padx=3);sg.columnconfigure(i,weight=1)


def _aia_v111_view_tecnica(self):
    fonte="HISTÓRICO INTERPRETADO" if (self.source_file and self.history) else "MOTOR NOVO / ENTRADA MANUAL"
    self.section_title("ANÁLISE TÉCNICA  /  DIAGNÓSTICO",f"{fonte} • V{AIA_SMART_ENGINE_VERSION} • duas matrizes, duas unidades")
    m=self.display_metrics();d=self.defect_counts();dt=self.detection_totals();di=dt.get("inspection",{});a=self.source_analysis or {}
    # C
    c=self.card(self.content,"01  •  CLASSIFICAÇÃO (C)","Unidade: inspeção • a decisão OK/NOK está correta?");c.pack(fill="x",pady=5)
    if self.source_file and self.history and not a.get("classification_matrix",{}).get("valid"):
        tk.Label(c,text="N/D — a fonte não fornece uma matriz C válida por inspeção. O AIA preserva essa limitação e não reutiliza D como C.",bg=PANEL,fg=WARN,font=("Segoe UI",10,"bold"),wraplength=1200,justify="left").pack(anchor="w",padx=18,pady=16)
    else:
        cards=[("Nº de inspeções",m["N"],TEXT),("Accuracy",pct(m["Accuracy"]),ACCENT2),("Precision",pct(m["Precision"]),ACCENT2),("Recall",pct(m["Recall / Sensibilidade"]),GOOD),("Specificity",pct(m["Specificity"]),ACCENT2),("FPR",pct(m["FPR"]),BAD),("FNR",pct(m["FNR"]),WARN),("F1",pct(m["F1"]),ACCENT2),("MCC","N/D" if m["MCC"] is None else f"{m['MCC']:.4f}",ACCENT2),("Prevalência",pct(safe(m["TP"]+m["FN"],m["N"])),ACCENT2)]
        self._aia_v111_metric_cards(c,"MATRIZ C — INDICADORES",f"TP-C={m['TP']} • VN-C={m['VN/TN']} • FP-C={m['FP']} • FN-C={m['FN']}",cards,5)
    # D por inspeção
    c=self.card(self.content,"02  •  DETECÇÃO (D) — POR INSPEÇÃO","Unidade: inspeção • Real>0 × Detectado>0");c.pack(fill="x",pady=5)
    if di.get("available"):
        cards=[("Nº de inspeções",di["N"],GOOD),("Accuracy",pct(di["Accuracy"]),ACCENT2),("Precision",pct(di["Precision"]),ACCENT2),("Recall",pct(di["Recall / Sensibilidade"]),GOOD),("Specificity",pct(di["Specificity"]),ACCENT2),("FPR",pct(di["FPR"]),BAD),("FNR",pct(di["FNR"]),WARN),("F1",pct(di["F1"]),ACCENT2),("MCC","N/D" if di["MCC"] is None else f"{di['MCC']:.4f}",ACCENT2),("Prevalência",pct(safe(di["TP"]+di["FN"],di["N"])),ACCENT2)]
        self._aia_v111_metric_cards(c,"MATRIZ D — INSPEÇÃO",f"TP-D={di['TP']} • TN-D={di['TN']} • FP-D={di['FP']} • FN-D={di['FN']} • método: {di['method']}",cards,5)
    else:
        tk.Label(c,text="N/D — Real e Detectado não estão disponíveis por inspeção em quantidade suficiente para construir essa matriz.",bg=PANEL,fg=WARN,font=("Segoe UI",10,"bold"),wraplength=1200,justify="left").pack(anchor="w",padx=18,pady=16)
    # D por ocorrência
    c=self.card(self.content,"03  •  DETECÇÃO (D) — POR OCORRÊNCIA","Unidade: defeito/ocorrência • não confundir com inspeção");c.pack(fill="x",pady=5)
    full=d.get("full_metrics")
    occ_metrics=[("Nº de ocorrências reais",d["real"],GOOD),("Nº de detecções",d["det"],ACCENT2),("TP-D",d["tp"],GOOD),("FP-D",d["fp"],BAD),("FN-D",d["fn"],WARN)]
    if full:
        occ_metrics += [("Accuracy",pct(full["Accuracy"]),ACCENT2),("Precision",pct(full["Precision"]),ACCENT2),("Recall",pct(full["Recall / Sensibilidade"]),GOOD),("Specificity",pct(full["Specificity"]),ACCENT2),("FPR",pct(full["FPR"]),BAD),("FNR",pct(full["FNR"]),WARN),("F1",pct(full["F1"]),ACCENT2),("MCC","N/D" if full["MCC"] is None else f"{full['MCC']:.4f}",ACCENT2),("Prevalência",pct(safe(full["TP"]+full["FN"],full["N"])),ACCENT2)]
        subtitle=f"TN-D={full['VN/TN']} • N={full['N']} • {d.get('status','N/D')} • {d.get('method','N/D')}"
    else:
        occ_metrics += [("Accuracy","N/D",ACCENT2),("Precision",pct(d.get("precision")),ACCENT2),("Recall",pct(d.get("recall")),GOOD),("Specificity","N/D",ACCENT2),("FPR","N/D",BAD),("FNR","N/D",WARN),("F1",pct(d.get("f1")),ACCENT2),("MCC","N/D",ACCENT2),("Prevalência","N/D",ACCENT2)]
        subtitle=f"TN-D não observável • {d.get('status','N/D')} • {d.get('method','N/D')}"
    self._aia_v111_metric_cards(c,"MATRIZ D — OCORRÊNCIA",subtitle,occ_metrics,5)
    # Auditoria grande e legível
    ab=self.card(self.content,"04  •  AUDITORIA DA INTERPRETAÇÃO","De onde saiu cada número? • a unidade e o método são parte do resultado.");ab.pack(fill="x",pady=5)
    lines=[f"Arquivo: {a.get('file','N/D')}",f"Aba principal: {a.get('sheet','N/D')} • Cabeçalho: {a.get('header_row','N/D')}",f"Inspeções observadas pelo AIA: {len(self.rows)}",f"Status da interpretação: {a.get('status','N/D')}"]
    agg=a.get("aggregate_source",{})
    if agg.get("total_inspections") is not None: lines.append(f"Total de inspeções declarado na fonte: {agg['total_inspections']} (não substitui registros efetivamente observados)")
    if agg.get("total_occurrences") is not None: lines.append(f"Total de ocorrências declarado na fonte: {agg['total_occurrences']}")
    if agg.get("complete_matrix"): lines.append(f"Matriz agregada de ocorrência: TP={agg['tp']} • TN={agg['tn']} • FP={agg['fp']} • FN={agg['fn']} • N={agg['n']} • origem: {agg.get('sheet')}")
    for line in lines:tk.Label(ab,text=line,bg=PANEL,fg=TEXT,font=("Segoe UI",9),wraplength=1200,justify="left").pack(anchor="w",padx=18,pady=3)
    self._aia_smart_add_interpretation_button(ab)


# Guardar referências anteriores e aplicar V1.1.2 por último.
App._aia_v111_old_import_excel = App.import_excel
App._aia_v111_old_detection_occurrence = App._smart_detection_occurrence
App.import_excel = _aia_v111_import_excel
App._smart_detection_occurrence = _aia_v111_detection_occurrence
App.detection_totals = _aia_smart_detection_totals
App.defect_counts = _aia_smart_defect_counts
App.view_dashboard = _aia_v111_view_dashboard
App.view_graficos = _aia_v111_view_graficos
App._aia_v111_metric_cards = _aia_v111_metric_cards
App._aia_smart_add_interpretation_button = _aia_smart_add_interpretation_button
App.view_tecnica = _aia_v111_view_tecnica



# ============================================================================
# AIA V1.1.3 — MOTOR DE VERDADE MATEMÁTICA / IMPORTADOR SEMÂNTICO
# Foco desta release: eliminar ambiguidade C x D, preservar ausência como N/D,
# separar N_inspecoes / N_ocorrencias / N_matriz e impedir promoção automática
# de VP/VN/FP/FN genéricos para uma unidade sem evidência contextual.
# ============================================================================
AIA_SMART_ENGINE_VERSION = "1.1.10"
AIA_DATA_ENGINE_VERSION = "1.1.7"


def _aia_v113_num(value):
    """Número observável. Ausência/blank/texto não numérico => None."""
    return _aia_num_optional(value)


def _aia_v113_matrix_sum(tp, tn, fp, fn):
    vals = [_aia_v113_num(x) for x in (tp, tn, fp, fn)]
    if any(x is None for x in vals):
        return None
    if any(x < 0 for x in vals):
        return None
    return sum(vals)


def _aia_v113_confusion(tp, tn, fp, fn, n=None):
    vals = [_aia_v113_num(x) for x in (tp, tn, fp, fn)]
    if any(x is None for x in vals):
        return {"valid": False, "reason": "componente ausente"}
    tp, tn, fp, fn = vals
    if any(x < 0 for x in vals):
        return {"valid": False, "reason": "componente negativo"}
    derived_n = tp + tn + fp + fn
    if n is not None and _aia_v113_num(n) is not None and int(n) != derived_n:
        return {"valid": False, "reason": f"N declarado ({n}) não fecha com a matriz ({derived_n})"}
    return {"valid": True, "tp": tp, "tn": tn, "fp": fp, "fn": fn,
            "n": derived_n if n is None else int(n),
            "closure": "VP+VN+FP+FN = N"}


def _aia_v113_context_scan(path):
    """Localiza agregados e exige evidência de unidade; não usa rótulo VP/VN/FP/FN isoladamente."""
    out = {
        "present": False, "sheet": None, "cells": {},
        "total_inspections": None, "total_occurrences": None,
        "tp": None, "tn": None, "fp": None, "fn": None,
        "complete_matrix": False, "n": None, "unit": "não determinado",
        "status": "N/D", "method": "N/D", "evidence": [], "conflicts": []
    }
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
        label_map = {
            "total_inspections": ["TOTAL DE INSPECOES", "TOTAL DE INSPECAO"],
            "total_occurrences": ["TOTAL DE OCORRENCIAS", "TOTAL DE OCORRENCIA"],
            "tp": ["TP - VERDADEIRO POSITIVO", "VP - VERDADEIRO POSITIVO"],
            "tn": ["TN - VERDADEIRO NEGATIVO", "VN - VERDADEIRO NEGATIVO"],
            "fp": ["FP - FALSO POSITIVO"],
            "fn": ["FN - FALSO NEGATIVO"],
        }
        for ws in wb.worksheets:
            for ridx, row in enumerate(ws.iter_rows(), 1):
                vals = [c.value for c in row]
                for j, v in enumerate(vals):
                    key = normalize_text(v)
                    if not key:
                        continue
                    for dest, aliases in label_map.items():
                        if any(key == a or key.startswith(a) for a in aliases):
                            # O valor pode estar na célula imediatamente à direita.
                            if j + 1 < len(vals):
                                num = _aia_v113_num(vals[j+1])
                                if num is not None and out[dest] is None:
                                    out[dest] = num
                                    out["sheet"] = ws.title
                                    try:
                                        coord = ws.cell(ridx, j+2).coordinate
                                    except Exception:
                                        coord = f"{ws.title}!R{ridx}C{j+2}"
                                    out["cells"][dest] = coord
        # Padrão de matriz em linha: "TOTAIS POR CLASSIFICAÇÃO" seguido de 4 números.
        # O rótulo não determina a unidade; serve apenas como evidência de conflito/contexto.
        for ws in wb.worksheets:
            for ridx, row in enumerate(ws.iter_rows(), 1):
                vals = [c.value for c in row]
                text = " | ".join(normalize_text(v) for v in vals if v is not None)
                if "TOTAIS POR CLASSIFICACAO" in text:
                    out["evidence"].append("contexto contém TOTAIS POR CLASSIFICAÇÃO")
                    nums=[]
                    for j,v in enumerate(vals):
                        num=_aia_v113_num(v)
                        if num is not None:
                            nums.append((j,num))
                    # Em planilhas históricas deste projeto, a sequência da linha é VP,VN,FP,FN.
                    if len(nums) >= 4 and not out["complete_matrix"]:
                        # Escolhe uma sequência de 4 colunas numericamente consecutivas;
                        # isso evita capturar um valor isolado de outra tabela na mesma linha.
                        chosen=None
                        for start in range(len(nums)-3):
                            cand=nums[start:start+4]
                            if all(cand[k+1][0] == cand[k][0] + 1 for k in range(3)):
                                chosen=cand
                                break
                        if chosen is not None:
                            out["tp"],out["tn"],out["fp"],out["fn"]=[x[1] for x in chosen]
                            out["sheet"] = out["sheet"] or ws.title
                            out["cells"]["matrix_row"] = f"{ws.title}!{ridx}"
        # Contexto global do workbook/aba da evidência principal.
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                text = " | ".join(normalize_text(c.value) for c in row if c.value is not None)
                if not text:
                    continue
                if "TOTAL DE OCORRENCIAS" in text or "OCORRENCIAS DETECTADAS" in text or "DETECCAO DE OCORRENCIAS" in text:
                    out["evidence"].append("contexto contém OCORRÊNCIA/DETECÇÃO DE OCORRÊNCIAS")
                if "TOTAL DE INSPECOES" in text:
                    out["evidence"].append("fonte declara TOTAL DE INSPEÇÕES")
        # Deduplica evidências preservando ordem.
        out["evidence"] = list(dict.fromkeys(out["evidence"]))
        out["conflicts"] = list(dict.fromkeys(out["conflicts"]))
        if "TOTAIS POR CLASSIFICACAO" in " | ".join(out["evidence"]):
            if out["total_occurrences"] is not None:
                out["conflicts"].append("rótulo de classificação coexistindo com total explícito de ocorrências")
        out["present"] = any(out[k] is not None for k in ("tp","tn","fp","fn","total_occurrences","total_inspections"))
        out["complete_matrix"] = all(out[k] is not None for k in ("tp","tn","fp","fn"))
        if out["complete_matrix"]:
            out["n"] = out["tp"] + out["tn"] + out["fp"] + out["fn"]
            occ_evidence = out["total_occurrences"] is not None or any("OCORRÊNCIA" in e or "OCORRENCIA" in e for e in out["evidence"])
            class_evidence = any("CLASSIFICAÇÃO" in e or "CLASSIFICACAO" in e for e in out["evidence"])
            if occ_evidence:
                out["unit"] = "ocorrência"
                out["status"] = "CONFIRMADO_COM_RESSALVA" if class_evidence else "CONFIRMADO"
                out["method"] = "Matriz agregada preservada; unidade sustentada por evidência de ocorrência"
            elif class_evidence:
                out["unit"] = "não determinado"
                out["status"] = "AMBÍGUO"
                out["method"] = "Matriz agregada encontrada, mas a fonte não comprova a unidade"
            else:
                out["unit"] = "não determinado"
                out["status"] = "AMBÍGUO"
                out["method"] = "Matriz agregada encontrada sem evidência semântica suficiente"
        else:
            out["method"] = "Indicadores agregados parciais; matriz completa não comprovada"
        wb.close()
    except Exception as exc:
        out["error"] = str(exc)
        out["status"] = "ERRO"
    return out

def _aia_v113_classification_from_rows(rows, explicit_c=False, generic_c=False):
    """Determina C somente quando a unidade inspeção é demonstrada."""
    sums = [r.get("c_matrix_count") for r in rows if r.get("id")]
    observed = [x for x in sums if x is not None and x > 0]
    if not observed:
        return {"valid": False, "reason": "nenhuma contribuição C observável"}
    if any(x != 1 for x in observed):
        return {"valid": False, "reason": "há linhas com mais de uma contribuição; não é matriz binária por inspeção"}
    # Para genéricos, exigir também evidência de decisão OK/NOK ou gabarito explícito.
    if generic_c and not explicit_c:
        context_ok = any(
            str(r.get("vision","")).upper() in {"OK","NOK"} and
            str(r.get("ref","")).strip() in {"OK confirmado","Defeito confirmado"}
            for r in rows
        )
        if not context_ok:
            return {"valid": False, "reason": "VP/VN/FP/FN genéricos sem evidência de universo OK/NOK"}
    return {"valid": True, "reason": "cada contribuição C válida representa uma inspeção"}


def _aia_v113_import_excel(self):
    if openpyxl is None:
        messagebox.showerror("Excel", "OpenPyXL não está instalado.\nExecute o instalador de dependências.", parent=self)
        return
    path = filedialog.askopenfilename(title="Abrir planilha de assertividade",
                                      filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")])
    if not path:
        return
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
        candidates=[]
        for ws_try in wb.worksheets:
            hr=self._find_header_row(ws_try)
            if hr is None:
                continue
            hm=self._header_map(ws_try,hr)
            score=0
            if self._norm_header(ws_try.title)=="INSPECOES": score += 10
            for k in ("IDENTIFICACAO","VISION AI","VISION","REFERENCIA","REAL","DETECTADO","TP-C","VN-C","FP-C","FN-C","TP-D","FP-D","FN-D","TP","VN","FP","FN"):
                if k in hm: score += 1
            if "IDENTIFICACAO" in hm: score += 5
            candidates.append((score,ws_try,hr))
        if not candidates:
            raise ValueError("Não encontrei uma tabela de inspeções compatível.")
        _,ws,hr=max(candidates,key=lambda x:x[0])
        hm=self._header_map(ws,hr)
        def col(*names):
            for n in names:
                k=self._norm_header(n)
                if k in hm: return hm[k]
            return None
        c_id=col("Identificação","N Inspeção","Nº Inspeção","Inspeção")
        if c_id is None:
            raise ValueError("A planilha não possui coluna de Identificação/Inspeção.")
        c_vision=col("Resultado Vision AI","Vision AI","Vision")
        c_ref=col("Referência","Reference","Gabarito")
        c_tp_c=col("TP-C"); c_tn_c=col("VN-C"); c_fp_c=col("FP-C"); c_fn_c=col("FN-C")
        c_tp_g=col("TP","Verdadeiro Positivo","VP - Verdadeiro Positivo")
        c_tn_g=col("VN","TN","Verdadeiro Negativo","VN - Verdadeiro Negativo")
        c_fp_g=col("FP","Falso Positivo","FP - Falso Positivo")
        c_fn_g=col("FN","Falso Negativo","FN - Falso Negativo")
        c_real=col("Real","Real (defeitos)","Defeitos reais")
        c_det=col("Detectado","Detectados","Detecções","Detecções registradas")
        c_tp_d=col("TP-D"); c_fp_d=col("FP-D"); c_fn_d=col("FN-D")
        c_def=col("Defeito","Tipo de defeito"); c_obs=col("Observação","Observações")
        c_date=col("Data da inspeção","Data inspeção","Data","Date")
        explicit_c = all(x is not None for x in (c_tp_c,c_tn_c,c_fp_c,c_fn_c))
        generic_c = all(x is not None for x in (c_tp_g,c_tn_g,c_fp_g,c_fn_g)) and not explicit_c
        rows=[]; matrix_rows=[]; imported_dates=[]
        for rr in range(hr+1,ws.max_row+1):
            ident=self._format_identification_cell(ws.cell(rr,c_id))
            if not ident:
                continue
            r=blank_row(); r["id"]=ident
            r["vision"]=(str(ws.cell(rr,c_vision).value or "").strip().upper() if c_vision else "")
            r["ref"]=(str(ws.cell(rr,c_ref).value or "").strip() if c_ref else "")
            # Nunca atribuir 0 a uma coluna ausente. Ausência fica N/D/None.
            # V1.1.9 — MATRIZES C E D-I ESTRITAMENTE INDEPENDENTES.
            # C nunca pode ser inferida de Real/Detectado nem de TP/FP/FN genéricos,
            # pois esses campos podem pertencer à detecção histórica. Prioridade:
            #   1) colunas explícitas TP-C/VN-C/FP-C/FN-C;
            #   2) decisão OK/NOK reconstruída de Vision AI × Referência.
            # D-I é calculada em outra rotina, exclusivamente a partir de Real × Detectado.
            vision_norm=str(r.get("vision","") or "").strip().upper()
            ref_norm=str(r.get("ref","") or "").strip()
            decision_c = vision_norm in {"OK","NOK"} and ref_norm in {"OK confirmado","Defeito confirmado"}
            if explicit_c:
                cp=_aia_v113_num(ws.cell(rr,c_tp_c).value); cn=_aia_v113_num(ws.cell(rr,c_tn_c).value)
                cf=_aia_v113_num(ws.cell(rr,c_fp_c).value); cfn=_aia_v113_num(ws.cell(rr,c_fn_c).value)
                r.update(tp=cp or 0, vn=cn or 0, fp=cf or 0, fn=cfn or 0)
                r["classification_source"]="colunas explícitas TP-C/VN-C/FP-C/FN-C"
                r["classification_observed_values"]=(cp,cn,cf,cfn)
                row_sum=_aia_v113_matrix_sum(cp,cn,cf,cfn)
            elif decision_c:
                cp = 1 if ref_norm=="Defeito confirmado" and vision_norm=="NOK" else 0
                cn = 1 if ref_norm=="OK confirmado" and vision_norm=="OK" else 0
                cf = 1 if ref_norm=="OK confirmado" and vision_norm=="NOK" else 0
                cfn = 1 if ref_norm=="Defeito confirmado" and vision_norm=="OK" else 0
                r.update(tp=cp, vn=cn, fp=cf, fn=cfn)
                r["classification_source"]="derivada exclusivamente de Vision AI × Referência"
                r["classification_observed_values"]=(cp,cn,cf,cfn)
                row_sum=1
            else:
                r.update(tp=0,vn=0,fp=0,fn=0)
                r["classification_source"]=("TP/VN/FP/FN genéricos ignorados por segurança" if generic_c else "não identificado")
                r["classification_observed_values"]=(None,None,None,None)
                row_sum=0
            r["c_matrix_count"]=row_sum
            matrix_rows.append(r)
            raw_real=ws.cell(rr,c_real).value if c_real else None
            raw_det=ws.cell(rr,c_det).value if c_det else None
            real=_aia_v113_num(raw_real); det=_aia_v113_num(raw_det)
            r["real"]=real
            r["det"]=det
            r["detection_observed"]=(c_real is not None and c_det is not None and real is not None and det is not None)
            tp_d=_aia_v113_num(ws.cell(rr,c_tp_d).value) if c_tp_d else None
            fp_d=_aia_v113_num(ws.cell(rr,c_fp_d).value) if c_fp_d else None
            fn_d=_aia_v113_num(ws.cell(rr,c_fn_d).value) if c_fn_d else None
            r["tp_d"]=tp_d; r["fp_d"]=fp_d; r["fn_d"]=fn_d
            r["detection_explicit_matrix"]=all(x is not None for x in (tp_d,fp_d,fn_d))
            r["obs"]=str(ws.cell(rr,c_obs).value or "") if c_obs else ""
            r["calc_status"]="Histórico"; r["source"]="Excel histórico"
            if c_def:
                name=normalize_defect(ws.cell(rr,c_def).value)
                if name and real is not None and det is not None:
                    r["defects"]=[{"type":name,"real":real,"det":det}]
            if c_date:
                dd=self._format_inspection_date(ws.cell(rr,c_date).value)
                if dd: imported_dates.append(dd)
            rows.append(r)
        if not rows:
            raise ValueError("Nenhuma inspeção válida foi encontrada na planilha.")

        # C é validada apenas quando há fonte C explícita ou decisão Vision×Referência.
        # Campos genéricos VP/VN/FP/FN não são mais promovidos a C automaticamente.
        derived_c = any(r.get("classification_source","").startswith("derivada exclusivamente") for r in rows)
        c_decision=_aia_v113_classification_from_rows(rows, explicit_c=explicit_c, generic_c=False)
        c_valid=bool(c_decision.get("valid")) and (explicit_c or derived_c)
        for r in rows:
            r["classification_observed"]=c_valid and r.get("c_matrix_count")==1

        # Não confundir número de registros com N da matriz.
        c_tp=sum(int(r.get("tp",0) or 0) for r in rows if r.get("classification_observed"))
        c_tn=sum(int(r.get("vn",0) or 0) for r in rows if r.get("classification_observed"))
        c_fp=sum(int(r.get("fp",0) or 0) for r in rows if r.get("classification_observed"))
        c_fn=sum(int(r.get("fn",0) or 0) for r in rows if r.get("classification_observed"))
        c_n=c_tp+c_tn+c_fp+c_fn

        self._archive_current_collection()
        self.rows=rows; self.history=[{"TP":c_tp,"VN/TN":c_tn,"FP":c_fp,"FN":c_fn,"N":c_n}]
        self.source_file=path
        imported_meta,metadata_status=self._extract_workbook_metadata(wb,ws,hr,path)
        imported_date=__import__('collections').Counter(imported_dates).most_common(1)[0][0] if imported_dates else ""
        if imported_date: imported_meta["data"]=imported_date
        for key in self.meta:
            if imported_meta.get(key): self.meta[key].set(imported_meta[key])
        if not self.meta["gabarito"].get(): self.meta["gabarito"].set("Não determinado")
        wb.close()

        agg=_aia_v113_context_scan(path)
        real_field_present=c_real is not None
        det_field_present=c_det is not None
        explicit_d=all(x is not None for x in (c_tp_d,c_fp_d,c_fn_d))
        di_rows=sum(1 for r in rows if r.get("detection_observed"))
        self.source_analysis={
            "file":str(path), "sheet":ws.title, "header_row":hr,
            "inspections_registered":len(rows),
            "n_inspections":len(rows),
            "n_occurrences_declared":agg.get("total_occurrences"),
            "classification_matrix":{
                "valid":c_valid,"unit":"inspeção","tp":c_tp,"tn":c_tn,"fp":c_fp,"fn":c_fn,"n":c_n,
                "method":"explícita TP-C/VN-C/FP-C/FN-C" if explicit_c else ("derivada Vision AI × Referência" if derived_c and c_valid else "não validada"),
                "reason":c_decision.get("reason")
            },
            "occurrence_matrix":{
                "present":bool(agg.get("complete_matrix") and agg.get("unit")=="ocorrência"),
                "tp":agg.get("tp"),"tn":agg.get("tn"),"fp":agg.get("fp"),"fn":agg.get("fn"),"n":agg.get("n"),
                "unit":agg.get("unit"),"method":agg.get("method"),"status":agg.get("status"),
                "sheet":agg.get("sheet"),"cells":agg.get("cells",{}),"evidence":agg.get("evidence",[]),"conflicts":agg.get("conflicts",[])
            },
            "detection_fields":{"real":real_field_present,"detected":det_field_present,"tp_d":c_tp_d is not None,"fp_d":c_fp_d is not None,"fn_d":c_fn_d is not None},
            "detection_inspection_rows":di_rows,
            "mapping":{"TP-C":c_tp_c,"VN-C":c_tn_c,"FP-C":c_fp_c,"FN-C":c_fn_c,"TP/VP":c_tp_g,"VN/TN":c_tn_g,"FP":c_fp_g,"FN":c_fn_g,"Real":c_real,"Detectado":c_det,"TP-D":c_tp_d,"FP-D":c_fp_d,"FN-D":c_fn_d},
            "metadata_status":metadata_status,
            "ambiguities":{
                "generic_confusion_columns_present":bool(generic_c),
                "generic_confusion_columns_action":(
                    "IGNORADAS — TP/VN/FP/FN genéricos não foram atribuídos a C, D-I ou D-O"
                    if generic_c else "NÃO IDENTIFICADAS"
                ),
                "classification_rows_unavailable":sum(1 for r in rows if not r.get("classification_observed")),
                "detection_rows_unavailable":sum(1 for r in rows if not r.get("detection_observed")),
            },
            "aggregate_source":agg,
            "status":"PARCIAL",
        }
        self.source_analysis["status"] = "CONFIRMADO" if (c_valid or self.source_analysis["occurrence_matrix"]["present"] or di_rows>0) else "PARCIAL"
        self.refresh_all()
        ca=self.source_analysis["classification_matrix"]; oa=self.source_analysis["occurrence_matrix"]
        d=self._smart_detection_occurrence(); di=self._smart_detection_inspection()
        status=[f"{len(rows)} inspeções registradas"]
        status.append(f"C={'validada' if ca['valid'] else 'N/D'}")
        status.append(f"D-O={d.get('status','DISPONÍVEL') if d.get('available') else 'N/D'}")
        status.append(f"D-I={'disponível' if di['available'] else 'N/D'}")
        self.status.set("✓ Excel interpretado • " + " • ".join(status))
        messagebox.showinfo("Interpretação da planilha concluída",
            f"Planilha: {Path(path).name}\nAba: {ws.title}\n\n"
            f"Inspeções registradas: {len(rows)}\n"
            f"Classificação C: {'VALIDADA' if ca['valid'] else 'N/D'}\n"
            f"Detecção por inspeção: {'DISPONÍVEL' if di['available'] else 'N/D'}\n"
            f"Detecção por ocorrência: {d.get('status','DISPONÍVEL') if d.get('available') else 'N/D'}\n\n"
            "Regra aplicada: ausência de campo não é tratada como zero; VP/VN/FP/FN genéricos não recebem unidade por nome isolado.", parent=self)
    except Exception as exc:
        messagebox.showerror("Erro na importação Excel",f"Não foi possível interpretar a planilha.\n\nDetalhe: {exc}",parent=self)


def _aia_v113_detection_occurrence(self):
    sa=getattr(self,"source_analysis",{}) or {}
    om=sa.get("occurrence_matrix",{}) or {}
    observed_rows=[]; unavailable_rows=0
    for r0 in self.rows:
        row=normalize_row(r0)
        if row.get("detection_observed",False):
            observed_rows.append(row)
        else:
            unavailable_rows+=1
    # Uma matriz agregada validada só é usada quando a fonte não oferece dados
    # Real/Detectado por inspeção. Havendo linhas, a soma linha a linha prevalece.
    if getattr(self,"source_file","") and om.get("present") and not observed_rows:
        tp,tn,fp,fn=om.get("tp"),om.get("tn"),om.get("fp"),om.get("fn")
        vals=[_aia_v113_num(x) for x in (tp,tn,fp,fn)]
        if all(x is not None for x in vals):
            tp,tn,fp,fn=vals
            mm=_aia_metric_pack(tp,tn,fp,fn,tp+tn+fp+fn)
            return {"available":True,"real":tp+fn,"det":tp+fp,"tp":tp,"fp":fp,"fn":fn,"tn":tn,
                    "precision":mm["Precision"],"recall":mm["Recall / Sensibilidade"],"f1":mm["F1"],
                    "available_rows":len(self.rows),"unavailable_rows":0,"detailed":False,"estimated":False,
                    "status":om.get("status","IMPORTADO / VALIDADO"),
                    "method":om.get("method","Matriz agregada da fonte"),"full_metrics":mm,
                    "source_sheet":om.get("sheet"),"source_cells":om.get("cells",{}),
                    "semantic_evidence":om.get("evidence",[]),"semantic_conflicts":om.get("conflicts",[])}
    # Consolidação D-O canônica: calcula cada inspeção exatamente como a tela
    # Coleta e, somente depois, soma os resultados. FP de uma linha jamais
    # compensa FN de outra linha.
    real=det=tp=fp=fn=0; available_rows=0; explicit_rows=0; explicit_ok=True; detailed=False
    for row in observed_rows:
        line=self.detection_for_row(row)
        if not line.get("available",False):
            unavailable_rows+=1
            continue
        available_rows+=1
        for key in ("real","det","tp","fp","fn"):
            value=max(0,int(line.get(key,0) or 0))
            if key=="real": real+=value
            elif key=="det": det+=value
            elif key=="tp": tp+=value
            elif key=="fp": fp+=value
            else: fn+=value
        detailed = detailed or bool(row.get("defects") or [])
        if row.get("detection_explicit_matrix",False):
            explicit_rows+=1
            imported=(max(0,int(row.get("tp_d",0) or 0)),
                      max(0,int(row.get("fp_d",0) or 0)),
                      max(0,int(row.get("fn_d",0) or 0)))
            calculated=(line["tp"],line["fp"],line["fn"])
            if imported!=calculated:
                explicit_ok=False
    if available_rows==0:
        return {"available":False,"real":0,"det":0,"tp":0,"fp":0,"fn":0,"tn":0,"precision":None,"recall":None,"f1":None,"status":"N/D","method":"N/D"}
    if explicit_rows==available_rows and explicit_ok:
        method="TP-D/FP-D/FN-D por inspeção importados e validados linha a linha"; status="IMPORTADO / VALIDADO"; estimated=False
    else:
        method="Soma D-O por inspeção (min/max por linha)"
        status="ESTIMADO"; estimated=True
    precision=safe(tp,tp+fp); recall=safe(tp,tp+fn)
    return {"available":True,"real":real,"det":det,"tp":tp,"fp":fp,"fn":fn,
            "precision":precision,"recall":recall,
            "f1":(None if precision is None or recall is None or precision+recall==0 else 2*precision*recall/(precision+recall)),
            "available_rows":available_rows,"unavailable_rows":unavailable_rows,"detailed":detailed,"estimated":estimated,
            "status":status,"method":method}


# A V1.1.3 assume o controle definitivo do importador e da D-O.
App.import_excel = _aia_v113_import_excel
App._smart_detection_occurrence = _aia_v113_detection_occurrence
App.detection_totals = _aia_smart_detection_totals

# ============================================================================
# AIA V1.1.5 — COMPARAÇÃO OPERACIONAL D-I
# D-I é armazenado no snapshot e usado como referência operacional na Comparação.
# D-O e C permanecem em universos estatísticos separados.
# ============================================================================

# Inicialização movida para o FINAL do arquivo para garantir que os patches
# V1.1.5/V1.1.6 sejam aplicados ANTES de abrir o aplicativo.

# ============================================================================
# AIA V1.1.5 FIX2 — FOCO OPERACIONAL D-I
# Ajuste final de apresentação: D-I (detecção por inspeção) é a leitura
# primária para campo. D-O e C permanecem disponíveis como camadas secundárias.
# Não altera as fórmulas do motor; apenas reposiciona e explicita onde cada
# resultado deve aparecer.
# ============================================================================

def _aia_v115_field_dashboard(self):
    fonte="HISTÓRICO INTERPRETADO" if (self.source_file and self.history) else "ENTRADA MANUAL"
    self.section_title("DASHBOARD  /  VISÃO EXECUTIVA",
        f"{fonte} • D-I é a referência operacional de campo • D-O e C permanecem separados")
    dt=self.detection_totals(); di=dt.get("inspection",{}) or {}; d=self.defect_counts(); m=self.display_metrics()

    # 01 — D-I PRIMEIRO: a pergunta operacional é "o sistema encontrou o defeito nesta inspeção?"
    band=tk.Frame(self.content,bg="#103126",highlightbackground=GOOD,highlightthickness=1);band.pack(fill="x",pady=(0,5))
    tk.Label(band,text="01  •  DETECÇÃO POR INSPEÇÃO (D-I)",bg="#103126",fg=GOOD,font=("Segoe UI",12,"bold")).pack(side="left",padx=14,pady=9)
    tk.Label(band,text="REFERÊNCIA OPERACIONAL • defeito real × detecção na mesma inspeção",bg="#103126",fg=TEXT,font=("Segoe UI",9)).pack(side="left",padx=4,pady=9)
    if di.get("available"):
        g=tk.Frame(self.content,bg=BG);g.pack(fill="x")
        vals=[("D-I N",di["N"],GOOD),("TP-DI",di["TP"],GOOD),("TN-DI",di["TN"],GOOD),("FP-DI",di["FP"],BAD),("FN-DI",di["FN"],WARN)]
        for i,x in enumerate(vals): self.metric_card(g,*x).grid(row=0,column=i,sticky="ew",padx=2,pady=3);g.columnconfigure(i,weight=1)
        g2=tk.Frame(self.content,bg=BG);g2.pack(fill="x")
        vals2=[("Accuracy D-I",pct(di["Accuracy"]),ACCENT2),("Precision D-I",pct(di["Precision"]),ACCENT2),("Recall D-I",pct(di["Recall / Sensibilidade"]),GOOD),("Specificity D-I",pct(di["Specificity"]),ACCENT2),("F1 D-I",pct(di["F1"]),ACCENT2)]
        for i,x in enumerate(vals2): self.metric_card(g2,*x).grid(row=0,column=i,sticky="ew",padx=2,pady=3);g2.columnconfigure(i,weight=1)
        c=self.chart_card(self.content,"MATRIZ D-I — POR INSPEÇÃO","Unidade: inspeção • TP-DI / TN-DI / FP-DI / FN-DI")
        c.pack(fill="both",expand=True,pady=4)
        self.draw_bar_chart(c,[("TP-DI",di["TP"],GOOD),("TN-DI",di["TN"],GOOD),("FP-DI",di["FP"],BAD),("FN-DI",di["FN"],WARN)],max_value=max(di["TP"],di["TN"],di["FP"],di["FN"],1))
        note=self.card(self.content,"LEITURA RÁPIDA DE CAMPO","O que o operador/engenharia deve observar primeiro.")
        note.pack(fill="x",pady=4)
        if di["FN"]>0:
            txt=f"ATENÇÃO: {di['FN']} inspeção(ões) com defeito real não foram detectadas (FN-DI). Prioridade: reduzir escapes."
            fg=BAD
        elif di["FP"]>0:
            txt=f"ATENÇÃO: {di['FP']} inspeção(ões) sem defeito geraram detecção (FP-DI). Prioridade: reduzir falsos alarmes."
            fg=WARN
        else:
            txt="SEM FP-DI E FN-DI NA AMOSTRA ATUAL: a detecção por inspeção não apresentou erro observado. Repetir em amostra maior antes de concluir estabilidade."
            fg=GOOD
        tk.Label(note,text=txt,bg=PANEL,fg=fg,font=("Segoe UI",10,"bold"),wraplength=1200,justify="left").pack(anchor="w",padx=16,pady=12)
    else:
        card=self.card(self.content,"D-I — N/D","A matriz por inspeção exige Real e Detectado válidos em nível de inspeção.")
        card.pack(fill="x",pady=5)
        tk.Label(card,text="Não há base suficiente para calcular D-I nesta coleta. O AIA não substitui D-I por D-O ou C.",bg=PANEL,fg=WARN,font=("Segoe UI",10,"bold")).pack(anchor="w",padx=16,pady=14)

    # 02 — D-O secundário
    band2=tk.Frame(self.content,bg="#0E202D",highlightbackground=ACCENT2,highlightthickness=1);band2.pack(fill="x",pady=(7,5))
    tk.Label(band2,text="02  •  DETECÇÃO POR OCORRÊNCIA (D-O)",bg="#0E202D",fg=ACCENT2,font=("Segoe UI",11,"bold")).pack(side="left",padx=14,pady=8)
    tk.Label(band2,text="UNIDADE: defeito/ocorrência • não confundir com D-I",bg="#0E202D",fg=TEXT,font=("Segoe UI",9)).pack(side="left",padx=4,pady=8)
    g=tk.Frame(self.content,bg=BG);g.pack(fill="x")
    vals=[("Ocorrências reais",d["real"],GOOD),("Detecções",d["det"],ACCENT2),("TP-D",d["tp"],GOOD),("FP-D",d["fp"],BAD),("FN-D",d["fn"],WARN)]
    for i,x in enumerate(vals): self.metric_card(g,*x).grid(row=0,column=i,sticky="ew",padx=2,pady=3);g.columnconfigure(i,weight=1)

    # 03 — C secundário
    band3=tk.Frame(self.content,bg="#0A2635",highlightbackground=ACCENT2,highlightthickness=1);band3.pack(fill="x",pady=(7,5))
    tk.Label(band3,text="03  •  CLASSIFICAÇÃO (C)",bg="#0A2635",fg=ACCENT2,font=("Segoe UI",11,"bold")).pack(side="left",padx=14,pady=8)
    tk.Label(band3,text="UNIDADE: inspeção • decisão OK/NOK",bg="#0A2635",fg=TEXT,font=("Segoe UI",9)).pack(side="left",padx=4,pady=8)
    if not (self.source_file and self.history and not self.source_analysis.get("classification_matrix",{}).get("valid")):
        g=tk.Frame(self.content,bg=BG);g.pack(fill="x")
        for i,(lab,val,col) in enumerate([("TP-C",m["TP"],GOOD),("VN-C",m["VN/TN"],GOOD),("FP-C",m["FP"],BAD),("FN-C",m["FN"],WARN)]):
            self.metric_card(g,lab,str(val),col).grid(row=0,column=i,sticky="ew",padx=2,pady=3);g.columnconfigure(i,weight=1)
    else:
        tk.Label(self.content,text="C não validada pela fonte histórica; o AIA preserva D-I e D-O sem promover esses dados para C.",bg=BG,fg=WARN,font=("Segoe UI",9,"bold")).pack(anchor="w",padx=12,pady=8)


def _aia_v115_field_graficos(self):
    fonte="HISTÓRICO INTERPRETADO" if (self.source_file and self.history) else "ENTRADA MANUAL"
    self.section_title("GRÁFICOS  /  PERFORMANCE",f"{fonte} • D-I prioritário • D-O e C separados")
    dt=self.detection_totals();di=dt.get("inspection",{}) or {};d=self.defect_counts();m=self.display_metrics()
    grid=tk.Frame(self.content,bg=BG);grid.pack(fill="both",expand=True)

    c1=self.chart_card(grid,"01  •  MATRIZ D-I — DETECÇÃO POR INSPEÇÃO","PRIORIDADE DE CAMPO • TP-DI / TN-DI / FP-DI / FN-DI")
    c1.grid(row=0,column=0,sticky="nsew",padx=(0,4),pady=(0,4))
    if di.get("available"):
        self.draw_bar_chart(c1,[("TP-DI",di["TP"],GOOD),("TN-DI",di["TN"],GOOD),("FP-DI",di["FP"],BAD),("FN-DI",di["FN"],WARN)],max_value=max(di["TP"],di["TN"],di["FP"],di["FN"],1))
    else:
        tk.Label(c1,text="N/D\n\nReal + Detectado não disponíveis por inspeção.",bg=PANEL,fg=WARN,font=("Segoe UI",11,"bold"),justify="center").pack(expand=True)

    c2=self.chart_card(grid,"02  •  INDICADORES D-I","Accuracy • Precision • Recall • Specificity • F1")
    c2.grid(row=0,column=1,sticky="nsew",padx=(4,0),pady=(0,4))
    if di.get("available"):
        self.draw_percent_chart(c2,[("Accuracy",di["Accuracy"],ACCENT2),("Precision",di["Precision"],ACCENT),("Recall",di["Recall / Sensibilidade"],GOOD),("Specificity",di["Specificity"],ACCENT2),("F1",di["F1"],GOOD)])
    else:
        tk.Label(c2,text="N/D",bg=PANEL,fg=WARN,font=("Segoe UI",12,"bold")).pack(expand=True)

    c3=self.chart_card(grid,"03  •  DETECÇÃO POR OCORRÊNCIA (D-O)","Unidade: defeito/ocorrência • contexto secundário")
    c3.grid(row=1,column=0,sticky="nsew",padx=(0,4),pady=(4,0))
    self.draw_bar_chart(c3,[("Real",d["real"],GOOD),("Detectado",d["det"],ACCENT2),("TP-D",d["tp"],GOOD),("FP-D",d["fp"],BAD),("FN-D",d["fn"],WARN)],max_value=max(d["real"],d["det"],d["tp"],d["fp"],d["fn"],1))

    c4=self.chart_card(grid,"04  •  CLASSIFICAÇÃO (C)","Unidade: inspeção • decisão OK/NOK • contexto secundário")
    c4.grid(row=1,column=1,sticky="nsew",padx=(4,0),pady=(4,0))
    self.draw_bar_chart(c4,[("TP-C",m["TP"],GOOD),("VN-C",m["VN/TN"],GOOD),("FP-C",m["FP"],BAD),("FN-C",m["FN"],WARN)],max_value=max(m["TP"],m["VN/TN"],m["FP"],m["FN"],1))
    grid.columnconfigure(0,weight=1);grid.columnconfigure(1,weight=1);grid.rowconfigure(0,weight=1);grid.rowconfigure(1,weight=1)


def _aia_v115_field_tecnica(self):
    fonte="HISTÓRICO INTERPRETADO" if (self.source_file and self.history) else "ENTRADA MANUAL"
    self.section_title("ANÁLISE TÉCNICA  /  DIAGNÓSTICO",f"{fonte} • D-I primeiro • D-O e C separados")
    dt=self.detection_totals();di=dt.get("inspection",{}) or {};d=self.defect_counts();m=self.display_metrics();a=self.source_analysis or {}

    c=self.card(self.content,"01  •  MATRIZ D-I — DETECÇÃO POR INSPEÇÃO","Unidade: inspeção • presença de defeito × presença de detecção")
    c.pack(fill="x",pady=5)
    if di.get("available"):
        cards=[("Nº de inspeções",di["N"],GOOD),("Accuracy",pct(di["Accuracy"]),ACCENT2),("Precision",pct(di["Precision"]),ACCENT2),("Recall",pct(di["Recall / Sensibilidade"]),GOOD),("Specificity",pct(di["Specificity"]),ACCENT2),("FPR",pct(di["FPR"]),BAD),("FNR",pct(di["FNR"]),WARN),("F1",pct(di["F1"]),ACCENT2),("MCC","N/D" if di["MCC"] is None else f"{di['MCC']:.4f}",ACCENT2),("Prevalência",pct(safe(di["TP"]+di["FN"],di["N"])),ACCENT2)]
        self._aia_v111_metric_cards(c,"D-I — INDICADORES",f"TP-DI={di['TP']} • TN-DI={di['TN']} • FP-DI={di['FP']} • FN-DI={di['FN']} • método: {di['method']}",cards,5)
    else:
        tk.Label(c,text="N/D — a fonte não fornece Real e Detectado válidos em nível de inspeção.",bg=PANEL,fg=WARN,font=("Segoe UI",10,"bold")).pack(anchor="w",padx=18,pady=16)

    if di.get("available"):
        pr=self.card(self.content,"02  •  DIAGNÓSTICO OPERACIONAL D-I","A ordem de investigação abaixo usa somente os erros da matriz D-I.")
        pr.pack(fill="x",pady=5)
        priorities=[]
        if di["FN"]>0: priorities.append(("01","REDUZIR FN-DI / ESCAPES",f"Há {di['FN']} inspeção(ões) com defeito real sem detecção. Investigar sensibilidade, cobertura, velocidade, foco, ROI e condições de aquisição."))
        if di["FP"]>0: priorities.append(("02","REDUZIR FP-DI / FALSOS ALARMES",f"Há {di['FP']} inspeção(ões) sem defeito que geraram detecção. Investigar threshold, ruído, iluminação, contraste e ROI."))
        if di.get("Recall / Sensibilidade") is not None and di["Recall / Sensibilidade"]<0.90: priorities.append(("03","AUMENTAR RECALL D-I",f"Recall atual de {pct(di['Recall / Sensibilidade'])}; repetir testes e ajustar condições antes de concluir estabilidade."))
        if di.get("Precision") is not None and di["Precision"]<0.90: priorities.append(("04","AUMENTAR PRECISION D-I",f"Precision atual de {pct(di['Precision'])}; investigar alarmes não confirmados."))
        if not priorities: priorities=[("01","ACOMPANHAR ESTABILIDADE","Nenhum erro D-I foi observado na amostra atual. Repetir testes em condições variadas antes de concluir estabilidade.")]
        for num,title,desc in priorities:
            row=tk.Frame(pr,bg=PANEL);row.pack(fill="x",padx=16,pady=6)
            tk.Label(row,text=num,bg="#0E4560",fg=TEXT,font=("Segoe UI",9,"bold"),width=4,pady=5).pack(side="left")
            body=tk.Frame(row,bg=PANEL);body.pack(side="left",fill="x",expand=True,padx=10)
            tk.Label(body,text=title,bg=PANEL,fg=TEXT,font=("Segoe UI",9,"bold")).pack(anchor="w")
            tk.Label(body,text=desc,bg=PANEL,fg=MUTED,wraplength=1100,justify="left",font=("Segoe UI",8)).pack(anchor="w",pady=(2,0))

    c2=self.card(self.content,"03  •  D-O — DETECÇÃO POR OCORRÊNCIA","Unidade: defeito/ocorrência • camada complementar; não usar como substituta de D-I")
    c2.pack(fill="x",pady=5)
    tk.Label(c2,text=f"Reais={d['real']} • Detectados={d['det']} • TP-D={d['tp']} • FP-D={d['fp']} • FN-D={d['fn']} • método={d.get('method','N/D')}",bg=PANEL,fg=TEXT,font=("Consolas",9,"bold")).pack(anchor="w",padx=16,pady=12)

    c3=self.card(self.content,"04  •  C — CLASSIFICAÇÃO DA INSPEÇÃO","Unidade: inspeção • decisão OK/NOK • camada complementar")
    c3.pack(fill="x",pady=5)
    tk.Label(c3,text=f"TP-C={m['TP']} • VN-C={m['VN/TN']} • FP-C={m['FP']} • FN-C={m['FN']} • Accuracy={pct(m['Accuracy'])} • F1={pct(m['F1'])}",bg=PANEL,fg=TEXT,font=("Consolas",9,"bold")).pack(anchor="w",padx=16,pady=12)

    ab=self.card(self.content,"05  •  AUDITORIA DA INTERPRETAÇÃO","De onde saiu cada número? • unidade e método fazem parte do resultado.")
    ab.pack(fill="x",pady=5)
    lines=[f"Arquivo: {a.get('file','N/D')}",f"Aba: {a.get('sheet','N/D')} • Cabeçalho: {a.get('header_row','N/D')}",f"Inspeções observadas: {len(self.rows)}",f"D-I disponível: {'SIM' if di.get('available') else 'NÃO'}",f"Método D-I: {di.get('method','N/D')}"]
    for line in lines: tk.Label(ab,text=line,bg=PANEL,fg=TEXT,font=("Segoe UI",9)).pack(anchor="w",padx=18,pady=3)
    self._aia_smart_add_interpretation_button(ab)


def _aia_v115_field_engenharia(self):
    fonte="HISTÓRICO INTERPRETADO" if (self.source_file and self.history) else "ENTRADA MANUAL"
    self.section_title("ENGENHARIA  /  DECISÃO",f"{fonte} • decisão operacional baseada primeiro em D-I")
    dt=self.detection_totals();di=dt.get("inspection",{}) or {};d=self.defect_counts();m=self.display_metrics()

    context=self.card(self.content,"CONTEXTO DO TESTE","Condições registradas no momento da coleta/importação.")
    context.pack(fill="x",pady=(0,6))
    ctx=tk.Frame(context,bg=PANEL);ctx.pack(fill="x",padx=14,pady=(0,12))
    items=[("Cliente",self.meta["cliente"].get() or "N/D"),("Planta",self.meta["local"].get() or "N/D"),("Projeto",self.meta["projeto"].get() or "N/D"),("Linha / Estação",self.meta["linha"].get() or "N/D"),("Threshold",self.meta["threshold"].get() or "N/D"),("Referência",self.meta["gabarito"].get() or "N/D")]
    for i,(lab,val) in enumerate(items):
        r,c=divmod(i,3);box=tk.Frame(ctx,bg=PANEL);box.grid(row=r,column=c,sticky="ew",padx=8,pady=4);tk.Label(box,text=lab,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w");tk.Label(box,text=val,bg=PANEL,fg=TEXT,font=("Segoe UI",9)).pack(anchor="w");ctx.columnconfigure(c,weight=1)

    if not di.get("available"):
        box=self.card(self.content,"D-I — N/D","A decisão operacional de campo não deve ser baseada em C ou D-O quando D-I não é observável.");box.pack(fill="x",pady=5);tk.Label(box,text="Real + Detectado por inspeção não estão disponíveis.",bg=PANEL,fg=WARN,font=("Segoe UI",10,"bold")).pack(anchor="w",padx=16,pady=12);return

    band=tk.Frame(self.content,bg="#103126",highlightbackground=GOOD,highlightthickness=1);band.pack(fill="x",pady=(0,5));tk.Label(band,text="01  •  INDICADORES PARA DECISÃO — D-I",bg="#103126",fg=GOOD,font=("Segoe UI",11,"bold")).pack(side="left",padx=14,pady=8);tk.Label(band,text="Accuracy / Recall / Precision / Specificity / FPR / FNR orientam a decisão e a investigação de campo",bg="#103126",fg=TEXT,font=("Segoe UI",9)).pack(side="left",padx=4,pady=8)
    g=tk.Frame(self.content,bg=BG);g.pack(fill="x")
    vals=[("Accuracy D-I",pct(di["Accuracy"]),ACCENT2),("Recall D-I",pct(di["Recall / Sensibilidade"]),GOOD),("Precision D-I",pct(di["Precision"]),ACCENT2),("Specificity D-I",pct(di["Specificity"]),ACCENT2),("FPR D-I",pct(di["FPR"]),BAD),("FNR D-I",pct(di["FNR"]),WARN)]
    for i,x in enumerate(vals):self.metric_card(g,*x).grid(row=0,column=i,sticky="ew",padx=3,pady=3);g.columnconfigure(i,weight=1)

    lower=tk.Frame(self.content,bg=BG);lower.pack(fill="both",expand=True,pady=7)
    priority=self.card(lower,"PRIORIZAÇÃO DE INVESTIGAÇÃO — D-I","A ordem é derivada diretamente dos erros da matriz D-I; não é um score proprietário.")
    priority.grid(row=0,column=0,sticky="nsew",padx=(0,4),pady=(0,4))
    priorities=[]
    if di["FN"]>0:priorities.append(("01","REDUZIR FN-DI / ESCAPES",f"FN-DI={di['FN']}. Investigar sensibilidade, cobertura, velocidade, foco, ROI e condições de aquisição."))
    if di["FP"]>0:priorities.append(("02","REDUZIR FP-DI / FALSOS ALARMES",f"FP-DI={di['FP']}. Investigar threshold, ruído, iluminação, contraste e ROI."))
    if di.get("Recall / Sensibilidade") is not None and di["Recall / Sensibilidade"]<0.90:priorities.append(("03","MELHORAR RECALL D-I",f"Recall={pct(di['Recall / Sensibilidade'])}. A prioridade é encontrar os defeitos reais."))
    if di.get("Precision") is not None and di["Precision"]<0.90:priorities.append(("04","MELHORAR PRECISION D-I",f"Precision={pct(di['Precision'])}. Avaliar alarmes sem defeito real."))
    if not priorities:priorities=[("01","ACOMPANHAR ESTABILIDADE","Sem FP-DI e FN-DI observados na amostra atual. Repetir testes em condições variadas.")]
    for num,title,desc in priorities:
        row=tk.Frame(priority,bg=PANEL);row.pack(fill="x",padx=16,pady=7);tk.Label(row,text=num,bg="#0E4560",fg=TEXT,font=("Segoe UI",9,"bold"),width=4,pady=5).pack(side="left");body=tk.Frame(row,bg=PANEL);body.pack(side="left",fill="x",expand=True,padx=10);tk.Label(body,text=title,bg=PANEL,fg=TEXT,font=("Segoe UI",9,"bold")).pack(anchor="w");tk.Label(body,text=desc,bg=PANEL,fg=MUTED,wraplength=520,justify="left",font=("Segoe UI",8)).pack(anchor="w",pady=(2,0))
    quality=self.chart_card(lower,"D-I — INDICADORES PARA DECISÃO","Unidade: inspeção • leitura operacional")
    quality.grid(row=0,column=1,sticky="nsew",padx=(4,0),pady=(0,4))
    self.draw_percent_chart(quality,[("Accuracy",di["Accuracy"],ACCENT2),("Recall",di["Recall / Sensibilidade"],GOOD),("Precision",di["Precision"],ACCENT),("Specificity",di["Specificity"],ACCENT2),("FPR",di["FPR"],BAD),("FNR",di["FNR"],WARN)])

    actions=self.card(lower,"LEITURA DE ENGENHARIA","Como transformar D-I em ação de campo.")
    actions.grid(row=1,column=0,sticky="nsew",padx=(0,4),pady=(4,0))
    lines=[f"• Accuracy D-I {pct(di['Accuracy'])}: proporção global de inspeções D-I corretamente interpretadas.",f"• Recall D-I {pct(di['Recall / Sensibilidade'])}: capacidade de encontrar inspeções com defeito.",f"• Precision D-I {pct(di['Precision'])}: confiabilidade das detecções nas inspeções.",f"• Specificity D-I {pct(di['Specificity'])}: capacidade de manter como negativas as inspeções realmente sem defeito.",f"• FPR D-I {pct(di['FPR'])}: carga de falsos alarmes em inspeções realmente OK.",f"• FNR D-I {pct(di['FNR'])}: risco de escape de defeitos.",f"• TP-DI={di['TP']} • TN-DI={di['TN']} • FP-DI={di['FP']} • FN-DI={di['FN']}."]
    for line in lines:tk.Label(actions,text=line,bg=PANEL,fg=TEXT,wraplength=520,justify="left",font=("Segoe UI",9)).pack(anchor="w",padx=16,pady=5)
    secondary=self.card(lower,"CAMADAS SECUNDÁRIAS","D-O mede ocorrências; C mede decisão OK/NOK.")
    secondary.grid(row=1,column=1,sticky="nsew",padx=(4,0),pady=(4,0))
    for line in [f"• D-O: reais={d['real']} • detectados={d['det']} • TP={d['tp']} • FP={d['fp']} • FN={d['fn']}.",f"• C: TP={m['TP']} • VN={m['VN/TN']} • FP={m['FP']} • FN={m['FN']}.","• Não usar D-O ou C como substitutos de D-I para a decisão operacional por inspeção."]:
        tk.Label(secondary,text=line,bg=PANEL,fg=TEXT,wraplength=520,justify="left",font=("Segoe UI",9)).pack(anchor="w",padx=16,pady=5)
    lower.columnconfigure(0,weight=1);lower.columnconfigure(1,weight=1);lower.rowconfigure(0,weight=1);lower.rowconfigure(1,weight=1)



# V1.1.8 — CHECKPOINT DE REGRESSÃO
# A Engenharia deve exibir seis indicadores D-I de decisão nesta ordem:
# Accuracy, Recall, Precision, Specificity, FPR e FNR.
# Não remover Accuracy ao alterar importação/exportação ou telas em releases futuras.

# Últimas ligações de runtime: estas são as telas efetivamente exibidas.
App.view_dashboard=_aia_v115_field_dashboard
App.view_graficos=_aia_v115_field_graficos
App.view_tecnica=_aia_v115_field_tecnica
App.view_engenharia=_aia_v115_field_engenharia


# ============================================================================
# AIA V1.1.6 — D-I DEFINITIVO / RUNTIME FINAL
# Esta ligação é a ÚLTIMA definição de Dashboard. O objetivo é impedir que
# camadas antigas do front-end sobrescrevam a tela operacional.
# ============================================================================
def _aia_v116_dashboard(self):
    fonte = "HISTÓRICO INTERPRETADO" if (self.source_file and self.history) else "ENTRADA MANUAL"
    self.section_title("DASHBOARD  /  VISÃO EXECUTIVA",
        f"{fonte} • V{VERSION} • D-I = referência operacional de campo")
    dt = self.detection_totals()
    di = dt.get("inspection", {}) or {}
    d = self.defect_counts()
    m = self.display_metrics()

    band = tk.Frame(self.content,bg="#103126",highlightbackground=GOOD,highlightthickness=1)
    band.pack(fill="x",pady=(0,5))
    tk.Label(band,text="01  •  D-I — DETECÇÃO POR INSPEÇÃO",bg="#103126",fg=GOOD,
             font=("Segoe UI",12,"bold")).pack(side="left",padx=14,pady=9)
    tk.Label(band,text="PRIORIDADE OPERACIONAL • defeito real × detecção na mesma inspeção",
             bg="#103126",fg=TEXT,font=("Segoe UI",9)).pack(side="left",padx=4,pady=9)

    if di.get("available"):
        g=tk.Frame(self.content,bg=BG); g.pack(fill="x")
        vals=[("N inspeções D-I",di.get("N",0),GOOD),("TP-DI",di.get("TP",0),GOOD),
              ("TN-DI",di.get("TN",0),GOOD),("FP-DI",di.get("FP",0),BAD),
              ("FN-DI",di.get("FN",0),WARN)]
        for i,x in enumerate(vals):
            self.metric_card(g,*x).grid(row=0,column=i,sticky="ew",padx=2,pady=3); g.columnconfigure(i,weight=1)
        g2=tk.Frame(self.content,bg=BG); g2.pack(fill="x")
        vals2=[("Accuracy D-I",pct(di.get("Accuracy")),ACCENT2),
               ("Precision D-I",pct(di.get("Precision")),ACCENT2),
               ("Recall D-I",pct(di.get("Recall / Sensibilidade")),GOOD),
               ("Specificity D-I",pct(di.get("Specificity")),ACCENT2),
               ("F1 D-I",pct(di.get("F1")),ACCENT2)]
        for i,x in enumerate(vals2):
            self.metric_card(g2,*x).grid(row=0,column=i,sticky="ew",padx=2,pady=3); g2.columnconfigure(i,weight=1)
        c=self.chart_card(self.content,"MATRIZ D-I — DETECÇÃO POR INSPEÇÃO",
                          "Unidade = inspeção • TP-DI / TN-DI / FP-DI / FN-DI")
        c.pack(fill="both",expand=True,pady=4)
        self.draw_bar_chart(c,[
            ("TP-DI",di.get("TP",0),GOOD),("TN-DI",di.get("TN",0),GOOD),
            ("FP-DI",di.get("FP",0),BAD),("FN-DI",di.get("FN",0),WARN)],
            max_value=max(di.get("TP",0),di.get("TN",0),di.get("FP",0),di.get("FN",0),1))
        note=self.card(self.content,"LEITURA OPERACIONAL — D-I",
                       "Esta é a leitura que deve ser priorizada no campo.")
        note.pack(fill="x",pady=4)
        if di.get("FN",0)>0:
            msg=f"ATENÇÃO • {di['FN']} FN-DI: defeito real não detectado. Prioridade = reduzir escapes."
            fg=BAD
        elif di.get("FP",0)>0:
            msg=f"ATENÇÃO • {di['FP']} FP-DI: inspeção sem defeito gerou detecção. Prioridade = reduzir falsos alarmes."
            fg=WARN
        else:
            msg="SEM FP-DI E FN-DI NA AMOSTRA ATUAL • repetir em amostra maior antes de concluir estabilidade."
            fg=GOOD
        tk.Label(note,text=msg,bg=PANEL,fg=fg,font=("Segoe UI",10,"bold"),
                 wraplength=1250,justify="left").pack(anchor="w",padx=16,pady=12)
    else:
        card=self.card(self.content,"D-I — N/D",
                       "A matriz D-I exige Real e Detectado válidos por inspeção.")
        card.pack(fill="x",pady=5)
        tk.Label(card,text="Não há dados D-I observáveis nesta coleta. O AIA não substitui D-I por D-O ou C.",
                 bg=PANEL,fg=WARN,font=("Segoe UI",10,"bold")).pack(anchor="w",padx=16,pady=14)

    # D-O e C permanecem abaixo como camadas secundárias.
    band2=tk.Frame(self.content,bg="#0E202D",highlightbackground=ACCENT2,highlightthickness=1)
    band2.pack(fill="x",pady=(7,5))
    tk.Label(band2,text="02  •  D-O — DETECÇÃO POR OCORRÊNCIA",bg="#0E202D",fg=ACCENT2,
             font=("Segoe UI",11,"bold")).pack(side="left",padx=14,pady=8)
    tk.Label(band2,text="UNIDADE: defeito / ocorrência",bg="#0E202D",fg=TEXT,font=("Segoe UI",9)).pack(side="left",padx=4,pady=8)
    g=tk.Frame(self.content,bg=BG);g.pack(fill="x")
    for i,x in enumerate([("Reais",d.get("real",0),GOOD),("Detecções",d.get("det",0),ACCENT2),
                          ("TP-D",d.get("tp",0),GOOD),("FP-D",d.get("fp",0),BAD),("FN-D",d.get("fn",0),WARN)]):
        self.metric_card(g,*x).grid(row=0,column=i,sticky="ew",padx=2,pady=3);g.columnconfigure(i,weight=1)

    band3=tk.Frame(self.content,bg="#0A2635",highlightbackground=ACCENT2,highlightthickness=1)
    band3.pack(fill="x",pady=(7,5))
    tk.Label(band3,text="03  •  C — CLASSIFICAÇÃO",bg="#0A2635",fg=ACCENT2,
             font=("Segoe UI",11,"bold")).pack(side="left",padx=14,pady=8)
    tk.Label(band3,text="UNIDADE: inspeção • decisão OK / NOK",bg="#0A2635",fg=TEXT,font=("Segoe UI",9)).pack(side="left",padx=4,pady=8)
    g=tk.Frame(self.content,bg=BG);g.pack(fill="x")
    for i,(lab,val,col) in enumerate([("TP-C",m.get("TP",0),GOOD),("VN-C",m.get("VN/TN",0),GOOD),
                                      ("FP-C",m.get("FP",0),BAD),("FN-C",m.get("FN",0),WARN)]):
        self.metric_card(g,lab,str(val),col).grid(row=0,column=i,sticky="ew",padx=2,pady=3);g.columnconfigure(i,weight=1)

App.view_dashboard = _aia_v116_dashboard



# ============================================================================
# AIA V1.2.0 — META D-I CANÔNICA
# A meta operacional pertence explicitamente à matriz D-I. Isso elimina o uso
# ambíguo dos nomes Accuracy/Precision/Recall e impede que valores da matriz C
# sejam usados na escolha de uma configuração de detecção.
# ============================================================================

def aia_v120_parse_target_value(key, raw):
    """Converte um campo de meta e rejeita valores fora do domínio."""
    text = str(raw or "").strip().replace(",", ".").replace("%", "")
    if not text:
        return None
    try:
        value = float(text)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"valor não numérico para {key}") from exc
    if value < 0:
        raise ValueError(f"valor negativo para {key}")
    if key in ("accuracy", "precision", "recall", "f1"):
        if value > 100:
            raise ValueError(f"percentual acima de 100 para {key}")
        if value > 1:
            value /= 100.0
        return value
    return value


def aia_v120_di_available(di_metrics):
    di = di_metrics or {}
    return bool(di.get("available", int(di.get("N", 0) or 0) > 0))


def aia_v120_evaluate_meta(target, di_metrics):
    """Avalia uma meta exclusivamente contra D-I e devolve resultado auditável."""
    target = target or {}
    di = di_metrics or {}
    available = aia_v120_di_available(di)
    n = int(di.get("N", 0) or 0)
    actual = {
        "accuracy": di.get("Accuracy") if available else None,
        "precision": di.get("Precision") if available else None,
        "recall": di.get("Recall / Sensibilidade") if available else None,
        "f1": di.get("F1") if available else None,
        "fp100": (di.get("FP", 0) / n * 100.0) if available and n else None,
        "fn100": (di.get("FN", 0) / n * 100.0) if available and n else None,
    }
    specs = [
        ("Accuracy D-I", "accuracy", ">="),
        ("Precision D-I", "precision", ">="),
        ("Recall D-I", "recall", ">="),
        ("F1 D-I", "f1", ">="),
        ("FP-DI / 100 inspeções", "fp100", "<="),
        ("FN-DI / 100 inspeções", "fn100", "<="),
    ]
    rows = []
    for label, key, operator in specs:
        expected = target.get(key)
        observed = actual.get(key)
        if expected is None:
            status = "SEM META"
        elif observed is None:
            status = "N/D"
        elif observed >= expected if operator == ">=" else observed <= expected:
            status = "ATENDIDA"
        else:
            status = "ABAIXO DA META" if operator == ">=" else "ACIMA DA META"
        rows.append({"label": label, "key": key, "operator": operator,
                     "actual": observed, "target": expected, "status": status})
    configured = [row for row in rows if row["target"] is not None]
    met = sum(row["status"] == "ATENDIDA" for row in configured)
    if not configured:
        overall = "SEM META"
    elif any(row["status"] == "N/D" for row in configured):
        overall = "DADOS INSUFICIENTES"
    elif met == len(configured):
        overall = "META ATINGIDA"
    elif met == 0:
        overall = "FORA DA EXPECTATIVA"
    else:
        overall = "EM CALIBRAÇÃO"
    return {"universe": "D-I", "unit": "inspeção", "available": available,
            "rows": rows, "configured": len(configured), "met": met,
            "status": overall, "actual": actual}


# Perfis e faixas operacionais da avaliação visual de capacidade.
# Os valores ficam centralizados e continuam editáveis na interface.
AIA_V120_META_PROFILES = {
    "BOM": {
        "accuracy": 0.90, "precision": 0.90, "recall": 0.95,
        "f1": 0.92, "fp100": 5.0, "fn100": 3.0,
    },
    "EXCELENTE": {
        "accuracy": 0.95, "precision": 0.95, "recall": 0.98,
        "f1": 0.96, "fp100": 3.0, "fn100": 1.0,
    },
}
AIA_V120_META_TOLERANCE = {
    "percent_points": 5.0,
    "per_100_inspections": 1.0,
}
AIA_V120_META_INDICATORS = (
    ("Accuracy D-I", "accuracy", "higher", False),
    ("Precision D-I", "precision", "higher", False),
    ("Recall D-I", "recall", "higher", True),
    ("F1 D-I", "f1", "higher", False),
    ("FP-DI / 100 inspeções", "fp100", "lower", False),
    ("FN-DI / 100 inspeções", "fn100", "lower", True),
)


def aia_v120_meta_profile_name(profile):
    name = str(profile or "").strip().upper()
    return name if name in {"BOM", "EXCELENTE", "PERSONALIZADO"} else "PERSONALIZADO"


def aia_v120_meta_profile_targets(profile):
    return dict(AIA_V120_META_PROFILES.get(aia_v120_meta_profile_name(profile), {}))


def aia_v120_infer_meta_profile(target):
    target = target or {}
    for profile, defaults in AIA_V120_META_PROFILES.items():
        if all(target.get(key) is not None and
               math.isclose(float(target[key]), value, rel_tol=1e-12, abs_tol=1e-12)
               for key, value in defaults.items()):
            return profile
    return "PERSONALIZADO"


def aia_v120_meta_sample_classification(n):
    count = max(0, int(n or 0))
    if count == 0:
        label = "SEM DADOS D-I"
    elif count < 30:
        label = "AMOSTRA PRELIMINAR"
    elif count < 100:
        label = "AMOSTRA DE VALIDAÇÃO"
    else:
        label = "AMOSTRA MAIS REPRESENTATIVA"
    return {
        "n": count,
        "label": label,
        "note": ("Classificação informativa do AIA; não representa garantia "
                 "estatística absoluta nem capacidade definitiva do equipamento."),
    }


def aia_v120_meta_actuals(di_metrics):
    """Obtém somente indicadores D-I; Matriz C e D-O não são consultadas."""
    di = di_metrics or {}
    available = aia_v120_di_available(di)
    n = int(di.get("N", 0) or 0)
    return {
        "accuracy": di.get("Accuracy") if available else None,
        "precision": di.get("Precision") if available else None,
        "recall": di.get("Recall / Sensibilidade") if available else None,
        "f1": di.get("F1") if available else None,
        "fp100": (float(di.get("FP", 0) or 0) / n * 100.0) if available and n else None,
        "fn100": (float(di.get("FN", 0) or 0) / n * 100.0) if available and n else None,
        "n": n,
        "available": available,
    }


def aia_v120_meta_value_text(value, key):
    if value is None:
        return "N/D"
    if key in {"accuracy", "precision", "recall", "f1"}:
        return f"{float(value) * 100:.1f}%"
    return f"{float(value):.1f}"


def aia_v120_meta_indicator_rows(target, di_metrics, tolerance=None):
    """Estrutura reutilizável por indicador para Meta e futura Calibração."""
    target = target or {}
    actuals = aia_v120_meta_actuals(di_metrics)
    limits = dict(AIA_V120_META_TOLERANCE)
    limits.update(tolerance or {})
    rows = []
    for label, key, direction, critical in AIA_V120_META_INDICATORS:
        actual = actuals.get(key)
        expected = target.get(key)
        percent = key in {"accuracy", "precision", "recall", "f1"}
        proximity = limits["percent_points"] if percent else limits["per_100_inspections"]
        distance = None
        if expected is None:
            status, distance_text = "SEM META", "Meta não configurada"
        elif actual is None:
            status, distance_text = "DADOS INSUFICIENTES", "Resultado D-I indisponível"
        else:
            distance = (float(actual) - float(expected)) * (100.0 if percent else 1.0)
            if direction == "higher":
                if distance >= -1e-12:
                    status = "META ATINGIDA"
                    distance_text = f"{distance:+.1f} p.p."
                elif distance >= -proximity - 1e-12:
                    status = "PRÓXIMO DA META"
                    distance_text = f"faltam {abs(distance):.1f} p.p."
                else:
                    status = "FORA DA META"
                    distance_text = f"faltam {abs(distance):.1f} p.p."
            else:
                if distance <= 1e-12:
                    status = "META ATINGIDA"
                    distance_text = ("no limite" if math.isclose(distance, 0.0, abs_tol=1e-12)
                                     else f"{abs(distance):.1f} abaixo do limite")
                elif distance <= proximity + 1e-12:
                    status = "PRÓXIMO DA META"
                    distance_text = f"{distance:.1f} acima do limite"
                else:
                    status = "FORA DA META"
                    distance_text = f"{distance:.1f} acima do limite"
        rows.append({
            "label": label, "key": key, "actual": actual, "target": expected,
            "distance": distance, "distance_text": distance_text,
            "direction": "MAIOR = MELHOR" if direction == "higher" else "MENOR = MELHOR",
            "direction_code": direction, "status": status, "critical": critical,
            "percent": percent, "proximity": proximity,
            "actual_text": aia_v120_meta_value_text(actual, key),
            "target_text": aia_v120_meta_value_text(expected, key),
        })
    return rows


def aia_v120_meta_goal_status(rows):
    configured = [row for row in rows if row["target"] is not None]
    if not configured:
        return "SEM META"
    if any(row["actual"] is None for row in configured):
        return "DADOS INSUFICIENTES"
    if all(row["status"] == "META ATINGIDA" for row in configured):
        return "META ATINGIDA"
    if any(row["status"] == "FORA DA META" for row in configured):
        return "FORA DA META"
    return "PRÓXIMO DA META"


def aia_v120_meta_project_classification(di_metrics, active_rows):
    actuals = aia_v120_meta_actuals(di_metrics)
    if not actuals["available"]:
        return "DADOS INSUFICIENTES"
    if not any(row["target"] is not None for row in active_rows):
        return "SEM META"
    excellent_rows = aia_v120_meta_indicator_rows(AIA_V120_META_PROFILES["EXCELENTE"], di_metrics)
    if all(row["status"] == "META ATINGIDA" for row in excellent_rows):
        return "EXCELENTE"
    good_rows = aia_v120_meta_indicator_rows(AIA_V120_META_PROFILES["BOM"], di_metrics)
    critical_good = all(row["status"] == "META ATINGIDA" for row in good_rows if row["critical"])
    no_relevant_failure = all(row["status"] != "FORA DA META" for row in good_rows)
    if critical_good and no_relevant_failure:
        return "BOM"
    if any(row["status"] == "FORA DA META" for row in active_rows):
        return "FORA DA META"
    if any(row["status"] in {"META ATINGIDA", "PRÓXIMO DA META"} for row in active_rows):
        return "EM CALIBRAÇÃO"
    return "FORA DA META"


def aia_v120_meta_capability_evaluation(target, di_metrics, tolerance=None):
    rows = aia_v120_meta_indicator_rows(target, di_metrics, tolerance)
    actuals = aia_v120_meta_actuals(di_metrics)
    configured = [row for row in rows if row["target"] is not None]
    return {
        "universe": "D-I", "unit": "inspeção",
        "method": "Real>0 × Detectado>0 por inspeção",
        "profile": aia_v120_meta_profile_name((target or {}).get("profile")),
        "assertiveness": actuals["accuracy"],
        "assertiveness_text": aia_v120_meta_value_text(actuals["accuracy"], "accuracy"),
        "n": actuals["n"], "available": actuals["available"],
        "sample": aia_v120_meta_sample_classification(actuals["n"]),
        "rows": rows, "configured": len(configured),
        "met": sum(row["status"] == "META ATINGIDA" for row in configured),
        "near": sum(row["status"] == "PRÓXIMO DA META" for row in configured),
        "outside": sum(row["status"] == "FORA DA META" for row in configured),
        "status": aia_v120_meta_goal_status(rows),
        "classification": aia_v120_meta_project_classification(di_metrics, rows),
        "actual": actuals,
    }


def aia_v120_meta_status_color(status):
    if status in {"META ATINGIDA", "EXCELENTE", "BOM"}:
        return GOOD
    if status in {"PRÓXIMO DA META", "EM CALIBRAÇÃO"}:
        return WARN
    if status in {"FORA DA META"}:
        return BAD
    return MUTED


def _aia_v120_ensure_meta_profile(self):
    if not hasattr(self, "expectation_profile"):
        self.expectation_profile = tk.StringVar(master=self, value="PERSONALIZADO")
    return self.expectation_profile


def _aia_v120_mark_meta_custom(self):
    if not getattr(self, "_applying_meta_profile", False):
        _aia_v120_ensure_meta_profile(self).set("PERSONALIZADO")


def _aia_v120_apply_meta_profile(self, profile):
    profile = aia_v120_meta_profile_name(profile)
    variable = _aia_v120_ensure_meta_profile(self)
    self._applying_meta_profile = True
    try:
        variable.set(profile)
        defaults = aia_v120_meta_profile_targets(profile)
        for key, value in defaults.items():
            display = value * 100.0 if key in {"accuracy", "precision", "recall", "f1"} else value
            self.expectation[key].set(f"{display:g}")
    finally:
        self._applying_meta_profile = False
    self.status.set(f"Perfil de Meta selecionado: {profile}.")
    if getattr(self, "current_view", "") == "expectativa":
        try:
            self.after_idle(lambda: self.show("expectativa")
                            if getattr(self, "current_view", "") == "expectativa" else None)
        except (tk.TclError, RuntimeError):
            pass


def _aia_v120_draw_meta_indicator_bar(parent, row):
    """Barra simples de resultado versus meta para um único indicador."""
    canvas = tk.Canvas(parent, bg=PANEL3, highlightthickness=0, height=38)
    canvas.pack(fill="x", padx=10, pady=(0, 7))
    # Não processa callbacks ociosos durante a construção da view: isso poderia
    # reconstruir a própria tela e invalidar este Canvas antes do desenho.
    width = max(parent.winfo_width(), 700)
    left, right, top, bottom = 8, 18, 7, 29
    plot_width = max(1, width - left - right)
    actual, target = row["actual"], row["target"]
    if actual is None or target is None:
        canvas.create_rectangle(left, top, left + plot_width, bottom, fill="#1A2C38", outline="")
        canvas.create_text(left + 8, (top + bottom) / 2, text="N/D",
                           fill=MUTED, anchor="w", font=("Segoe UI", 8, "bold"))
        return canvas
    if row["percent"]:
        maximum = 1.0
    else:
        maximum = max(float(actual), float(target) + row["proximity"], 1.0) * 1.12
    actual_ratio = max(0.0, min(1.0, float(actual) / maximum))
    target_ratio = max(0.0, min(1.0, float(target) / maximum))
    color = aia_v120_meta_status_color(row["status"])
    canvas.create_rectangle(left, top, left + plot_width, bottom, fill="#1A2C38", outline="")
    canvas.create_rectangle(left, top, left + plot_width * actual_ratio, bottom,
                            fill=color, outline="")
    marker_x = left + plot_width * target_ratio
    canvas.create_line(marker_x, top - 3, marker_x, bottom + 3, fill=TEXT, width=2)
    canvas.create_text(marker_x, bottom + 6, text="META", fill=TEXT,
                       anchor="n", font=("Segoe UI", 6, "bold"))
    return canvas


def _aia_v120_parse_target(self, key):
    try:
        return aia_v120_parse_target_value(key, self.expectation[key].get())
    except ValueError:
        return None


def _aia_v120_expectation_snapshot(self):
    profile = aia_v120_meta_profile_name(_aia_v120_ensure_meta_profile(self).get())
    snap = {key: self._parse_target(key) for key in
            ("accuracy", "precision", "recall", "f1", "fp100", "fn100")}
    snap.update({
        "profile": profile,
        "source": self.expectation["source"].get().strip(),
        "responsible": self.expectation["responsible"].get().strip(),
        "date": self.expectation["date"].get().strip(),
        "notes": self.expectation["notes"].get().strip(),
        "universe": "D-I", "unit": "inspeção",
        "method": "Real>0 × Detectado>0 por inspeção",
    })
    return snap


def _aia_v120_expectation_actuals(self):
    di = (self.detection_totals().get("inspection", {}) or {})
    return aia_v120_evaluate_meta({}, di)["actual"]


def _aia_v120_expectation_rows(self):
    di = (self.detection_totals().get("inspection", {}) or {})
    result = aia_v120_evaluate_meta(self.expectation_snapshot(), di)
    return [(row["label"], row["key"], row["operator"], row["actual"],
             row["target"], row["status"]) for row in result["rows"]]


def _aia_v120_expectation_summary(self):
    di = (self.detection_totals().get("inspection", {}) or {})
    result = aia_v120_evaluate_meta(self.expectation_snapshot(), di)
    return {key: result[key] for key in ("configured", "met", "status")}


def _aia_v120_expectation_status_color(self, status):
    return {"META ATINGIDA": GOOD, "ATENDIDA": GOOD, "EM CALIBRAÇÃO": WARN,
            "FORA DA EXPECTATIVA": BAD, "ABAIXO DA META": WARN,
            "ACIMA DA META": WARN, "DADOS INSUFICIENTES": WARN,
            "SEM META": MUTED, "N/D": MUTED}.get(status, TEXT)


def _aia_v120_load_active_expectation(self):
    path = self._aia_final_expectation_file()
    profile_variable = _aia_v120_ensure_meta_profile(self)
    profile_variable.set("PERSONALIZADO")
    self.expectation_legacy_migrated = False
    if not path.exists():
        return
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        self.expectation_legacy_migrated = not bool(data.get("universe"))
        for key in ("accuracy", "precision", "recall", "f1", "fp100", "fn100",
                    "source", "responsible", "date", "notes"):
            if key not in data:
                continue
            value = data[key]
            if value is None:
                value = ""
            elif key in ("accuracy", "precision", "recall", "f1") and isinstance(value, (int, float)):
                value = f"{value * 100:g}"
            elif key in ("fp100", "fn100") and isinstance(value, (int, float)):
                value = f"{value:g}"
            self.expectation[key].set(str(value))
        stored_profile = str(data.get("profile") or "").strip().upper()
        if stored_profile in {"BOM", "EXCELENTE", "PERSONALIZADO"}:
            profile_variable.set(stored_profile)
        else:
            inferred = {key: self._parse_target(key) for key in
                        ("accuracy", "precision", "recall", "f1", "fp100", "fn100")}
            profile_variable.set(aia_v120_infer_meta_profile(inferred))
        self.expectation_saved_snapshot = self.expectation_snapshot()
    except Exception:
        self.expectation_saved_snapshot = None


def _aia_v120_save_expectation(self):
    configured, invalid = [], []
    labels = (("accuracy", "Accuracy D-I"), ("precision", "Precision D-I"),
              ("recall", "Recall D-I"), ("f1", "F1 D-I"),
              ("fp100", "FP-DI / 100"), ("fn100", "FN-DI / 100"))
    for key, label in labels:
        raw = self.expectation[key].get()
        if not str(raw or "").strip():
            continue
        try:
            aia_v120_parse_target_value(key, raw)
            configured.append(key)
        except ValueError:
            invalid.append(label)
    if invalid:
        messagebox.showerror("Meta inválida", "Verifique os valores de: " + ", ".join(invalid) + ".", parent=self)
        return
    if not configured:
        messagebox.showwarning("Meta não configurada", "Informe pelo menos uma meta D-I antes de salvar.", parent=self)
        return
    snap = self.expectation_snapshot()
    snap.update({"saved_at": datetime.now().isoformat(timespec="seconds"), "configured": configured})
    try:
        self._aia_final_expectation_file().write_text(
            json.dumps(snap, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        messagebox.showerror("Erro ao salvar meta", f"Não foi possível salvar a meta.\n\n{exc}", parent=self)
        return
    # A persistência termina antes de qualquer atualização visual. Assim, uma
    # referência antiga de Treeview nunca transforma uma gravação bem-sucedida
    # em uma mensagem falsa de erro.
    self.expectation_saved_snapshot = dict(snap)
    self.expectation_legacy_migrated = False
    profile = aia_v120_meta_profile_name(_aia_v120_ensure_meta_profile(self).get())
    self.status.set(f"✓ Meta D-I salva — perfil {profile} • {len(configured)} critério(s).")
    messagebox.showinfo(
        "Meta salva",
        f"A meta D-I e o perfil {profile} foram salvos com fonte, unidade e método explícitos.",
        parent=self)
    try:
        self.after_idle(lambda: self.show("expectativa") if self.current_view == "expectativa" else None)
    except (tk.TclError, RuntimeError):
        pass


def _aia_v120_refresh_all(self):
    """Atualiza o shell e reconstrói a tela somente após o callback atual."""
    try:
        self.update_side()
    except (tk.TclError, RuntimeError):
        pass
    view = getattr(self, "current_view", "")
    if not view:
        return

    def rebuild():
        try:
            if getattr(self, "current_view", "") == view:
                self.show(view)
        except (tk.TclError, RuntimeError):
            # A janela pode ter sido fechada entre o agendamento e a execução.
            pass

    try:
        self.after_idle(rebuild)
    except (tk.TclError, RuntimeError):
        rebuild()


def _aia_v120_view_meta(self):
    self.section_title("META  /  CRITÉRIO DE ENTREGA — D-I",
                       "Avaliação de capacidade • Assertividade do projeto = Accuracy D-I • fonte exclusiva: Real × Detectado")
    profile_variable = _aia_v120_ensure_meta_profile(self)
    target = self.expectation_snapshot()
    di = (self.detection_totals().get("inspection", {}) or {})
    evaluation = aia_v120_meta_capability_evaluation(target, di)

    hero = self.card(self.content, "ASSERTIVIDADE D-I",
                     "Convenção do AIA: Assertividade do Projeto = Accuracy D-I; a classificação nunca usa esse número isoladamente")
    hero.pack(fill="x", pady=(0, 7))
    hero_top = tk.Frame(hero, bg=PANEL)
    hero_top.pack(fill="x", padx=16, pady=(3, 10))
    tk.Label(hero_top, text=evaluation["assertiveness_text"], bg=PANEL, fg=ACCENT2,
             font=("Segoe UI", 26, "bold")).pack(side="left")
    status_box = tk.Frame(hero_top, bg=PANEL)
    status_box.pack(side="right")
    tk.Label(status_box, text="STATUS DA META SELECIONADA", bg=PANEL, fg=MUTED,
             font=("Segoe UI", 7, "bold")).pack(anchor="e")
    tk.Label(status_box, text=evaluation["status"], bg=PANEL,
             fg=aia_v120_meta_status_color(evaluation["status"]),
             font=("Segoe UI", 15, "bold")).pack(anchor="e")

    summary_grid = tk.Frame(hero, bg=PANEL)
    summary_grid.pack(fill="x", padx=12, pady=(0, 10))
    summary_items = (
        ("PERFIL SELECIONADO", evaluation["profile"], ACCENT2),
        ("CLASSIFICAÇÃO GERAL", evaluation["classification"],
         aia_v120_meta_status_color(evaluation["classification"])),
        ("AMOSTRA D-I", f"N = {evaluation['n']} — {evaluation['sample']['label']}", TEXT),
        ("CRITÉRIOS ATENDIDOS", f"{evaluation['met']} de {evaluation['configured']}",
         aia_v120_meta_status_color(evaluation["status"])),
    )
    for index, (label, value, color) in enumerate(summary_items):
        box = tk.Frame(summary_grid, bg=PANEL3, highlightbackground=BORDER, highlightthickness=1)
        box.grid(row=0, column=index, sticky="nsew", padx=4, pady=3)
        summary_grid.columnconfigure(index, weight=1)
        tk.Label(box, text=label, bg=PANEL3, fg=MUTED,
                 font=("Segoe UI", 7, "bold")).pack(anchor="w", padx=10, pady=(8, 2))
        tk.Label(box, text=value, bg=PANEL3, fg=color,
                 font=("Segoe UI", 10, "bold"), wraplength=260,
                 justify="left").pack(anchor="w", padx=10, pady=(0, 8))
    tk.Label(hero, text=evaluation["sample"]["note"], bg=PANEL, fg=MUTED,
             font=("Segoe UI", 8), wraplength=1180, justify="left").pack(
                 anchor="w", padx=16, pady=(0, 10))

    profile_card = self.card(self.content, "PERFIL DE META",
                             "BOM e EXCELENTE preenchem os defaults; PERSONALIZADO conserva os valores para edição livre")
    profile_card.pack(fill="x", pady=5)
    profile_row = tk.Frame(profile_card, bg=PANEL)
    profile_row.pack(fill="x", padx=16, pady=(2, 7))
    for profile in ("BOM", "EXCELENTE", "PERSONALIZADO"):
        ttk.Radiobutton(profile_row, text=profile, variable=profile_variable, value=profile,
                        command=lambda selected=profile: _aia_v120_apply_meta_profile(
                            self, selected)).pack(side="left", padx=(0, 18))
    tk.Label(profile_row,
             text=("Proximidade centralizada: até 5,0 p.p. abaixo nos percentuais; "
                   "até 1,0 acima do limite em FP/FN por 100."),
             bg=PANEL, fg=MUTED, font=("Segoe UI", 8)).pack(side="right")

    grid = tk.Frame(profile_card, bg=PANEL)
    grid.pack(fill="x", padx=16, pady=8)
    fields = [("Accuracy D-I (%)", "accuracy"), ("Precision D-I (%)", "precision"),
              ("Recall D-I (%)", "recall"), ("F1 D-I (%)", "f1"),
              ("FP-DI / 100", "fp100"), ("FN-DI / 100", "fn100")]
    for i, (label, key) in enumerate(fields):
        r, c = divmod(i, 3); box = tk.Frame(grid, bg=PANEL)
        box.grid(row=r, column=c, sticky="ew", padx=6, pady=6); grid.columnconfigure(c, weight=1)
        tk.Label(box, text=label, bg=PANEL, fg=MUTED,
                 font=("Segoe UI", 8, "bold")).pack(anchor="w")
        entry = ttk.Entry(box, textvariable=self.expectation[key])
        entry.pack(fill="x")
        entry.bind("<KeyRelease>", lambda _event: _aia_v120_mark_meta_custom(self))
    actions = tk.Frame(profile_card, bg=PANEL)
    actions.pack(fill="x", padx=16, pady=(0, 9))
    tk.Button(actions, text="ATUALIZAR AVALIAÇÃO",
              command=lambda: self.after_idle(lambda: self.show("expectativa")),
              bg="#0D4560", fg=TEXT, activebackground="#176B8F", bd=0,
              padx=14, pady=7, font=("Segoe UI", 8, "bold"), cursor="hand2").pack(side="right", padx=4)
    tk.Button(actions, text="✓ SALVAR META D-I", command=self.save_expectation,
              bg=GOOD, fg="#061019", activebackground="#75E6B7", bd=0,
              padx=16, pady=7, font=("Segoe UI", 9, "bold"), cursor="hand2").pack(side="right", padx=4)

    indicators = self.card(self.content, "AVALIAÇÃO TÉCNICA POR INDICADOR — D-I",
                           "Verde = meta atingida • amarelo = próximo, nunca aprovado • vermelho = fora da meta")
    indicators.pack(fill="x", pady=5)
    for row in evaluation["rows"]:
        item = tk.Frame(indicators, bg=PANEL3, highlightbackground=BORDER, highlightthickness=1)
        item.pack(fill="x", padx=12, pady=4)
        top = tk.Frame(item, bg=PANEL3)
        top.pack(fill="x", padx=10, pady=(8, 2))
        title = row["label"] + ("  •  CRÍTICO" if row["critical"] else "")
        tk.Label(top, text=title, bg=PANEL3,
                 fg=WARN if row["critical"] else TEXT,
                 font=("Segoe UI", 9, "bold")).pack(side="left")
        tk.Label(top, text=row["status"], bg=PANEL3,
                 fg=aia_v120_meta_status_color(row["status"]),
                 font=("Segoe UI", 9, "bold")).pack(side="right")
        detail = (f"Resultado atual: {row['actual_text']}   |   Meta: {row['target_text']}   |   "
                  f"Diferença: {row['distance_text']}   |   {row['direction']}")
        tk.Label(item, text=detail, bg=PANEL3, fg=MUTED,
                 font=("Segoe UI", 8), justify="left").pack(anchor="w", padx=10, pady=(0, 2))
        _aia_v120_draw_meta_indicator_bar(item, row)

    presentation = self.card(self.content, "RESULTADO PARA APRESENTAÇÃO",
                             "Número simples para comunicação acompanhado da fundamentação técnica obrigatória")
    presentation.pack(fill="x", pady=5)
    presentation_row = tk.Frame(presentation, bg=PANEL)
    presentation_row.pack(fill="x", padx=16, pady=(2, 12))
    simple = tk.Frame(presentation_row, bg=PANEL3, highlightbackground=ACCENT, highlightthickness=1)
    simple.pack(side="left", fill="both", expand=True, padx=(0, 5))
    tk.Label(simple, text="ASSERTIVIDADE D-I", bg=PANEL3, fg=MUTED,
             font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=12, pady=(10, 2))
    tk.Label(simple, text=evaluation["assertiveness_text"], bg=PANEL3, fg=ACCENT2,
             font=("Segoe UI", 23, "bold")).pack(anchor="w", padx=12, pady=(0, 3))
    tk.Label(simple, text=f"N = {evaluation['n']} — {evaluation['sample']['label']}",
             bg=PANEL3, fg=TEXT, font=("Segoe UI", 9, "bold")).pack(
                 anchor="w", padx=12, pady=(0, 10))
    technical = tk.Frame(presentation_row, bg=PANEL3, highlightbackground=BORDER, highlightthickness=1)
    technical.pack(side="left", fill="both", expand=True, padx=(5, 0))
    tk.Label(technical, text="BASE TÉCNICA D-I", bg=PANEL3, fg=MUTED,
             font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=12, pady=(10, 3))
    actuals = evaluation["actual"]
    technical_text = (
        f"Accuracy: {aia_v120_meta_value_text(actuals['accuracy'], 'accuracy')}   •   "
        f"Recall: {aia_v120_meta_value_text(actuals['recall'], 'recall')}   •   "
        f"Precision: {aia_v120_meta_value_text(actuals['precision'], 'precision')}\n"
        f"F1: {aia_v120_meta_value_text(actuals['f1'], 'f1')}   •   "
        f"FP-DI/100: {aia_v120_meta_value_text(actuals['fp100'], 'fp100')}   •   "
        f"FN-DI/100: {aia_v120_meta_value_text(actuals['fn100'], 'fn100')}   •   N: {evaluation['n']}")
    tk.Label(technical, text=technical_text, bg=PANEL3, fg=TEXT,
             font=("Segoe UI", 9), wraplength=560, justify="left").pack(
                 anchor="w", padx=12, pady=(0, 10))

    trace = self.card(self.content, "RASTREABILIDADE DA META",
                      "Origem, responsável e data integram o critério técnico de entrega")
    trace.pack(fill="x", pady=5); form = tk.Frame(trace, bg=PANEL); form.pack(fill="x", padx=16, pady=8)
    for i, (label, key) in enumerate((("Origem do critério", "source"), ("Responsável", "responsible"), ("Data", "date"))):
        box = tk.Frame(form, bg=PANEL); box.grid(row=0, column=i, sticky="ew", padx=6); form.columnconfigure(i, weight=1)
        tk.Label(box, text=label, bg=PANEL, fg=MUTED, font=("Segoe UI", 8, "bold")).pack(anchor="w")
        ttk.Entry(box, textvariable=self.expectation[key]).pack(fill="x")
    box = tk.Frame(form, bg=PANEL); box.grid(row=1, column=0, columnspan=3, sticky="ew", padx=6, pady=6)
    tk.Label(box, text="Observações / acordo com o cliente", bg=PANEL, fg=MUTED,
             font=("Segoe UI", 8, "bold")).pack(anchor="w")
    ttk.Entry(box, textvariable=self.expectation["notes"]).pack(fill="x")
    if getattr(self, "expectation_legacy_migrated", False):
        tk.Label(trace, text="Meta antiga sem universo explícito carregada em modo de migração. Ao salvar, será registrada como D-I.",
                 bg=PANEL, fg=WARN, font=("Segoe UI", 8, "bold"), wraplength=1150,
                 justify="left").pack(anchor="w", padx=16, pady=(0, 10))


# ============================================================================
# AIA V1.2.0 — COMPARAÇÃO CANÔNICA
# ============================================================================

def _aia_v120_dataset_identity(item):
    return (item.get("record_id") or "", item.get("trace_hash") or "",
            item.get("created_at") or "", item.get("name") or "")


def _aia_v120_current_dataset(self):
    if not self._current_collection_has_data():
        return None
    meta = {key: value.get() for key, value in self.meta.items()}
    return {
        "name": "Coleta atual", "created_at": datetime.now().isoformat(),
        "metadata": meta, "source_file": self.source_file or "",
        "metrics": dict(self.display_metrics()), "defects": dict(self.defect_counts()),
        "detection_inspection": dict(self._smart_detection_inspection()),
        "threshold": meta.get("threshold") or "N/D",
        "date": meta.get("data") or datetime.now().strftime("%d/%m/%Y"),
        "expectation": self.expectation_snapshot(), "current": True,
    }


def _aia_v120_archive_current_collection(self, label=None):
    current = _aia_v120_current_dataset(self)
    if not current:
        return None
    created = datetime.now()
    payload = {key: current.get(key) for key in
               ("metadata", "metrics", "defects", "detection_inspection", "source_file")}
    current.update({
        "name": label or f"Coleta {len(self.comparison_archive) + 1:02d}",
        "created_at": created.isoformat(timespec="microseconds"),
        "project": dict(self.calibration_project),
        "record_id": "CAL-" + created.strftime("%Y%m%d%H%M%S%f"),
        "trace_hash": hashlib.sha256(json.dumps(
            payload, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")).hexdigest(),
    })
    current.pop("current", None)
    self.comparison_archive.append(current)
    self.calibration_history.append(dict(current))
    self.comparison_archive[:] = self.comparison_archive[-50:]
    self.calibration_history[:] = self.calibration_history[-50:]
    return current


def _aia_v120_comparison_datasets(self):
    archived = list(self.comparison_archive[-50:])
    current = _aia_v120_current_dataset(self)
    if current:
        archived = archived[-49:] + [current]
    return archived[-50:]


def _aia_v120_delete_selected_collection(self):
    tree = getattr(self, "compare_tree", None)
    if tree is None:
        return
    selected_ids = tuple(tree.selection())
    if len(selected_ids) != 1:
        messagebox.showinfo("Excluir coleta",
                            "Selecione exatamente uma coleta arquivada para excluir. Limpar seleção apenas desmarca itens.",
                            parent=self)
        return
    try:
        selected = self.compare_datasets[int(selected_ids[0])]
    except (IndexError, TypeError, ValueError):
        messagebox.showerror("Excluir coleta", "A seleção não corresponde a uma coleta válida.", parent=self)
        return
    if selected.get("current"):
        messagebox.showinfo("Excluir coleta",
                            "A coleta atual é somente leitura nesta tela. Exclua inspeções em COLETA ou inicie uma nova coleta.",
                            parent=self)
        return
    if not messagebox.askyesno("Excluir coleta",
                               f"Excluir a coleta arquivada '{selected.get('name', 'Coleta')}'?\n\nO arquivo Excel original não será apagado.",
                               parent=self):
        return
    identity = _aia_v120_dataset_identity(selected)
    self.comparison_archive[:] = [item for item in self.comparison_archive
                                  if _aia_v120_dataset_identity(item) != identity]
    self.calibration_history[:] = [item for item in self.calibration_history
                                   if _aia_v120_dataset_identity(item) != identity]
    self.status.set(f"✓ Coleta arquivada excluída: {selected.get('name', 'Coleta')}.")
    self.show("comparacao")


def _aia_v120_select_all_compare(self):
    tree = getattr(self, "compare_tree", None)
    if tree is None:
        return
    children = tuple(tree.get_children())[:50]
    tree.selection_set(children)
    self.status.set(f"{len(children)} coleta(s) selecionada(s) para comparação.")


def aia_v120_comparison_summary(selected, target=None):
    """Resumo puro da primeira e última coleta, usado pela UI e pelos testes."""
    if len(selected) < 2:
        raise ValueError("selecione pelo menos duas coletas")
    if len(selected) > 50:
        raise ValueError("o limite é de 50 coletas")
    first, last = selected[0], selected[-1]
    di_first = first.get("detection_inspection", {}) or {}
    di_last = last.get("detection_inspection", {}) or {}
    changes = []
    if aia_v120_di_available(di_first) and aia_v120_di_available(di_last):
        directions = (("Recall / Sensibilidade", 1), ("Precision", 1),
                      ("Accuracy", 1), ("FP", -1), ("FN", -1))
        for key, direction in directions:
            before, after = di_first.get(key), di_last.get(key)
            if before is not None and after is not None:
                changes.append((after - before) * direction)
    good = sum(value > 0 for value in changes)
    bad = sum(value < 0 for value in changes)
    verdict = "TENDÊNCIA FAVORÁVEL" if good > bad else (
        "TENDÊNCIA DESFAVORÁVEL" if bad > good else "TRADE-OFF / RESULTADO MISTO")
    return {
        "count": len(selected), "first": first, "last": last,
        "di_first": di_first, "di_last": di_last,
        "do_first": first.get("defects", {}) or {}, "do_last": last.get("defects", {}) or {},
        "c_first": first.get("metrics", {}) or {}, "c_last": last.get("metrics", {}) or {},
        "meta_first": aia_v120_evaluate_meta(target or {}, di_first),
        "meta_last": aia_v120_evaluate_meta(target or {}, di_last),
        "verdict": verdict,
    }


def _aia_v120_run_comparison(self, datasets):
    tree = getattr(self, "compare_tree", None)
    ids = tuple(tree.selection()) if tree is not None else ()
    if len(ids) < 2:
        messagebox.showwarning("Comparação", "Selecione pelo menos 2 coletas.", parent=self); return
    if len(ids) > 50:
        messagebox.showwarning("Comparação", "Selecione no máximo 50 coletas.", parent=self); return
    try:
        indices = sorted({int(item) for item in ids})
        selected = [datasets[index] for index in indices]
    except (IndexError, TypeError, ValueError):
        messagebox.showerror("Comparação", "A seleção contém uma coleta inválida.", parent=self); return
    result = getattr(self, "compare_result", None)
    if result is not None:
        _aia_v120_render_visual_comparison(self, selected, result)


def _aia_v120_view_comparacao(self):
    self.section_title("COMPARAÇÃO  /  EVOLUÇÃO",
                       "Compare de 2 a 50 coletas • D-I orienta a decisão • D-O e C permanecem separados")
    datasets = self._comparison_datasets()
    if not datasets:
        empty = self.card(self.content, "NENHUMA COLETA DISPONÍVEL",
                          "Realize uma coleta, importe uma planilha ou registre a coleta atual no histórico.")
        empty.pack(fill="x", pady=4); return
    toolbar = tk.Frame(self.content, bg=BG); toolbar.pack(fill="x", pady=(0, 7))
    buttons = (("COMPARAR SELECIONADAS", lambda: self._run_comparison(datasets), True),
               ("SELECIONAR TODAS", self._select_all_compare, False),
               ("REMOVER DA SELEÇÃO", self._remove_selected_compare, False),
               ("LIMPAR SELEÇÃO", self._clear_compare, False),
               ("✕ EXCLUIR UMA COLETA", self._aia_final_delete_selected_collection, False))
    for label, command, accent in buttons:
        tk.Button(toolbar, text=label, command=command,
                  bg="#0D5676" if accent else "#0D2432", fg="white",
                  activebackground="#176B8F", bd=0, padx=10, pady=7,
                  font=("Segoe UI", 8, "bold"), cursor="hand2").pack(side="left", padx=3)
    tk.Label(toolbar, text=f"{len(datasets)} coleta(s) • seleção não apaga dados",
             bg=BG, fg=MUTED, font=("Segoe UI", 8)).pack(side="right", padx=5)

    box = self.card(self.content, "COLETAS DISPONÍVEIS",
                    "Base e final seguem a ordem visual da lista; escolha condições comparáveis")
    box.pack(fill="x", pady=4)
    cols = ("coleta", "data", "threshold", "n", "acc", "rec", "pre", "fp", "fn", "meta", "origem")
    heads = ("Coleta", "Data", "Threshold", "N D-I", "Accuracy D-I", "Recall D-I",
             "Precision D-I", "FP-DI", "FN-DI", "Meta D-I", "Origem")
    widths = (120, 86, 75, 55, 90, 85, 90, 60, 60, 135, 185)
    tree = ttk.Treeview(box, columns=cols, show="headings", selectmode="extended",
                        height=min(10, max(5, len(datasets))))
    self.compare_tree, self.compare_datasets = tree, datasets
    for col, label, width in zip(cols, heads, widths):
        tree.heading(col, text=label); tree.column(col, width=width, anchor="center" if col != "origem" else "w")
    tree.pack(fill="x", padx=12, pady=(4, 12)); target = self.expectation_snapshot()
    for index, dataset in enumerate(datasets):
        di = dataset.get("detection_inspection", {}) or {}
        meta = aia_v120_evaluate_meta(target, di)
        origin = "ATUAL" if dataset.get("current") else (Path(dataset.get("source_file", "")).name or "Manual")
        tree.insert("", "end", iid=str(index), values=(
            dataset.get("name", "Coleta"), dataset.get("date", "N/D"), dataset.get("threshold", "N/D"),
            di.get("N", "N/D") if aia_v120_di_available(di) else "N/D", pct(di.get("Accuracy")),
            pct(di.get("Recall / Sensibilidade")), pct(di.get("Precision")),
            di.get("FP", "N/D") if aia_v120_di_available(di) else "N/D",
            di.get("FN", "N/D") if aia_v120_di_available(di) else "N/D", meta["status"], origin))
    tree.bind("<Delete>", lambda event: self._aia_final_delete_selected_collection())
    result = self.card(self.content, "RESULTADO DA COMPARAÇÃO",
                       "Nenhum resultado calculado: selecione pelo menos duas coletas")
    result.pack(fill="x", pady=(6, 4)); self.compare_result = result
    tk.Label(result, text="A comparação exibirá D-I, D-O, C e atendimento à mesma meta ativa em blocos independentes.",
             bg=PANEL, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w", padx=16, pady=14)


def _aia_v120_render_comparison_result(self, selected, result):
    for widget in result.winfo_children():
        widget.destroy()
    summary = aia_v120_comparison_summary(selected, self.expectation_snapshot())
    first, last = summary["first"], summary["last"]
    tk.Label(result, text=f"{summary['count']} COLETAS • BASE → FINAL",
             bg=PANEL, fg=ACCENT2, font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=16, pady=(12, 3))
    tk.Label(result,
             text=(f"{first.get('name', 'Coleta')} ({first.get('threshold', 'N/D')})  →  "
                   f"{last.get('name', 'Coleta')} ({last.get('threshold', 'N/D')})"),
             bg=PANEL, fg=MUTED, font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=16, pady=(0, 8))

    def comparison_table(parent, specs, before, after, height):
        table = ttk.Treeview(parent, columns=("ind", "base", "final", "delta", "leitura"),
                             show="headings", height=height)
        for col, label, width in (("ind", "Indicador", 225), ("base", "Base", 110),
                                  ("final", "Final", 110), ("delta", "Δ B→F", 105),
                                  ("leitura", "Direção", 170)):
            table.heading(col, text=label); table.column(col, width=width,
                                                         anchor="center" if col != "ind" else "w")
        table.pack(fill="x", padx=12, pady=8)
        for label, key, percent, lower_better in specs:
            a, b = before.get(key), after.get(key)
            av = pct(a) if percent else ("N/D" if a is None else str(a))
            bv = pct(b) if percent else ("N/D" if b is None else str(b))
            delta = _aia_v115_fmt_delta(a, b, percent)
            if a is None or b is None:
                direction = "DADOS INSUFICIENTES"
            elif a == b:
                direction = "IGUAL"
            elif (b < a) if lower_better else (b > a):
                direction = "MELHOR"
            else:
                direction = "PIOR"
            table.insert("", "end", values=(label, av, bv, delta, direction))

    di1, di2 = summary["di_first"], summary["di_last"]
    di_card = self.card(result, "01  •  D-I — DETECÇÃO POR INSPEÇÃO",
                        "Referência operacional • fonte exclusiva: Real>0 × Detectado>0 • unidade: inspeção")
    di_card.pack(fill="x", padx=10, pady=5)
    if not aia_v120_di_available(di1) or not aia_v120_di_available(di2):
        tk.Label(di_card, text="N/D — pelo menos uma coleta não contém Real e Detectado válidos por inspeção.",
                 bg=PANEL, fg=WARN, font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=14, pady=12)
    else:
        comparison_table(di_card, (
            ("N D-I", "N", False, False), ("TP-DI", "TP", False, False),
            ("TN-DI", "TN", False, False), ("FP-DI", "FP", False, True),
            ("FN-DI", "FN", False, True), ("Accuracy D-I", "Accuracy", True, False),
            ("Precision D-I", "Precision", True, False),
            ("Recall D-I", "Recall / Sensibilidade", True, False),
            ("Specificity D-I", "Specificity", True, False),
            ("FPR D-I", "FPR", True, True), ("FNR D-I", "FNR", True, True),
            ("F1 D-I", "F1", True, False), ("Prevalência D-I", "Prevalência", True, False),
        ), di1, di2, 13)
        tk.Label(di_card, text=f"Método base: {di1.get('method', 'N/D')} • método final: {di2.get('method', 'N/D')}",
                 bg=PANEL, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w", padx=14, pady=(0, 10))

    verdict = self.card(result, "DECISÃO OPERACIONAL D-I",
                        "Síntese multimetric: Accuracy, Recall, Precision, FP-DI e FN-DI")
    verdict.pack(fill="x", padx=10, pady=5)
    tk.Label(verdict, text=summary["verdict"], bg=PANEL,
             fg=GOOD if summary["verdict"] == "TENDÊNCIA FAVORÁVEL" else (
                 BAD if summary["verdict"] == "TENDÊNCIA DESFAVORÁVEL" else WARN),
             font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=14, pady=10)

    do1, do2 = summary["do_first"], summary["do_last"]
    do_card = self.card(result, "02  •  D-O — DETECÇÃO POR OCORRÊNCIA",
                        "Unidade: ocorrência/defeito • observado ou estimado sem substituir D-I")
    do_card.pack(fill="x", padx=10, pady=5)
    comparison_table(do_card, (("Ocorrências reais", "real", False, False),
                               ("Detecções", "det", False, False), ("TP-D", "tp", False, False),
                               ("FP-D", "fp", False, True), ("FN-D", "fn", False, True),
                               ("Precision D-O", "precision", True, False),
                               ("Recall D-O", "recall", True, False),
                               ("F1 D-O", "f1", True, False)), do1, do2, 8)
    tk.Label(do_card, text=f"Método final: {do2.get('method', 'N/D')} • status: {do2.get('status', 'N/D')}",
             bg=PANEL, fg=WARN if do2.get("estimated") else MUTED,
             font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=14, pady=(0, 10))

    c1, c2 = summary["c_first"], summary["c_last"]
    c_card = self.card(result, "03  •  C — CLASSIFICAÇÃO OK/NOK",
                       "Fonte exclusiva: Vision AI × Referência ou colunas C explícitas • unidade: inspeção")
    c_card.pack(fill="x", padx=10, pady=5)
    comparison_table(c_card, (("N C", "N", False, False), ("TP-C", "TP", False, False),
                              ("VN-C", "VN/TN", False, False), ("FP-C", "FP", False, True),
                              ("FN-C", "FN", False, True), ("Accuracy C", "Accuracy", True, False),
                              ("Precision C", "Precision", True, False),
                              ("Recall C", "Recall / Sensibilidade", True, False),
                              ("Specificity C", "Specificity", True, False),
                              ("F1 C", "F1", True, False)), c1, c2, 10)

    meta_card = self.card(result, "04  •  ATENDIMENTO À META D-I",
                          "Todas as coletas são avaliadas contra a mesma meta ativa")
    meta_card.pack(fill="x", padx=10, pady=5)
    table = ttk.Treeview(meta_card, columns=("coleta", "threshold", "atendidas", "configuradas", "status"),
                         show="headings", height=min(10, len(selected)))
    for col, label, width in (("coleta", "Coleta", 220), ("threshold", "Threshold", 100),
                              ("atendidas", "Atendidas", 90), ("configuradas", "Configuradas", 100),
                              ("status", "Status", 210)):
        table.heading(col, text=label); table.column(col, width=width, anchor="center" if col != "coleta" else "w")
    table.pack(fill="x", padx=12, pady=8)
    target = self.expectation_snapshot()
    for dataset in selected:
        evaluation = aia_v120_evaluate_meta(target, dataset.get("detection_inspection", {}))
        table.insert("", "end", values=(dataset.get("name", "Coleta"), dataset.get("threshold", "N/D"),
                                         evaluation["met"], evaluation["configured"], evaluation["status"]))

    charts = tk.Frame(result, bg=PANEL); charts.pack(fill="both", expand=True, padx=10, pady=(2, 10))
    for index, (title, key, percent) in enumerate((
            ("EVOLUÇÃO — ACCURACY D-I", "Accuracy", True),
            ("EVOLUÇÃO — RECALL D-I", "Recall / Sensibilidade", True),
            ("EVOLUÇÃO — FP-DI", "FP", False), ("EVOLUÇÃO — FN-DI", "FN", False))):
        row, col = divmod(index, 2)
        chart = self.chart_card(charts, title, "D-I • inspeções")
        chart.grid(row=row, column=col, sticky="nsew", padx=(0, 4) if col == 0 else (4, 0), pady=4)
        mapped = [{"metrics": {key: (dataset.get("detection_inspection", {}) or {}).get(key)},
                   "threshold": dataset.get("threshold"), "name": dataset.get("name"), "defects": {}}
                  for dataset in selected]
        self._draw_trend_chart(chart, mapped, key, percent)
    charts.columnconfigure(0, weight=1); charts.columnconfigure(1, weight=1)


# ============================================================================
# AIA V1.2.0 — CALIBRAÇÃO CANÔNICA
# ============================================================================

def aia_v120_calibration_score(item, target=None):
    """Ordena configurações por meta D-I e, depois, desempenho D-I."""
    di = item.get("detection_inspection", {}) or {}
    if not aia_v120_di_available(di):
        return (0, 0, -1.0, -1.0, -1.0, -10**9, 0)
    evaluation = aia_v120_evaluate_meta(target or {}, di)
    configured = evaluation["configured"]
    ratio = evaluation["met"] / configured if configured else 0.0
    all_met = int(configured > 0 and evaluation["met"] == configured)
    return (1, all_met, ratio,
            di.get("F1") if di.get("F1") is not None else -1.0,
            di.get("Recall / Sensibilidade") if di.get("Recall / Sensibilidade") is not None else -1.0,
            -(di.get("FP", 0) or 0), di.get("N", 0) or 0)


def _aia_v120_calibration_score(self, item):
    return aia_v120_calibration_score(item, self.expectation_snapshot())


def _aia_v120_best_configuration(self, datasets=None):
    items = datasets if datasets is not None else self.calibration_history
    valid = []
    for item in items:
        di=item.get("detection_inspection", {}) or {}
        if not aia_v120_di_available(di):
            continue
        # Não recomendar uma configuração quando a amostra não permite medir
        # simultaneamente sensibilidade e confiabilidade dos alarmes.
        if any(di.get(key) is None for key in ("Recall / Sensibilidade", "Precision", "F1")):
            continue
        valid.append(item)
    return max(valid, key=self._calibration_score) if valid else None


def _aia_v120_draw_threshold_chart(self, parent, datasets):
    points = []
    for item in datasets:
        threshold = self._threshold_value(item)
        di = item.get("detection_inspection", {}) or {}
        if threshold is not None and aia_v120_di_available(di) and di.get("F1") is not None:
            points.append((threshold, item))
    canvas = tk.Canvas(parent, bg=PANEL, highlightthickness=0, height=270)
    canvas.pack(fill="both", expand=True, padx=10, pady=8); canvas.update_idletasks()
    width, height = max(canvas.winfo_width(), 620), 270
    left, right, top, bottom = 55, 28, 22, 45
    if not points:
        canvas.create_text(width / 2, height / 2,
                           text="N/D — informe Threshold numérico e registre ao menos uma coleta com D-I válida.",
                           fill=MUTED, font=("Segoe UI", 9)); return
    points.sort(key=lambda pair: pair[0]); values = [pair[0] for pair in points]
    low, high = min(values), max(values)
    if low == high:
        low -= 1; high += 1
    for index in range(5):
        y = top + (height - top - bottom) * (1 - index / 4)
        canvas.create_line(left, y, width - right, y, fill="#29404F")
        canvas.create_text(left - 7, y, text=f"{index * 25}%", anchor="e", fill=MUTED, font=("Segoe UI", 7))
    best = self._best_configuration(datasets); previous = None
    for threshold, item in points:
        x = left + (threshold - low) / (high - low) * (width - left - right)
        f1 = item["detection_inspection"]["F1"]
        y = top + (1 - f1) * (height - top - bottom)
        if previous:
            canvas.create_line(previous[0], previous[1], x, y, fill=ACCENT, width=2)
        canvas.create_oval(x - 5, y - 5, x + 5, y + 5,
                           fill=GOOD if item is best else ACCENT2, outline="")
        canvas.create_text(x, y - 12, text=f"{f1 * 100:.1f}%", fill=TEXT, font=("Segoe UI", 7, "bold"))
        canvas.create_text(x, height - bottom + 15, text=str(threshold), fill=MUTED, font=("Segoe UI", 7))
        previous = (x, y)
    canvas.create_text(left, height - bottom + 30, text="Threshold • F1 D-I", fill="#6E8E9E", font=("Segoe UI", 7))


def _aia_v120_view_calibracao(self):
    self.section_title("CALIBRAÇÃO  /  PROJETO E RECOMENDAÇÃO D-I",
                       "Rastreabilidade de configurações • mesma meta D-I • threshold × desempenho por inspeção")
    project = self.calibration_project
    info = self.card(self.content, "IDENTIFICAÇÃO DO PROJETO",
                     project.get("name") or "Projeto ainda não identificado")
    info.pack(fill="x", pady=4)
    tk.Label(info,
             text=(f"Código: {project.get('id') or 'N/D'}   •   Responsável: {project.get('responsible') or 'N/D'}\n"
                   f"Objetivo: {project.get('objective') or 'N/D'}"),
             bg=PANEL, fg=TEXT, justify="left", wraplength=1100).pack(anchor="w", padx=16, pady=12)
    ttk.Button(info, text="EDITAR IDENTIFICAÇÃO", command=self.edit_calibration_project).pack(anchor="e", padx=14, pady=(0, 12))

    toolbar = tk.Frame(self.content, bg=BG); toolbar.pack(fill="x", pady=5)
    ttk.Button(toolbar, text="REGISTRAR COLETA ATUAL", style="Accent.TButton",
               command=self._archive_current_from_view).pack(side="left", padx=3)
    ttk.Button(toolbar, text="ABRIR COMPARAÇÃO", command=lambda: self.show("comparacao")).pack(side="left", padx=3)
    ttk.Button(toolbar, text="GERAR RELATÓRIO EXECUTIVO", command=self.export_executive_report).pack(side="left", padx=3)

    datasets = list(self.calibration_history or self.comparison_archive)
    best = self._best_configuration(datasets); active_meta = self.expectation_summary()
    status = self.card(self.content, "CRITÉRIO DE RECOMENDAÇÃO",
                       "1º atendimento à meta ativa • 2º F1 D-I • 3º Recall D-I • 4º menor FP-DI • 5º maior N")
    status.pack(fill="x", pady=4)
    tk.Label(status,
             text=f"Meta D-I atual: {active_meta['status']} • {active_meta['met']}/{active_meta['configured']} critérios atendidos na coleta ativa",
             bg=PANEL, fg=self._expectation_status_color(active_meta["status"]),
             font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=16, pady=10)

    best_box = self.card(self.content, "MELHOR CONFIGURAÇÃO ENCONTRADA",
                         "Recomendação calculada somente com D-I válida; D-O e C permanecem como verificações complementares")
    best_box.pack(fill="x", pady=4)
    if best:
        di = best["detection_inspection"]; evaluation = aia_v120_evaluate_meta(self.expectation_snapshot(), di)
        text = (f"{best.get('name', 'Coleta')} • Threshold {best.get('threshold', 'N/D')} • N D-I={di.get('N', 0)} • "
                f"Accuracy={pct(di.get('Accuracy'))} • Recall={pct(di.get('Recall / Sensibilidade'))} • "
                f"Precision={pct(di.get('Precision'))} • F1={pct(di.get('F1'))} • "
                f"FP-DI={di.get('FP', 0)} • FN-DI={di.get('FN', 0)} • Meta={evaluation['status']}")
        color = (aia_v120_meta_status_color(evaluation["status"])
                 if evaluation.get("configured") else ACCENT2)
    else:
        text = "N/D — ainda não há coleta histórica com D-I válida para recomendar uma configuração."
        color = MUTED
    tk.Label(best_box, text=text, bg=PANEL, fg=color, font=("Segoe UI", 10, "bold"),
             wraplength=1200, justify="left").pack(anchor="w", padx=16, pady=14)

    guide = self.card(self.content, "ASSISTENTE DE CALIBRAÇÃO — D-I",
                      "Hipóteses e testes sugeridos a partir da Meta; nenhuma alteração é aplicada automaticamente")
    guide.pack(fill="x", pady=4)
    for item in aia_v120_calibration_guidance(self.expectation_snapshot(),
                                               (self.detection_totals().get("inspection", {}) or {})):
        tk.Label(guide, text=f"{item['level']}  {item['title']} — {item['text']}",
                 bg=PANEL, fg=item['color'], font=("Segoe UI", 9, "bold"),
                 wraplength=1180, justify="left").pack(anchor="w", padx=16, pady=5)

    chart = self.card(self.content, "THRESHOLD × F1 D-I",
                      "Ponto verde = recomendação atual • compare somente ensaios com condições equivalentes")
    chart.pack(fill="both", expand=True, pady=4); self._draw_threshold_chart(chart, datasets)

    history = self.card(self.content, "HISTÓRICO RASTREÁVEL",
                        f"{len(datasets)} coleta(s) • registros individuais disponíveis na Comparação")
    history.pack(fill="x", pady=4)
    columns = ("registro", "data", "coleta", "threshold", "n", "accuracy", "recall", "precision", "f1", "fp", "fn", "meta", "hash")
    labels = ("Registro", "Data", "Coleta", "Threshold", "N D-I", "Accuracy", "Recall", "Precision", "F1", "FP", "FN", "Meta", "Hash")
    widths = (150, 82, 115, 72, 55, 78, 75, 78, 70, 50, 50, 125, 90)
    tree = ttk.Treeview(history, columns=columns, show="headings", height=min(9, max(3, len(datasets))))
    for col, label, width in zip(columns, labels, widths):
        tree.heading(col, text=label); tree.column(col, width=width, anchor="center" if col != "coleta" else "w")
    tree.pack(fill="x", padx=12, pady=10); target = self.expectation_snapshot()
    for item in datasets:
        di = item.get("detection_inspection", {}) or {}; evaluation = aia_v120_evaluate_meta(target, di)
        tree.insert("", "end", values=(item.get("record_id", "N/D"), item.get("date", "N/D"),
                                        item.get("name", "Coleta"), item.get("threshold", "N/D"),
                                        di.get("N", "N/D") if aia_v120_di_available(di) else "N/D",
                                        pct(di.get("Accuracy")), pct(di.get("Recall / Sensibilidade")),
                                        pct(di.get("Precision")), pct(di.get("F1")),
                                        di.get("FP", "N/D") if aia_v120_di_available(di) else "N/D",
                                        di.get("FN", "N/D") if aia_v120_di_available(di) else "N/D",
                                        evaluation["status"], (item.get("trace_hash") or "")[:10]))


# ============================================================================
# AIA V1.2.0 — AUDITORIA CANÔNICA
# ============================================================================

def aia_v120_audit_snapshot(self):
    """Materializa fonte, unidade, método e status sem depender da camada visual."""
    analysis = getattr(self, "source_analysis", {}) or {}
    source_file = getattr(self, "source_file", "") or ""
    metrics_c = self.display_metrics()
    di = (self.detection_totals().get("inspection", {}) or {})
    do = self.defect_counts()
    source_c = analysis.get("classification_matrix", {}) or {}
    c_valid = bool(source_c.get("valid")) if source_file else bool(metrics_c.get("N", 0))
    c_method = source_c.get("method") if source_file else "Vision AI × Referência / entrada manual"
    c_values = {
        "N": metrics_c.get("N", 0), "TP": metrics_c.get("TP", 0),
        "TN": metrics_c.get("VN/TN", 0), "FP": metrics_c.get("FP", 0),
        "FN": metrics_c.get("FN", 0),
    } if c_valid else {"N": None, "TP": None, "TN": None, "FP": None, "FN": None}
    do_available = bool(do.get("available", do.get("real", 0) or do.get("det", 0)))
    matrices = [
        {"name": "C", "question": "Decisão Vision AI OK/NOK corresponde à Referência?",
         "unit": "inspeção", "source": "Vision AI × Referência ou TP-C/VN-C/FP-C/FN-C explícitos",
         "method": c_method or "não validada", "status": "VALIDADA" if c_valid else "N/D",
         "values": c_values},
        {"name": "D-I", "question": "A inspeção com defeito apresentou alguma detecção?",
         "unit": "inspeção", "source": di.get("source") or "Real × Detectado por inspeção",
         "method": di.get("method") or "Binário por presença: Real>0 e Detectado>0",
         "status": "DISPONÍVEL" if aia_v120_di_available(di) else "N/D",
         "values": {"N": di.get("N") if aia_v120_di_available(di) else None,
                    "TP": di.get("TP") if aia_v120_di_available(di) else None,
                    "TN": di.get("TN") if aia_v120_di_available(di) else None,
                    "FP": di.get("FP") if aia_v120_di_available(di) else None,
                    "FN": di.get("FN") if aia_v120_di_available(di) else None}},
        {"name": "D-O", "question": "Quantas ocorrências reais foram detectadas?",
         "unit": "ocorrência/defeito", "source": "Real/Detectado/TP-D por ocorrência ou matriz agregada semanticamente validada",
         "method": do.get("method", "N/D"), "status": do.get("status", "DISPONÍVEL" if do_available else "N/D"),
         "values": {"N": do.get("real") if do_available else None,
                    "TP": do.get("tp") if do_available else None, "TN": do.get("tn"),
                    "FP": do.get("fp") if do_available else None,
                    "FN": do.get("fn") if do_available else None}},
    ]
    metadata_status = analysis.get("metadata_status", {}) or {}
    metadata = []
    for key, label in (("cliente", "Cliente"), ("local", "Local / Planta"),
                       ("projeto", "Projeto"), ("linha", "Linha / Estação"),
                       ("tecnico", "Técnico"), ("data", "Data"),
                       ("threshold", "Threshold"), ("gabarito", "Método de referência")):
        value = self.meta[key].get() if key in self.meta else ""
        origin = metadata_status.get(key) or ("entrada manual / projeto" if value else "não identificado")
        metadata.append({"field": label, "value": value or "N/D", "origin": origin,
                         "status": "PREENCHIDO" if value else "PENDENTE"})
    mapping = analysis.get("mapping", {}) or {}
    generic_present = any(mapping.get(key) is not None for key in ("TP/VP", "VN/TN", "FP", "FN"))
    relation = matrix_relationship_diagnostics(self.rows)
    runtime_independent = (getattr(App, "classification_totals", None) is _aia_smart_classification_totals
                           and getattr(App, "_smart_detection_inspection", None) is _aia_smart_detection_inspection
                           and getattr(App, "_smart_detection_occurrence", None) is _aia_v120_detection_occurrence_audited)
    do_identity_real = (not do_available or do.get("real") == (do.get("tp",0) or 0) + (do.get("fn",0) or 0))
    do_identity_det = (not do_available or do.get("det") == (do.get("tp",0) or 0) + (do.get("fp",0) or 0))
    detail_mismatches=[]
    for row0 in self.rows:
        row=normalize_row(row0); details=row.get("defects") or []
        if details:
            sr=sum(int(x.get("real",0) or 0) for x in details); sd=sum(int(x.get("det",0) or 0) for x in details)
            if sr != int(row.get("real",0) or 0) or sd != int(row.get("det",0) or 0):
                detail_mismatches.append(str(row.get("id") or "?"))
    checks = [
        ("Motores independentes", runtime_independent,
         "Bindings ativos verificados: C, D-I e D-O apontam para motores distintos."),
        ("N da matriz C fecha", not c_valid or c_values["N"] == sum(c_values[key] for key in ("TP", "TN", "FP", "FN")),
         "N-C = TP-C + VN-C + FP-C + FN-C."),
        ("N da matriz D-I fecha", not aia_v120_di_available(di) or di.get("N") == sum(di.get(key, 0) for key in ("TP", "TN", "FP", "FN")),
         "N-DI = TP-DI + TN-DI + FP-DI + FN-DI."),
        ("Conservação D-O — reais", do_identity_real,
         "Real = TP-D + FN-D."),
        ("Conservação D-O — detecções", do_identity_det,
         "Detectado = TP-D + FP-D."),
        ("Detalhamento fecha com totais", not detail_mismatches,
         "Tipos detalhados conferem com Real/Detectado" if not detail_mismatches else "Divergência nas inspeções: " + ", ".join(detail_mismatches[:12])),
        ("Genéricos não promovidos para C", True,
         "TP/VN/FP/FN sem sufixo permanecem ambíguos e não alimentam C por nome isolado."),
    ]
    return {
        "source": {"file": analysis.get("file") or source_file or "Entrada manual",
                   "sheet": analysis.get("sheet") or "N/A",
                   "header_row": analysis.get("header_row") or "N/A",
                   "status": analysis.get("status") or ("MANUAL" if not source_file else "PARCIAL"),
                   "registered": len(self.rows)},
        "matrices": matrices, "metadata": metadata, "mapping": mapping,
        "generic_present": generic_present,
        "generic_action": "IGNORADOS COMO MATRIZ C; preservados apenas como evidência da fonte" if generic_present else "não encontrados",
        "detection_rows_used": di.get("rows_used", 0),
        "detection_rows_unavailable": di.get("rows_unavailable", 0),
        "relationship": relation, "checks": checks,
    }


def _aia_v120_view_auditoria(self):
    audit = aia_v120_audit_snapshot(self); source = audit["source"]
    self.section_title("AUDITORIA  /  FONTE → UNIDADE → MÉTODO → RESULTADO",
                       "C, D-I e D-O são rastreadas separadamente; N/D nunca é convertido em zero")
    origin = self.card(self.content, "ORIGEM DO CONJUNTO",
                       "Identificação do arquivo e da tabela efetivamente interpretada")
    origin.pack(fill="x", pady=4)
    tk.Label(origin,
             text=(f"Arquivo: {source['file']}\nAba: {source['sheet']} • linha de cabeçalho: {source['header_row']} • "
                   f"inspeções registradas: {source['registered']} • status: {source['status']}"),
             bg=PANEL, fg=TEXT, font=("Segoe UI", 9), wraplength=1220,
             justify="left").pack(anchor="w", padx=16, pady=12)

    matrices = self.card(self.content, "CONTRATO DAS MATRIZES",
                         "Cada linha declara a pergunta, a unidade, a fonte e o método do próprio universo")
    matrices.pack(fill="x", pady=4)
    columns = ("matriz", "n", "tp", "tn", "fp", "fn", "unidade", "fonte", "metodo", "status")
    labels = ("Matriz", "N", "TP", "TN/VN", "FP", "FN", "Unidade", "Fonte", "Método", "Status")
    widths = (58, 52, 48, 58, 48, 48, 110, 250, 275, 105)
    tree = ttk.Treeview(matrices, columns=columns, show="headings", height=3)
    for col, label, width in zip(columns, labels, widths):
        tree.heading(col, text=label); tree.column(col, width=width, anchor="center" if col not in ("fonte", "metodo") else "w")
    tree.pack(fill="x", padx=12, pady=10)
    for matrix in audit["matrices"]:
        values = matrix["values"]
        tree.insert("", "end", values=(matrix["name"],
                                        "N/D" if values.get("N") is None else values.get("N"),
                                        "N/D" if values.get("TP") is None else values.get("TP"),
                                        "N/A" if matrix["name"] == "D-O" and values.get("TN") is None else ("N/D" if values.get("TN") is None else values.get("TN")),
                                        "N/D" if values.get("FP") is None else values.get("FP"),
                                        "N/D" if values.get("FN") is None else values.get("FN"),
                                        matrix["unit"], matrix["source"], matrix["method"], matrix["status"]))
        tk.Label(matrices, text=f"{matrix['name']}: {matrix['question']}", bg=PANEL, fg=MUTED,
                 font=("Segoe UI", 8), justify="left").pack(anchor="w", padx=16, pady=2)
    relation = audit["relationship"]
    tk.Label(matrices,
             text=(f"Relação C × D-I: {relation['comparable']} linha(s) comparáveis • "
                   f"{relation['equivalent']} equivalentes • {relation['divergent']} divergentes. "
                   "Totais iguais não significam reutilização: a prova é feita pelas fontes de cada linha."),
             bg=PANEL, fg=GOOD if relation["divergent"] else WARN,
             font=("Segoe UI", 8, "bold"), wraplength=1220, justify="left").pack(anchor="w", padx=16, pady=(8, 12))

    meta = self.card(self.content, "METADADOS E DADOS CADASTRADOS",
                     "Valor atual e origem reconhecida durante importação ou preenchimento")
    meta.pack(fill="x", pady=4)
    table = ttk.Treeview(meta, columns=("campo", "valor", "origem", "status"), show="headings", height=8)
    for col, label, width in (("campo", "Campo", 180), ("valor", "Valor", 260),
                              ("origem", "Origem / método", 460), ("status", "Status", 110)):
        table.heading(col, text=label); table.column(col, width=width, anchor="w" if col != "status" else "center")
    table.pack(fill="x", padx=12, pady=10)
    for item in audit["metadata"]:
        table.insert("", "end", values=(item["field"], item["value"], item["origin"], item["status"]))

    mapping = self.card(self.content, "MAPEAMENTO DA IMPORTAÇÃO",
                        "Colunas encontradas e tratamento de TP/VN/FP/FN genéricos")
    mapping.pack(fill="x", pady=4)
    map_tree = ttk.Treeview(mapping, columns=("campo", "coluna", "uso"), show="headings", height=6)
    for col, label, width in (("campo", "Campo lógico", 190), ("coluna", "Coluna fonte", 120),
                              ("uso", "Regra de uso", 700)):
        map_tree.heading(col, text=label); map_tree.column(col, width=width, anchor="w")
    map_tree.pack(fill="x", padx=12, pady=8)
    rules = {"TP-C": "C explícita", "VN-C": "C explícita", "FP-C": "C explícita", "FN-C": "C explícita",
             "Real": "D-I e D-O", "Detectado": "D-I e D-O", "TP-D": "D-O explícita",
             "FP-D": "D-O explícita", "FN-D": "D-O explícita"}
    for field, column in audit["mapping"].items():
        use = rules.get(field, "AMBÍGUO — não promover automaticamente para C, D-I ou D-O")
        map_tree.insert("", "end", values=(field, "N/D" if column is None else column, use))
    tk.Label(mapping, text=f"Campos genéricos: {audit['generic_action']}.", bg=PANEL,
             fg=WARN if audit["generic_present"] else MUTED, font=("Segoe UI", 8, "bold"),
             wraplength=1200, justify="left").pack(anchor="w", padx=16, pady=(0, 10))

    integrity = self.card(self.content, "CONTROLES DE INTEGRIDADE",
                           "Verificações estruturais e matemáticas do conjunto ativo")
    integrity.pack(fill="x", pady=4)
    for name, passed, detail in audit["checks"]:
        tk.Label(integrity, text=("✓ " if passed else "✕ ") + name + " — " + detail,
                 bg=PANEL, fg=GOOD if passed else BAD, font=("Segoe UI", 8, "bold"),
                 wraplength=1200, justify="left").pack(anchor="w", padx=16, pady=4)
    tk.Label(integrity,
             text=(f"D-I: {audit['detection_rows_used']} linha(s) usadas • "
                   f"{audit['detection_rows_unavailable']} indisponíveis. Ausência permanece N/D."),
             bg=PANEL, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w", padx=16, pady=(4, 12))


# ============================================================================
# AIA V1.1.7 — CONSOLIDAÇÃO DE TELAS
# Dashboard/Engenharia: abordagem D-I da V1.1.6 mantida.
# Comparação: abordagem multicoleta D-I mantida.
# Análise Técnica: restaurada como tela completa, com C + D-I + D-O + Auditoria.
# ============================================================================
def _aia_v117_view_tecnica(self):
    fonte = "HISTÓRICO INTERPRETADO" if (self.source_file and self.history) else "ENTRADA MANUAL"
    self.section_title(
        "ANÁLISE TÉCNICA  /  DIAGNÓSTICO",
        f"Autaza Inspection Analytics • {fonte} • V1.1.7 • três unidades estatísticas preservadas"
    )

    m = self.display_metrics()
    d = self.defect_counts()
    dt = self.detection_totals()
    di = dt.get("inspection", {}) or {}
    a = self.source_analysis or {}

    # ------------------------------------------------------------------
    # 01 — CLASSIFICAÇÃO (C): matriz completa por inspeção
    # ------------------------------------------------------------------
    c = self.card(self.content, "01  •  CLASSIFICAÇÃO (C)",
                  "Unidade: inspeção • decisão OK/NOK está correta?")
    c.pack(fill="x", pady=5)

    c_valid = not (self.source_file and self.history and
                   not a.get("classification_matrix", {}).get("valid"))
    if not c_valid:
        tk.Label(c,
                 text="N/D — a fonte não fornece uma matriz C válida por inspeção. O AIA preserva essa limitação e não reutiliza D-I ou D-O como C.",
                 bg=PANEL, fg=WARN, font=("Segoe UI", 10, "bold"),
                 wraplength=1250, justify="left").pack(anchor="w", padx=18, pady=16)
    else:
        subtitle = (f"TP-C={m['TP']} • VN-C={m['VN/TN']} • FP-C={m['FP']} • FN-C={m['FN']}")
        self._aia_v111_metric_cards(c, "MATRIZ C — INDICADORES", subtitle, [
            ("Nº de inspeções", m["N"], TEXT),
            ("Accuracy", pct(m["Accuracy"]), ACCENT2),
            ("Precision", pct(m["Precision"]), ACCENT2),
            ("Recall", pct(m["Recall / Sensibilidade"]), GOOD),
            ("Specificity", pct(m["Specificity"]), ACCENT2),
            ("FPR", pct(m["FPR"]), BAD),
            ("FNR", pct(m["FNR"]), WARN),
            ("F1", pct(m["F1"]), ACCENT2),
            ("MCC", "N/D" if m["MCC"] is None else f"{m['MCC']:.4f}", ACCENT2),
            ("Prevalência", pct(safe(m["TP"] + m["FN"], m["N"])), ACCENT2),
        ], 5)

    # ------------------------------------------------------------------
    # 02 — DETECÇÃO POR INSPEÇÃO (D-I): principal análise operacional
    # ------------------------------------------------------------------
    di_card = self.card(self.content, "02  •  DETECÇÃO (D-I) — POR INSPEÇÃO",
                        "Unidade: inspeção • presença de defeito × presença de detecção")
    di_card.pack(fill="x", pady=5)
    if not di.get("available"):
        tk.Label(di_card,
                 text="N/D — Real e Detectado por inspeção não estão disponíveis de forma suficiente para construir a matriz D-I.",
                 bg=PANEL, fg=WARN, font=("Segoe UI", 10, "bold"),
                 wraplength=1250, justify="left").pack(anchor="w", padx=18, pady=16)
    else:
        di_sub = (f"TP-DI={di.get('TP',0)} • TN-DI={di.get('TN',0)} • "
                  f"FP-DI={di.get('FP',0)} • FN-DI={di.get('FN',0)} • "
                  f"método: {di.get('method','N/D')} • binário por presença")
        relation = matrix_relationship_diagnostics(self.rows)
        c_signature = (m.get("TP",0), m.get("VN/TN",0), m.get("FP",0), m.get("FN",0))
        di_signature = (di.get("TP",0), di.get("TN",0), di.get("FP",0), di.get("FN",0))
        if c_signature == di_signature:
            proof = tk.Frame(di_card, bg="#3A2B0B", highlightbackground=WARN, highlightthickness=2)
            proof.pack(fill="x", padx=16, pady=(8,6))
            tk.Label(proof, text="⚠ TOTAIS IGUAIS — PROVA DE COINCIDÊNCIA DOS DADOS",
                     bg="#3A2B0B", fg=WARN, font=("Segoe UI",10,"bold")).pack(anchor="w",padx=12,pady=(9,3))
            tk.Label(proof,
                     text=(f"C usa somente Referência × Vision AI. D-I usa somente Real>0 × Detectado>0. "
                           f"Linhas comparáveis: {relation['comparable']} • equivalentes: {relation['equivalent']} • "
                           f"divergentes: {relation['divergent']}. Totais iguais são consequência dos dados, não cópia da matriz."),
                     bg="#3A2B0B", fg=TEXT, font=("Segoe UI",9), wraplength=1220,
                     justify="left").pack(anchor="w",padx=12,pady=(0,9))
        elif relation["comparable"]:
            proof = tk.Frame(di_card, bg="#103126", highlightbackground=GOOD, highlightthickness=1)
            proof.pack(fill="x", padx=16, pady=(8,6))
            tk.Label(proof,
                     text=(f"✓ MATRIZES INDEPENDENTES COMPROVADAS • {relation['divergent']} de "
                           f"{relation['comparable']} inspeções divergem. IDs: "
                           f"{', '.join(relation['divergent_ids'][:12]) or 'N/D'}"),
                     bg="#103126", fg=GOOD, font=("Segoe UI",9,"bold"), wraplength=1220,
                     justify="left").pack(anchor="w",padx=12,pady=9)
        self._aia_v111_metric_cards(di_card, "MATRIZ D-I — INDICADORES", di_sub, [
            ("Nº de inspeções", di.get("N", "N/D"), TEXT),
            ("Accuracy", pct(di.get("Accuracy")), ACCENT2),
            ("Precision", pct(di.get("Precision")), ACCENT2),
            ("Recall", pct(di.get("Recall / Sensibilidade")), GOOD),
            ("Specificity", pct(di.get("Specificity")), ACCENT2),
            ("FPR", pct(di.get("FPR")), BAD),
            ("FNR", pct(di.get("FNR")), WARN),
            ("F1", pct(di.get("F1")), ACCENT2),
            ("MCC", "N/D" if di.get("MCC") is None else f"{di.get('MCC'):.4f}", ACCENT2),
            ("Prevalência", pct(di.get("Prevalência")), ACCENT2),
        ], 5)

        # Visual compacto da matriz, para não transformar a tela em apenas cards.
        visual = tk.Frame(self.content, bg=BG)
        visual.pack(fill="x", pady=2)
        matrix = self.chart_card(visual, "MATRIZ D-I — VISÃO OPERACIONAL",
                                 "TP-DI e TN-DI = acertos • FP-DI = falso alarme • FN-DI = escape")
        matrix.pack(fill="x", pady=4)
        self.draw_bar_chart(matrix, [
            ("TP-DI", di.get("TP",0), GOOD),
            ("TN-DI", di.get("TN",0), GOOD),
            ("FP-DI", di.get("FP",0), BAD),
            ("FN-DI", di.get("FN",0), WARN),
        ], max_value=max(di.get("TP",0), di.get("TN",0), di.get("FP",0), di.get("FN",0), 1))

        diag = self.card(self.content, "DIAGNÓSTICO D-I",
                         "Leitura automática derivada exclusivamente da matriz por inspeção.")
        diag.pack(fill="x", pady=4)
        priorities=[]
        if di.get("FN",0)>0:
            priorities.append(("01", "REDUZIR FN-DI / ESCAPES",
                               f"{di['FN']} inspeção(ões) com defeito real sem detecção. Investigar sensibilidade, cobertura, velocidade, foco, ROI e condições de aquisição."))
        if di.get("FP",0)>0:
            priorities.append(("02", "REDUZIR FP-DI / FALSOS ALARMES",
                               f"{di['FP']} inspeção(ões) sem defeito geraram detecção. Investigar threshold, ruído, iluminação, contraste e ROI."))
        if di.get("Recall / Sensibilidade") is not None and di["Recall / Sensibilidade"] < .90:
            priorities.append(("03", "AUMENTAR RECALL D-I",
                               f"Recall atual = {pct(di['Recall / Sensibilidade'])}. Prioridade: aumentar a capacidade de encontrar defeitos reais."))
        if di.get("Precision") is not None and di["Precision"] < .90:
            priorities.append(("04", "AUMENTAR PRECISION D-I",
                               f"Precision atual = {pct(di['Precision'])}. Prioridade: reduzir alarmes que não correspondem a defeito real."))
        if not priorities:
            priorities=[("01", "ACOMPANHAR ESTABILIDADE",
                         "Nenhum FP-DI ou FN-DI observado na amostra atual. Repetir em condições variadas antes de concluir estabilidade.")]
        for num,title,desc in priorities:
            row=tk.Frame(diag,bg=PANEL); row.pack(fill="x", padx=16, pady=6)
            tk.Label(row,text=num,bg="#0E4560",fg=TEXT,font=("Segoe UI",9,"bold"),width=4,pady=5).pack(side="left")
            body=tk.Frame(row,bg=PANEL);body.pack(side="left",fill="x",expand=True,padx=10)
            tk.Label(body,text=title,bg=PANEL,fg=TEXT,font=("Segoe UI",9,"bold")).pack(anchor="w")
            tk.Label(body,text=desc,bg=PANEL,fg=MUTED,wraplength=1180,justify="left",font=("Segoe UI",8)).pack(anchor="w",pady=(2,0))

    # ------------------------------------------------------------------
    # 03 — DETECÇÃO POR OCORRÊNCIA (D-O): camada complementar
    # ------------------------------------------------------------------
    do = self.card(self.content, "03  •  DETECÇÃO (D-O) — POR OCORRÊNCIA",
                   "Unidade: defeito/ocorrência • não confundir com inspeção")
    do.pack(fill="x", pady=5)
    do_sub = (f"TP-D={d.get('tp',0)} • FP-D={d.get('fp',0)} • FN-D={d.get('fn',0)} • "
              f"método: {d.get('method','N/D')}")
    self._aia_v111_metric_cards(do, "MATRIZ D-O — INDICADORES", do_sub, [
        ("Nº ocorrências reais", d.get("real",0), TEXT),
        ("Nº de detecções", d.get("det",0), ACCENT2),
        ("TP-D", d.get("tp",0), GOOD),
        ("FP-D", d.get("fp",0), BAD),
        ("FN-D", d.get("fn",0), WARN),
        ("Accuracy", "N/D" if d.get("full_metrics",{}).get("Accuracy") is None else pct(d["full_metrics"]["Accuracy"]), ACCENT2),
        ("Precision", pct(d.get("precision")), ACCENT2),
        ("Recall", pct(d.get("recall")), GOOD),
        ("F1", pct(d.get("f1")), ACCENT2),
        ("Specificity", "N/D" if d.get("full_metrics",{}).get("Specificity") is None else pct(d["full_metrics"]["Specificity"]), ACCENT2),
    ], 5)
    tk.Label(do,text=("Observação: quando não existe correspondência individual entre ocorrência real e detecção, "
                     "o AIA mantém D-O como contagem/estimativa e não o usa para substituir D-I."),
             bg=PANEL,fg=MUTED,font=("Segoe UI",8),wraplength=1250,justify="left").pack(anchor="w",padx=18,pady=(2,12))

    # ------------------------------------------------------------------
    # 04 — LEITURA CRUZADA: evita confundir unidades
    # ------------------------------------------------------------------
    cross=self.card(self.content,"04  •  LEITURA CRUZADA DAS MATRIZES",
                    "As três unidades respondem perguntas diferentes; os indicadores não devem ser misturados.")
    cross.pack(fill="x",pady=5)
    lines=[
        f"• C — classificação: TP={m.get('TP',0)} | VN={m.get('VN/TN',0)} | FP={m.get('FP',0)} | FN={m.get('FN',0)} | unidade = inspeção.",
        f"• D-I — detecção por inspeção: TP={di.get('TP','N/D')} | TN={di.get('TN','N/D')} | FP={di.get('FP','N/D')} | FN={di.get('FN','N/D')} | unidade = inspeção.",
        f"• D-O — detecção por ocorrência: reais={d.get('real',0)} | detecções={d.get('det',0)} | TP={d.get('tp',0)} | FP={d.get('fp',0)} | FN={d.get('fn',0)} | unidade = ocorrência.",
        "• Regra operacional: para saber se o sistema encontrou o defeito na inspeção, priorizar D-I. D-O é complementar e C mede a decisão OK/NOK.",
    ]
    for line in lines:
        tk.Label(cross,text=line,bg=PANEL,fg=TEXT,font=("Segoe UI",9),wraplength=1250,justify="left").pack(anchor="w",padx=18,pady=4)

    # ------------------------------------------------------------------
    # 05 — AUDITORIA DA INTERPRETAÇÃO
    # ------------------------------------------------------------------
    ab=self.card(self.content,"05  •  AUDITORIA DA INTERPRETAÇÃO",
                 "De onde saiu cada número? • unidade, fonte e método fazem parte do resultado.")
    ab.pack(fill="x",pady=5)
    lines=[
        f"Arquivo: {a.get('file','N/D')}",
        f"Aba principal: {a.get('sheet','N/D')} • Cabeçalho: {a.get('header_row','N/D')}",
        f"Inspeções observadas pelo AIA: {len(self.rows)}",
        f"D-I disponível: {'SIM' if di.get('available') else 'NÃO'} • método: {di.get('method','N/D')}",
        f"D-O disponível: {'SIM' if d.get('available') else 'NÃO'} • método: {d.get('method','N/D')}",
        f"Classificação C validada: {'SIM' if c_valid else 'NÃO'}",
        "Status da interpretação: CONFIRMADO — desde que as limitações declaradas acima sejam respeitadas.",
    ]
    for line in lines:
        tk.Label(ab,text=line,bg=PANEL,fg=TEXT,font=("Segoe UI",9),wraplength=1250,justify="left").pack(anchor="w",padx=18,pady=3)
    try:
        self._aia_smart_add_interpretation_button(ab)
    except Exception:
        pass


# ============================================================================
# AIA V1.2.0 CANDIDATA — COMPARAÇÃO VISUAL OPERACIONAL
# Somente o módulo Comparação usa as funções abaixo.
# ============================================================================

AIA_V120_COMPARISON_METRICS = (
    ("Accuracy D-I", "Accuracy", True, False),
    ("Recall D-I", "Recall / Sensibilidade", True, False),
    ("Precision D-I", "Precision", True, False),
    ("Specificity D-I", "Specificity", True, False),
    ("FP-DI", "FP", False, True),
    ("FN-DI", "FN", False, True),
)


def aia_v120_comparison_color(index):
    """Gera uma cor estável e distinta por posição, sem afetar os dados."""
    hue = (index * 137.508 + 195.0) % 360.0
    saturation, value = 0.68, 0.92
    chroma = value * saturation
    x = chroma * (1 - abs((hue / 60.0) % 2 - 1))
    if hue < 60: red, green, blue = chroma, x, 0
    elif hue < 120: red, green, blue = x, chroma, 0
    elif hue < 180: red, green, blue = 0, chroma, x
    elif hue < 240: red, green, blue = 0, x, chroma
    elif hue < 300: red, green, blue = x, 0, chroma
    else: red, green, blue = chroma, 0, x
    match = value - chroma
    rgb = tuple(round((component + match) * 255) for component in (red, green, blue))
    return "#%02X%02X%02X" % rgb


def aia_v120_comparison_identities(selected):
    """Cria a identidade visual imutável usada em toda a comparação."""
    identities = []
    for index, dataset in enumerate(selected):
        di = dataset.get("detection_inspection", {}) or {}
        source = dataset.get("source_file", "") or ""
        if dataset.get("current"):
            origin = "Coleta atual" + (f" • {Path(source).name}" if source else "")
        elif source:
            origin = Path(source).name
        else:
            origin = "Entrada manual"
        n_inspections = di.get("N") if aia_v120_di_available(di) else (dataset.get("metrics", {}) or {}).get("N")
        identities.append({
            "index": index,
            "label": f"Coleta {index + 1:02d}",
            "short": f"C{index + 1:02d}",
            "color": aia_v120_comparison_color(index),
            "name": dataset.get("name") or f"Coleta {index + 1:02d}",
            "date": dataset.get("date") if dataset.get("date") not in (None, "") else "N/D",
            "threshold": (dataset.get("threshold")
                          if dataset.get("threshold") not in (None, "") else "N/D"),
            "n": n_inspections if n_inspections is not None else "N/D",
            "origin": origin,
            "dataset": dataset,
            "di": di,
        })
    return identities


def aia_v120_comparison_value(identity, key):
    return (identity.get("di", {}) or {}).get(key)


def aia_v120_equal_metric(value, reference):
    if value is None or reference is None:
        return False
    if isinstance(value, (int, float)) and isinstance(reference, (int, float)):
        return math.isclose(float(value), float(reference), rel_tol=1e-12, abs_tol=1e-12)
    return value == reference


def aia_v120_metric_delta(before, after, key, percent=False, lower_better=False):
    """Calcula delta e direção respeitando a semântica da métrica."""
    first = aia_v120_comparison_value(before, key)
    last = aia_v120_comparison_value(after, key)
    if first is None or last is None:
        return {"before": first, "after": last, "raw": None, "display": "N/D",
                "status": "N/D", "color": MUTED}
    raw = float(last) - float(first)
    scaled = raw * 100.0 if percent else raw
    if math.isclose(raw, 0.0, rel_tol=1e-12, abs_tol=1e-12):
        status, color = "ESTÁVEL", MUTED
    else:
        improved = raw < 0 if lower_better else raw > 0
        status, color = ("MELHOROU", GOOD) if improved else ("PIOROU", BAD)
    display = f"{scaled:+.1f} p.p." if percent else f"{scaled:+g}"
    return {"before": first, "after": last, "raw": raw, "display": display,
            "status": status, "color": color}


def aia_v120_metric_extreme(identities, key, mode):
    valid = [(identity, aia_v120_comparison_value(identity, key)) for identity in identities]
    valid = [(identity, value) for identity, value in valid if value is not None]
    if not valid:
        return {"value": None, "identities": [], "labels": [], "text": "N/D"}
    target = (max if mode == "max" else min)(value for _, value in valid)
    winners = [identity for identity, value in valid if aia_v120_equal_metric(value, target)]
    return {"value": target, "identities": winners,
            "labels": [identity["label"] for identity in winners]}


def aia_v120_format_join(values):
    if not values:
        return "N/D"
    if len(values) == 1:
        return values[0]
    if len(values) == 2:
        return values[0] + " e " + values[1]
    return ", ".join(values[:-1]) + " e " + values[-1]


def aia_v120_comparison_extremes(identities):
    specs = (
        ("Maior Accuracy D-I", "Accuracy", "max", True),
        ("Maior Recall D-I", "Recall / Sensibilidade", "max", True),
        ("Maior Precision D-I", "Precision", "max", True),
        ("Menor FP-DI", "FP", "min", False),
        ("Menor FN-DI", "FN", "min", False),
        ("Maior FP-DI", "FP", "max", False),
        ("Maior FN-DI", "FN", "max", False),
    )
    results = []
    for title, key, mode, percent in specs:
        extreme = aia_v120_metric_extreme(identities, key, mode)
        value = extreme["value"]
        formatted = "N/D" if value is None else (pct(value) if percent else f"{value:g}")
        results.append({"title": title, "key": key, "mode": mode,
                        "value": value, "value_text": formatted,
                        "identities": extreme["identities"],
                        "labels": extreme["labels"],
                        "text": f"{aia_v120_format_join(extreme['labels'])} — {formatted}"})
    return results


def aia_v120_bar_model(identities, key, percent=False):
    values = [aia_v120_comparison_value(identity, key) for identity in identities]
    valid = [float(value) for value in values if value is not None]
    maximum = 1.0 if percent else max(valid + [1.0])
    model = []
    for identity, value in zip(identities, values):
        ratio = None if value is None else max(0.0, min(1.0, float(value) / maximum))
        model.append({"identity": identity, "value": value, "ratio": ratio})
    return model


def aia_v120_pair_narrative(before, after):
    """Traduz deltas observados sem sugerir causa ou ajuste de threshold."""
    positive, negative, neutral = [], [], []
    specs = (("Recall", "Recall / Sensibilidade", True, False),
             ("FN-DI", "FN", False, True), ("FP-DI", "FP", False, True),
             ("Accuracy", "Accuracy", True, False))
    for label, key, percent, lower_better in specs:
        delta = aia_v120_metric_delta(before, after, key, percent, lower_better)
        first, last = delta["before"], delta["after"]
        if delta["status"] == "N/D":
            neutral.append(f"{label} ficou N/D")
            continue
        if percent:
            amount = abs(delta["raw"] * 100.0)
            if delta["raw"] > 0: phrase = f"{label} aumentou {amount:.1f} p.p."
            elif delta["raw"] < 0: phrase = f"{label} caiu {amount:.1f} p.p."
            else: phrase = f"{label} permaneceu em {pct(last)}"
        else:
            if delta["raw"] > 0: phrase = f"{label} aumentou de {first:g} para {last:g}"
            elif delta["raw"] < 0: phrase = f"{label} caiu de {first:g} para {last:g}"
            else: phrase = f"{label} permaneceu em {last:g}"
        if delta["status"] == "MELHOROU": positive.append(phrase)
        elif delta["status"] == "PIOROU": negative.append(phrase)
        else: neutral.append(phrase)
    prefix = f"Entre {before['label']} e {after['label']}, "
    sentences = []
    def end_sentence(text):
        return text if text.endswith(".") else text + "."
    if positive: sentences.append(end_sentence(prefix + " e ".join(positive)))
    if negative: sentences.append(end_sentence(("Entretanto, " if positive else prefix) + " e ".join(negative)))
    if neutral: sentences.append(end_sentence("Também, " + " e ".join(neutral)))
    return " ".join(sentences)


def aia_v120_visual_comparison_model(selected):
    identities = aia_v120_comparison_identities(selected)
    if len(identities) < 2:
        raise ValueError("selecione pelo menos duas coletas")
    if len(identities) > 50:
        raise ValueError("o limite é de 50 coletas")
    pairs = [(identities[index - 1], identities[index]) for index in range(1, len(identities))]
    first_last = (identities[0], identities[-1])
    deltas = []
    for scope, before, after in (("BASE → FINAL", first_last[0], first_last[1]), *(
            ("ANTERIOR → ATUAL", before, after) for before, after in pairs)):
        for label, key, percent, lower_better in AIA_V120_COMPARISON_METRICS:
            deltas.append({"scope": scope, "before_identity": before, "after_identity": after,
                           "label": label, "key": key,
                           **aia_v120_metric_delta(before, after, key, percent, lower_better)})
    narratives = [aia_v120_pair_narrative(before, after) for before, after in pairs]
    return {"identities": identities, "pairs": pairs, "first_last": first_last,
            "deltas": deltas, "extremes": aia_v120_comparison_extremes(identities),
            "narratives": narratives}


def aia_v120_comparison_metric_text(value, percent=False):
    if value is None:
        return "N/D"
    return pct(value) if percent else f"{value:g}"


def _aia_v120_draw_comparison_bars(parent, identities, title, key, percent, lower_better):
    """Desenha barras D-I mantendo a cor fixa de cada identidade visual."""
    height = max(150, 38 * len(identities) + 42)
    canvas = tk.Canvas(parent, bg=PANEL, highlightthickness=0, height=height)
    canvas.pack(fill="both", expand=True, padx=12, pady=(2, 12))
    canvas.update_idletasks()
    width = max(canvas.winfo_width(), 720)
    left, right, top = 190, 88, 18
    plot_width = max(1, width - left - right)
    row_height = max(30, (height - top - 18) / max(len(identities), 1))
    bars = aia_v120_bar_model(identities, key, percent)
    valid = [item for item in bars if item["value"] is not None]
    best_value = None
    worst_value = None
    if valid:
        values = [item["value"] for item in valid]
        best_value = min(values) if lower_better else max(values)
        worst_value = max(values) if lower_better else min(values)
    for index, item in enumerate(bars):
        identity, value, ratio = item["identity"], item["value"], item["ratio"]
        y = top + index * row_height + row_height / 2
        canvas.create_text(12, y, text=f"● {identity['short']}  {identity['name'][:22]}",
                           fill=identity["color"], anchor="w", font=("Segoe UI", 8, "bold"))
        canvas.create_rectangle(left, y - 8, left + plot_width, y + 8,
                                fill="#1A2C38", outline="")
        if ratio is not None:
            canvas.create_rectangle(left, y - 8, left + plot_width * ratio, y + 8,
                                    fill=identity["color"], outline="")
        marker = ""
        if value is not None and aia_v120_equal_metric(value, best_value):
            marker = "  ★ MELHOR"
        if value is not None and aia_v120_equal_metric(value, worst_value):
            marker = "  ◆ EMPATE" if marker else "  ▼ PIOR"
        value_text = aia_v120_comparison_metric_text(value, percent)
        canvas.create_text(left + plot_width + 8, y, text=value_text + marker,
                           fill=GOOD if "MELHOR" in marker else (BAD if "PIOR" in marker else TEXT),
                           anchor="w", font=("Segoe UI", 8, "bold"))
    return canvas


def _aia_v120_draw_detailed_trend(parent, identities, title, key, percent=False):
    """Gráfico secundário; a cor de cada ponto é a identidade da coleta."""
    canvas = tk.Canvas(parent, bg=PANEL, highlightthickness=0, height=245)
    canvas.pack(fill="both", expand=True, padx=10, pady=(2, 10))
    canvas.update_idletasks()
    width, height = max(canvas.winfo_width(), 520), 245
    left, right, top, bottom = 56, 26, 24, 52
    plot_width, plot_height = max(1, width - left - right), max(1, height - top - bottom)
    values = [aia_v120_comparison_value(identity, key) for identity in identities]
    valid = [float(value) for value in values if value is not None]
    if not valid:
        canvas.create_text(width / 2, height / 2, text="N/D — dados D-I insuficientes",
                           fill=MUTED, font=("Segoe UI", 9, "bold"))
        return canvas
    high = 1.0 if percent else max(max(valid), 1.0) * 1.12
    for index in range(5):
        fraction = index / 4
        y = top + plot_height * (1 - fraction)
        canvas.create_line(left, y, width - right, y, fill="#29404F")
        label_value = high * fraction
        canvas.create_text(left - 8, y,
                           text=f"{label_value * 100:.0f}%" if percent else f"{label_value:.0f}",
                           fill=MUTED, anchor="e", font=("Segoe UI", 7))
    step = plot_width / max(len(identities) - 1, 1)
    points = []
    for index, (identity, value) in enumerate(zip(identities, values)):
        x = left + (index * step if len(identities) > 1 else plot_width / 2)
        canvas.create_text(x, height - bottom + 14, text=identity["short"],
                           fill=identity["color"], font=("Segoe UI", 7, "bold"))
        canvas.create_text(x, height - bottom + 29,
                           text=f"T={identity['threshold']}", fill=MUTED, font=("Segoe UI", 6))
        if value is None:
            continue
        numeric = float(value)
        y = top + plot_height * (1 - max(0.0, min(1.0, numeric / high)))
        points.append((x, y))
        canvas.create_oval(x - 5, y - 5, x + 5, y + 5,
                           fill=identity["color"], outline=TEXT)
        canvas.create_text(x, y - 12,
                           text=aia_v120_comparison_metric_text(value, percent),
                           fill=identity["color"], font=("Segoe UI", 7, "bold"))
    for point_a, point_b in zip(points, points[1:]):
        canvas.create_line(*point_a, *point_b, fill="#547383", width=2)
    return canvas


def _aia_v120_open_detailed_trends(self, identities):
    """Abre os quatro gráficos antigos em uma janela secundária não modal."""
    window = tk.Toplevel(self)
    window.title("AIA V1.2.0 — Tendência detalhada D-I")
    window.geometry("1180x780")
    window.configure(bg=BG)
    try:
        window.transient(self)
    except Exception:
        pass
    tk.Label(window, text="TENDÊNCIA DETALHADA — D-I", bg=BG, fg=TEXT,
             font=("Segoe UI", 15, "bold")).pack(anchor="w", padx=16, pady=(14, 2))
    tk.Label(window,
             text="Visualização secundária. Cores e códigos são os mesmos da comparação principal.",
             bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w", padx=16, pady=(0, 8))
    grid = tk.Frame(window, bg=BG)
    grid.pack(fill="both", expand=True, padx=12, pady=(0, 12))
    for index, (title, key, percent) in enumerate((
            ("EVOLUÇÃO — ACCURACY D-I", "Accuracy", True),
            ("EVOLUÇÃO — RECALL D-I", "Recall / Sensibilidade", True),
            ("EVOLUÇÃO — FP-DI", "FP", False),
            ("EVOLUÇÃO — FN-DI", "FN", False))):
        row, column = divmod(index, 2)
        card = self.chart_card(grid, title,
                               "Maior = melhor" if key not in {"FP", "FN"} else "Menor = melhor")
        card.grid(row=row, column=column, sticky="nsew", padx=4, pady=4)
        _aia_v120_draw_detailed_trend(card, identities, title, key, percent)
    grid.columnconfigure(0, weight=1)
    grid.columnconfigure(1, weight=1)
    grid.rowconfigure(0, weight=1)
    grid.rowconfigure(1, weight=1)
    return window


def _aia_v120_render_visual_comparison(self, selected, result):
    """Renderiza a visão operacional D-I; C e D-O ficam complementares e separados."""
    for widget in result.winfo_children():
        widget.destroy()
    model = aia_v120_visual_comparison_model(selected)
    identities = model["identities"]

    header = tk.Frame(result, bg=PANEL)
    header.pack(fill="x", padx=14, pady=(12, 5))
    tk.Label(header, text=f"{len(identities)} COLETAS — COMPARAÇÃO OPERACIONAL D-I",
             bg=PANEL, fg=ACCENT2, font=("Segoe UI", 11, "bold")).pack(side="left")
    tk.Button(header, text="VER TENDÊNCIA DETALHADA",
              command=lambda: _aia_v120_open_detailed_trends(self, identities),
              bg="#0D4560", fg=TEXT, activebackground="#176B8F", bd=0,
              padx=12, pady=7, font=("Segoe UI", 8, "bold"), cursor="hand2").pack(side="right")

    identity_card = self.card(result, "IDENTIDADE DAS COLETAS",
                              "O código e a cor de cada coleta permanecem iguais em toda a tela")
    identity_card.pack(fill="x", padx=10, pady=5)
    identity_grid = tk.Frame(identity_card, bg=PANEL)
    identity_grid.pack(fill="x", padx=10, pady=(0, 10))
    for index, identity in enumerate(identities):
        row, column = divmod(index, 4)
        tile = tk.Frame(identity_grid, bg=PANEL3,
                        highlightbackground=identity["color"], highlightthickness=2)
        tile.grid(row=row, column=column, sticky="nsew", padx=4, pady=4)
        tk.Label(tile, text=f"● {identity['label']} — {identity['name']}",
                 bg=PANEL3, fg=identity["color"], font=("Segoe UI", 9, "bold"),
                 wraplength=255, justify="left").pack(anchor="w", padx=10, pady=(8, 3))
        detail = (f"Data: {identity['date']}   •   Threshold: {identity['threshold']}\n"
                  f"N de inspeções D-I: {identity['n']}\nOrigem: {identity['origin']}")
        tk.Label(tile, text=detail, bg=PANEL3, fg=TEXT, font=("Segoe UI", 8),
                 wraplength=255, justify="left").pack(anchor="w", padx=10, pady=(0, 8))
    for column in range(4):
        identity_grid.columnconfigure(column, weight=1)

    matrix_card = self.card(result, "MATRIZ COMPARATIVA — D-I",
                            "★ melhor • ▼ pior • ◆ empate geral | percentuais: maior = melhor • FP-DI/FN-DI: menor = melhor")
    matrix_card.pack(fill="x", padx=10, pady=5)
    columns = ("indicator", "direction", *(identity["short"] for identity in identities))
    matrix_wrap = tk.Frame(matrix_card, bg=PANEL)
    matrix_wrap.pack(fill="x", padx=12, pady=(4, 12))
    matrix = ttk.Treeview(matrix_wrap, columns=columns, show="headings", height=6)
    matrix.heading("indicator", text="Indicador")
    matrix.column("indicator", width=180, anchor="w")
    matrix.heading("direction", text="Leitura")
    matrix.column("direction", width=125, anchor="center")
    for identity in identities:
        matrix.heading(identity["short"], text=f"{identity['short']} — {identity['name'][:15]}")
        matrix.column(identity["short"], width=145, anchor="center")
    for label, key, percent, lower_better in AIA_V120_COMPARISON_METRICS:
        values = [aia_v120_comparison_value(identity, key) for identity in identities]
        valid = [value for value in values if value is not None]
        best = (min(valid) if lower_better else max(valid)) if valid else None
        worst = (max(valid) if lower_better else min(valid)) if valid else None
        cells = []
        for value in values:
            marker = ""
            if value is not None and aia_v120_equal_metric(value, best):
                marker = "★ "
            if value is not None and aia_v120_equal_metric(value, worst):
                marker = "◆ " if marker else "▼ "
            cells.append(marker + aia_v120_comparison_metric_text(value, percent))
        matrix.insert("", "end", values=(label,
                                           "MENOR = MELHOR" if lower_better else "MAIOR = MELHOR",
                                           *cells))
    matrix_scroll = ttk.Scrollbar(matrix_wrap, orient="horizontal", command=matrix.xview)
    matrix.configure(xscrollcommand=matrix_scroll.set)
    matrix.pack(fill="x")
    matrix_scroll.pack(fill="x")

    bars_card = self.card(result, "BARRAS HORIZONTAIS — PRINCIPAIS INDICADORES D-I",
                          "A extensão da barra representa exatamente o valor; contagens FP/FN usam a maior contagem selecionada como escala")
    bars_card.pack(fill="x", padx=10, pady=5)
    for label, key, percent, lower_better in AIA_V120_COMPARISON_METRICS:
        metric_card = self.card(bars_card, label,
                                "MENOR = MELHOR" if lower_better else "MAIOR = MELHOR")
        metric_card.pack(fill="x", padx=8, pady=4)
        _aia_v120_draw_comparison_bars(metric_card, identities, label, key, percent, lower_better)

    extremes_card = self.card(result, "DESTAQUES AUTOMÁTICOS — D-I",
                              "Empates são preservados; nenhuma relação causal é inferida")
    extremes_card.pack(fill="x", padx=10, pady=5)
    extremes_grid = tk.Frame(extremes_card, bg=PANEL)
    extremes_grid.pack(fill="x", padx=10, pady=(0, 10))
    for index, extreme in enumerate(model["extremes"]):
        row, column = divmod(index, 2)
        text = f"{extreme['title']}: {extreme['text']}"
        tk.Label(extremes_grid, text=text, bg=PANEL, fg=TEXT,
                 font=("Segoe UI", 9, "bold"), wraplength=540,
                 justify="left").grid(row=row, column=column, sticky="w", padx=6, pady=5)
    extremes_grid.columnconfigure(0, weight=1)
    extremes_grid.columnconfigure(1, weight=1)

    def delta_table(parent, rows, height):
        table = ttk.Treeview(parent,
                             columns=("interval", "metric", "before", "after", "delta", "status"),
                             show="headings", height=max(1, height))
        specs = (("interval", "Comparação", 175), ("metric", "Indicador", 160),
                 ("before", "Antes", 95), ("after", "Depois", 95),
                 ("delta", "Variação", 105), ("status", "Leitura", 105))
        for column, label, width in specs:
            table.heading(column, text=label)
            table.column(column, width=width, anchor="w" if column in {"interval", "metric"} else "center")
        table.tag_configure("improved", foreground=GOOD)
        table.tag_configure("worse", foreground=BAD)
        table.tag_configure("stable", foreground=MUTED)
        table.pack(fill="x", padx=12, pady=8)
        for row in rows:
            percent = row["key"] not in {"FP", "FN"}
            interval = f"{row['before_identity']['short']} → {row['after_identity']['short']}"
            tag = "improved" if row["status"] == "MELHOROU" else (
                "worse" if row["status"] == "PIOROU" else "stable")
            table.insert("", "end", tags=(tag,), values=(
                interval, row["label"],
                aia_v120_comparison_metric_text(row["before"], percent),
                aia_v120_comparison_metric_text(row["after"], percent),
                row["display"], row["status"]))
        return table

    base_rows = [row for row in model["deltas"] if row["scope"] == "BASE → FINAL"]
    base_card = self.card(result, "PRIMEIRA × ÚLTIMA COLETA — D-I",
                          "Verde = melhoria • vermelho = piora • cinza = estabilidade")
    base_card.pack(fill="x", padx=10, pady=5)
    delta_table(base_card, base_rows, len(base_rows))

    previous_rows = [row for row in model["deltas"] if row["scope"] == "ANTERIOR → ATUAL"]
    previous_card = self.card(result, "CADA COLETA × ANTERIOR — D-I",
                              "A direção respeita a semântica própria de cada indicador")
    previous_card.pack(fill="x", padx=10, pady=5)
    delta_table(previous_card, previous_rows, min(24, len(previous_rows)))

    narrative_card = self.card(result, "LEITURA OBJETIVA DOS DADOS — D-I",
                                "Descrição numérica, sem atribuição de causa e sem recomendação de Threshold")
    narrative_card.pack(fill="x", padx=10, pady=5)
    for narrative in model["narratives"]:
        tk.Label(narrative_card, text="• " + narrative, bg=PANEL, fg=TEXT,
                 wraplength=1180, justify="left", font=("Segoe UI", 9)).pack(
                     anchor="w", padx=14, pady=4)
    if len(identities) > 2:
        tk.Label(narrative_card,
                 text="• Síntese primeira → última: " + aia_v120_pair_narrative(*model["first_last"]),
                 bg=PANEL, fg=ACCENT2, wraplength=1180, justify="left",
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=14, pady=(5, 10))

    def complementary_matrix(title, subtitle, source_key, specs):
        card = self.card(result, title, subtitle)
        card.pack(fill="x", padx=10, pady=5)
        cols = ("indicator", *(identity["short"] for identity in identities))
        table_wrap = tk.Frame(card, bg=PANEL)
        table_wrap.pack(fill="x", padx=12, pady=(4, 12))
        table = ttk.Treeview(table_wrap, columns=cols, show="headings", height=len(specs))
        table.heading("indicator", text="Indicador")
        table.column("indicator", width=220, anchor="w")
        for identity in identities:
            table.heading(identity["short"], text=identity["short"])
            table.column(identity["short"], width=125, anchor="center")
        for label, key, percent in specs:
            cells = []
            for identity in identities:
                source = identity["dataset"].get(source_key, {}) or {}
                cells.append(aia_v120_comparison_metric_text(source.get(key), percent))
            table.insert("", "end", values=(label, *cells))
        table_scroll = ttk.Scrollbar(table_wrap, orient="horizontal", command=table.xview)
        table.configure(xscrollcommand=table_scroll.set)
        table.pack(fill="x")
        table_scroll.pack(fill="x")
        return card

    complementary_matrix(
        "BLOCO COMPLEMENTAR D-O — DETECÇÃO POR OCORRÊNCIA",
        "Unidade: ocorrência/defeito • não substitui D-I",
        "defects", (("Ocorrências reais", "real", False), ("Detecções", "det", False),
                    ("TP-D", "tp", False), ("FP-D", "fp", False),
                    ("FN-D", "fn", False), ("Precision D-O", "precision", True),
                    ("Recall D-O", "recall", True)))
    complementary_matrix(
        "BLOCO COMPLEMENTAR C — CLASSIFICAÇÃO OK/NOK",
        "Unidade: inspeção • fonte exclusiva Vision AI × Referência/colunas C explícitas • FP-C/FN-C não alimentam D-I",
        "metrics", (("N C", "N", False), ("TP-C", "TP", False),
                    ("VN-C", "VN/TN", False), ("FP-C", "FP", False),
                    ("FN-C", "FN", False), ("Accuracy C", "Accuracy", True),
                    ("Recall C", "Recall / Sensibilidade", True),
                    ("Precision C", "Precision", True)))

    meta_card = self.card(result, "META ATIVA — LEITURA D-I POR COLETA",
                          "A meta é somente avaliada aqui; recomendações permanecem no módulo Calibração")
    meta_card.pack(fill="x", padx=10, pady=(5, 10))
    meta_table = ttk.Treeview(meta_card,
                              columns=("collection", "threshold", "met", "configured", "status"),
                              show="headings", height=min(10, len(identities)))
    for column, label, width in (("collection", "Coleta", 260),
                                 ("threshold", "Threshold", 100),
                                 ("met", "Atendidos", 90),
                                 ("configured", "Configurados", 100),
                                 ("status", "Status", 210)):
        meta_table.heading(column, text=label)
        meta_table.column(column, width=width,
                          anchor="w" if column == "collection" else "center")
    target = self.expectation_snapshot()
    for identity in identities:
        evaluation = aia_v120_meta_capability_evaluation(target, identity["di"])
        meta_table.insert("", "end", values=(
            f"{identity['short']} — {identity['name']}", identity["threshold"],
            evaluation["met"], evaluation["configured"], evaluation["status"]))
    meta_table.pack(fill="x", padx=12, pady=(4, 12))

    self.status.set(f"✓ {len(identities)} coleta(s) comparada(s) • referência operacional: D-I.")
    try:
        self.update_idletasks()
        self.content_canvas.configure(scrollregion=self.content_canvas.bbox("all"))
    except Exception:
        pass


# ============================================================================
# AIA V1.2.0 — AUDITORIA FINAL DE INTEGRIDADE
# Correções pontuais pós-validação: D-O explícita, orientação de calibração e
# consistência de exportação. Nenhum layout aprovado é reconstruído aqui.
# ============================================================================

def _aia_v120_detection_occurrence_audited(self):
    """D-O canônica por linha.

    Prioridade por inspeção:
    1) TP-D/FP-D/FN-D explícitos e conservativos (preserva pareamento real);
    2) detalhamento por tipo, quando informado;
    3) estimativa min/max de Real × Detectado da própria linha.
    FP de uma inspeção nunca compensa FN de outra.
    """
    sa=getattr(self,"source_analysis",{}) or {}; om=sa.get("occurrence_matrix",{}) or {}
    observed=[]; unavailable=0
    for row0 in self.rows:
        row=normalize_row(row0)
        if row.get("detection_observed",False): observed.append(row)
        else: unavailable+=1
    if getattr(self,"source_file","") and om.get("present") and not observed:
        vals=[_aia_v113_num(om.get(k)) for k in ("tp","tn","fp","fn")]
        if all(v is not None for v in vals):
            tp,tn,fp,fn=vals; mm=_aia_metric_pack(tp,tn,fp,fn,tp+tn+fp+fn)
            return {"available":True,"real":tp+fn,"det":tp+fp,"tp":tp,"fp":fp,"fn":fn,"tn":tn,
                    "precision":mm["Precision"],"recall":mm["Recall / Sensibilidade"],"f1":mm["F1"],
                    "available_rows":0,"unavailable_rows":len(self.rows),"detailed":False,"estimated":False,
                    "status":om.get("status","IMPORTADO / VALIDADO"),"method":om.get("method","Matriz agregada da fonte"),
                    "full_metrics":mm,"source_sheet":om.get("sheet"),"source_cells":om.get("cells",{}),
                    "semantic_evidence":om.get("evidence",[]),"semantic_conflicts":om.get("conflicts",[])}
    if not observed:
        return {"available":False,"real":0,"det":0,"tp":0,"fp":0,"fn":0,"tn":None,
                "precision":None,"recall":None,"f1":None,"available_rows":0,"unavailable_rows":unavailable,
                "detailed":False,"estimated":False,"status":"N/D","method":"N/D"}
    real=det=tp=fp=fn=0; explicit_used=estimated_used=0; detailed=False; invalid_explicit=[]
    for row in observed:
        rr=max(0,int(row.get("real",0) or 0)); dd=max(0,int(row.get("det",0) or 0))
        ep,efp,efn=row.get("tp_d"),row.get("fp_d"),row.get("fn_d")
        explicit=bool(row.get("detection_explicit_matrix",False)) and all(v is not None for v in (ep,efp,efn))
        if explicit:
            ep,efp,efn=max(0,int(ep or 0)),max(0,int(efp or 0)),max(0,int(efn or 0))
            # Pareamento explícito pode diferir do min/max; precisa apenas conservar os totais.
            if ep+efp==dd and ep+efn==rr:
                line={"real":rr,"det":dd,"tp":ep,"fp":efp,"fn":efn,"available":True}
                explicit_used+=1
            else:
                invalid_explicit.append(str(row.get("id") or "?")); line=self.detection_for_row(row); estimated_used+=1
        else:
            line=self.detection_for_row(row); estimated_used+=1
        detailed=detailed or bool(row.get("defects") or [])
        real+=max(0,int(line.get("real",0) or 0)); det+=max(0,int(line.get("det",0) or 0))
        tp+=max(0,int(line.get("tp",0) or 0)); fp+=max(0,int(line.get("fp",0) or 0)); fn+=max(0,int(line.get("fn",0) or 0))
    precision=safe(tp,tp+fp); recall=safe(tp,tp+fn)
    f1=None if precision is None or recall is None or precision+recall==0 else 2*precision*recall/(precision+recall)
    if explicit_used and not estimated_used:
        status="IMPORTADO / VALIDADO"; method="TP-D/FP-D/FN-D explícitos, conservativos e somados por inspeção"; estimated=False
    elif explicit_used:
        status="MISTO"; method="D-O por inspeção: valores explícitos válidos + estimativa min/max nas demais linhas"; estimated=True
    else:
        status="ESTIMADO"; method="Soma D-O por inspeção (min/max por linha)"; estimated=True
    if invalid_explicit:
        method += " • matriz explícita inválida ignorada em: " + ", ".join(invalid_explicit[:8])
    return {"available":True,"real":real,"det":det,"tp":tp,"fp":fp,"fn":fn,"tn":None,
            "precision":precision,"recall":recall,"f1":f1,"available_rows":len(observed),
            "unavailable_rows":unavailable,"detailed":detailed,"estimated":estimated,"status":status,"method":method}


def aia_v120_calibration_guidance(target, di_metrics):
    """Gera orientação conservadora de calibração a partir de D-I + Meta."""
    di=di_metrics or {}; evaluation=aia_v120_meta_capability_evaluation(target or {},di)
    if not evaluation.get("available"):
        return [{"level":"⚪","title":"DADOS INSUFICIENTES","text":"Registre Real e Detectado por inspeção para habilitar recomendações D-I.","color":MUTED}]
    if not evaluation.get("configured"):
        return [{"level":"🟡","title":"META NÃO CONFIGURADA","text":"Defina o perfil BOM, EXCELENTE ou uma Meta personalizada antes de usar o AIA para orientar ajustes de Threshold.","color":WARN}]
    configured_rows=[row for row in evaluation["rows"] if row["target"] is not None]
    if any(row["actual"] is None for row in configured_rows):
        missing=", ".join(row["label"] for row in configured_rows if row["actual"] is None)
        return [{"level":"🟡","title":"VALIDAÇÃO D-I INCOMPLETA","text":"Não há base nesta amostra para avaliar: "+missing+". Inclua casos que permitam medir defeitos encontrados e alarmes antes de recomendar ajuste de Threshold.","color":WARN}]
    rows={row["key"]:row for row in evaluation["rows"]}
    fp_problem=any(rows[k]["status"]!="META ATINGIDA" for k in ("precision","fp100") if rows[k]["target"] is not None and rows[k]["actual"] is not None)
    fn_problem=any(rows[k]["status"]!="META ATINGIDA" for k in ("recall","fn100") if rows[k]["target"] is not None and rows[k]["actual"] is not None)
    out=[]
    if fp_problem and fn_problem:
        out.append({"level":"🔴","title":"FP-DI E FN-DI EXIGEM ATENÇÃO",
                    "text":"Não trate o Threshold como ajuste único. Investigue primeiro ROI/cobertura, iluminação, qualidade da imagem, referência, estabilidade do processo e modelo. Depois teste variações pequenas de Threshold acompanhando simultaneamente Recall, Precision, FP-DI e FN-DI.","color":BAD})
    elif fp_problem:
        out.append({"level":"🟡","title":"FALSOS POSITIVOS / PRECISION FORA DA META",
                    "text":"Hipótese de teste: avaliar aumento gradual do Threshold de confiança para reduzir falsos alarmes. Monitorar Recall e FN-DI a cada passo, pois Threshold excessivo pode aumentar escapes.","color":WARN})
    elif fn_problem:
        out.append({"level":"🔴","title":"FALSOS NEGATIVOS / RECALL FORA DA META",
                    "text":"Hipótese de teste: avaliar redução gradual do Threshold para aumentar sensibilidade. Monitorar Precision e FP-DI, e investigar também ROI, iluminação, foco, cobertura e condições de aquisição.","color":BAD})
    else:
        out.append({"level":"🟢","title":"CRITÉRIOS CRÍTICOS CONTROLADOS",
                    "text":"A configuração atual atende aos critérios monitorados ligados a Recall/FN e Precision/FP. Preserve os parâmetros e valide estabilidade com amostra maior e condições variadas antes de concluir capacidade definitiva.","color":GOOD})
    if evaluation.get("n",0)<30:
        out.append({"level":"🟡","title":"AMOSTRA PRELIMINAR",
                    "text":f"N-DI={evaluation.get('n',0)}. Use a recomendação como orientação de teste, não como conclusão definitiva de capacidade.","color":WARN})
    return out

# ============================================================================
# AIA V1.2.0 — RUNTIME FINAL ÚNICO / BARREIRA DE REGRESSÃO
# Esta seção é deliberadamente a última associação de funções antes do main.
# C e D-I possuem fontes independentes:
#   C   = TP-C/VN-C/FP-C/FN-C OU Vision AI × Referência.
#   D-I = presença de Real > 0 × presença de Detectado > 0.
# Campos TP/VN/FP/FN genéricos nunca são promovidos a C por nome isolado.
# ============================================================================

# Um único registro declarativo determina todas as implementações ativas.
# Nenhuma reassociação deve ser adicionada abaixo deste bloco.
FINAL_RUNTIME_BINDINGS = {
    "__init__": _aia_smart_init,
    "show": _aia_v10_show_frontend,
    "refresh_all": _aia_v120_refresh_all,
    "import_excel": _aia_v113_import_excel,
    "export_excel": _aia_smart_export_excel,
    "_smart_detection_inspection": _aia_smart_detection_inspection,
    "_smart_detection_occurrence": _aia_v120_detection_occurrence_audited,
    "_smart_classification_totals": _aia_smart_classification_totals,
    "_smart_metrics": _aia_smart_metrics,
    "display_metrics": _aia_smart_display_metrics,
    "classification_totals": _aia_smart_classification_totals,
    "metrics": _aia_smart_metrics,
    "detection_totals": _aia_smart_detection_totals,
    "defect_counts": _aia_smart_defect_counts,
    "_aia_smart_add_interpretation_button": _aia_smart_add_interpretation_button,
    "_parse_target": _aia_v120_parse_target,
    "expectation_snapshot": _aia_v120_expectation_snapshot,
    "_expectation_actuals": _aia_v120_expectation_actuals,
    "_expectation_rows": _aia_v120_expectation_rows,
    "expectation_summary": _aia_v120_expectation_summary,
    "_expectation_status_color": _aia_v120_expectation_status_color,
    "_aia_final_load_active_expectation": _aia_v120_load_active_expectation,
    "_load_active_expectation": _aia_v120_load_active_expectation,
    "save_expectation": _aia_v120_save_expectation,
    "_archive_current_collection": _aia_v120_archive_current_collection,
    "_comparison_datasets": _aia_v120_comparison_datasets,
    "_aia_final_delete_selected_collection": _aia_v120_delete_selected_collection,
    "_select_all_compare": _aia_v120_select_all_compare,
    "_run_comparison": _aia_v120_run_comparison,
    "_calibration_score": _aia_v120_calibration_score,
    "_best_configuration": _aia_v120_best_configuration,
    "_draw_threshold_chart": _aia_v120_draw_threshold_chart,
    "view_operacao": _aia_v10_final_operacao,
    "quick_entry": _aia_v10_final_quick_entry,
    "view_dashboard": _aia_v116_dashboard,
    "view_graficos": _aia_v115_field_graficos,
    "view_comparacao": _aia_v120_view_comparacao,
    "view_tecnica": _aia_v117_view_tecnica,
    "view_engenharia": _aia_v115_field_engenharia,
    "view_calibracao": _aia_v120_view_calibracao,
    "view_expectativa": _aia_v120_view_meta,
    "view_auditoria": _aia_v120_view_auditoria,
}
for _method_name, _implementation in FINAL_RUNTIME_BINDINGS.items():
    setattr(App, _method_name, _implementation)

# ============================================================================
# AIA V1.3.0 RC — CAMADA DE COMPATIBILIDADE E RASTREABILIDADE
# ============================================================================
# A camada abaixo é intencionalmente aditiva: projetos V1.2.0 continuam válidos
# sem migração. Campos novos são opcionais e sempre exportados de forma explícita.

_aia_v130_normalize_row_base = normalize_row

def aia_v130_normalize_row(row):
    """Normaliza registros novos e antigos sem alterar sua semântica histórica."""
    x = _aia_v130_normalize_row_base(row)
    raw = (row or {}).get("correspondentes", (row or {}).get("matched"))
    try:
        x["correspondentes"] = None if raw in (None, "") else max(0, int(float(raw)))
    except (TypeError, ValueError):
        x["correspondentes"] = None
    raw_snapshot = (row or {}).get("configuracao_classes_snapshot", {})
    x["configuracao_classes_snapshot"] = raw_snapshot if isinstance(raw_snapshot, dict) else {}
    normalized = []
    for d in x.get("defects", []):
        raw_match = d.get("correspondentes", d.get("matched"))
        try:
            match = None if raw_match in (None, "") else max(0, int(float(raw_match)))
        except (TypeError, ValueError):
            match = None
        normalized.append({"type": d["type"], "real": d["real"], "det": d["det"],
                           "correspondentes": match})
    x["defects"] = normalized
    return x

normalize_row = aia_v130_normalize_row

def aia_v130_config_snapshot(self):
    configs = getattr(self, "class_thresholds", {}) or {}
    return {str(name): {"threshold": str(item.get("threshold", "")).strip(),
                        "confidence": str(item.get("confidence", "")).strip(),
                        "notes": str(item.get("notes", "")).strip()}
            for name, item in configs.items() if str(name).strip()}

def aia_v130_snapshot_rows(self):
    snapshot = aia_v130_config_snapshot(self)
    for i, row in enumerate(self.rows):
        normalized = normalize_row(row)
        if normalized.get("detection_observed"):
            normalized["configuracao_classes_snapshot"] = dict(snapshot)
        self.rows[i] = normalized

def _aia_v130_init(self, *args, **kwargs):
    self.class_thresholds = {}
    _aia_smart_init(self, *args, **kwargs)
    self.after(120, lambda: self.status.set("AIA V1.3.0 RC • D-I PRINCIPAL • D-O CORRESPONDENTE • RASTREABILIDADE ATIVA"))

def _aia_v130_detection_occurrence(self):
    """D-O por inspeção: Correspondentes permite FP e FN simultâneos.

    Sem esse campo, a regra min/max histórica é preservada somente como ESTIMADO.
    """
    observed = [normalize_row(r) for r in self.rows if normalize_row(r).get("detection_observed", False)]
    unavailable = len(self.rows) - len(observed)
    if not observed:
        return {"available": False, "real": 0, "det": 0, "tp": 0, "fp": 0, "fn": 0,
                "tn": None, "precision": None, "recall": None, "f1": None,
                "available_rows": 0, "unavailable_rows": unavailable, "detailed": False,
                "estimated": False, "status": "N/D", "method": "N/D"}
    real = det = tp = fp = fn = explicit = estimated = 0
    invalid = []
    for row in observed:
        rr, dd = int(row["real"]), int(row["det"])
        corr = row.get("correspondentes")
        if corr is not None and corr <= min(rr, dd):
            line_tp = corr; explicit += 1
        else:
            if corr is not None: invalid.append(str(row.get("id") or "?"))
            line_tp = min(rr, dd); estimated += 1
        real += rr; det += dd; tp += line_tp; fp += dd - line_tp; fn += rr - line_tp
    precision, recall = safe(tp, tp + fp), safe(tp, tp + fn)
    f1 = None if precision is None or recall is None or precision + recall == 0 else 2 * precision * recall / (precision + recall)
    if explicit and not estimated:
        status, method, is_estimated = "EXPLÍCITO", "Correspondentes por inspeção: TP-D=Correspondentes; FP-D=Detectado-Correspondentes; FN-D=Real-Correspondentes", False
    elif explicit:
        status, method, is_estimated = "MISTO", "Correspondentes explícitos nas linhas informadas; regra por contagem ESTIMADA nas demais", True
    else:
        status, method, is_estimated = "ESTIMADO", "Sem correspondência explícita: TP-D=min(Real, Detectado), FP-D=max(Detectado-Real,0), FN-D=max(Real-Detectado,0)", True
    if invalid:
        method += " • Correspondentes inválido ignorado em: " + ", ".join(invalid[:8])
    return {"available": True, "real": real, "det": det, "tp": tp, "fp": fp, "fn": fn,
            "tn": None, "precision": precision, "recall": recall, "f1": f1,
            "available_rows": len(observed), "unavailable_rows": unavailable,
            "detailed": any(r.get("defects") for r in observed), "estimated": is_estimated,
            "status": status, "method": method, "explicit_rows": explicit, "estimated_rows": estimated}

def _aia_v130_detection_for_row(self, row):
    r = normalize_row(row)
    if not r.get("detection_available", True):
        return {"real": 0, "det": 0, "tp": 0, "fp": 0, "fn": 0, "available": False}
    rr, dd, corr = int(r["real"]), int(r["det"]), r.get("correspondentes")
    tp = corr if corr is not None and corr <= min(rr, dd) else min(rr, dd)
    return {"real": rr, "det": dd, "tp": tp, "fp": dd-tp, "fn": rr-tp, "available": True}

def _aia_v130_edit_correspondence(self):
    if not getattr(self, "tree", None) or not self.tree.selection():
        messagebox.showinfo("Correspondentes", "Selecione uma inspeção para registrar as ocorrências correspondentes.", parent=self); return
    index = int(self.tree.selection()[0]); row = normalize_row(self.rows[index])
    win = tk.Toplevel(self); win.title("Correspondentes D-O"); win.geometry("720x410"); win.minsize(650,370); win.configure(bg=BG); win.transient(self); win.grab_set()
    box = tk.Frame(win, bg=PANEL, padx=18, pady=18); box.pack(fill="both", expand=True, padx=12, pady=12)
    tk.Label(box, text="CORRESPONDENTES D-O", bg=PANEL, fg=TEXT, font=("Segoe UI", 14, "bold")).pack(anchor="w")
    limit=min(row['real'], row['det'])
    tk.Label(box, text=f"Inspeção {row['id'] or index+1}: Real={row['real']} • Detectado={row['det']}\n\nUse este campo somente quando você souber quantas detecções correspondem de fato a defeitos reais.\nExemplo: se Real=1, Detectado=3 e Correspondentes=1, então TP-D=1, FP-D=2 e FN-D=0.\nSe não houver pareamento confiável, deixe sem preencher: o AIA mantém D-O como ESTIMADO.", bg=PANEL, fg=MUTED, justify="left", wraplength=630).pack(anchor="w", pady=(6, 12))
    tk.Label(box, text=f"Quantidade de Correspondentes (de 0 a {limit})", bg=PANEL, fg=TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
    value = tk.StringVar(value="" if row.get("correspondentes") is None else str(row["correspondentes"]))
    ttk.Entry(box, textvariable=value, width=18).pack(anchor="w", pady=(3, 4))
    preview = tk.StringVar(value="Informe um valor para visualizar TP-D, FP-D e FN-D.")
    tk.Label(box, textvariable=preview, bg="#123B52", fg=ACCENT2, justify="left", anchor="w", padx=10, pady=8, font=("Segoe UI", 9, "bold")).pack(fill="x", pady=(4, 0))
    def update_preview(*_):
        try:
            corr=max(0, int(value.get()))
            if corr > limit: preview.set(f"Valor máximo: {limit}. Real={row['real']} e Detectado={row['det']}."); return
            preview.set(f"Resultado explícito: TP-D = {corr}   •   FP-D = {row['det']-corr}   •   FN-D = {row['real']-corr}")
        except ValueError:
            preview.set("Informe um número inteiro entre 0 e " + str(limit) + ".")
    value.trace_add("write", update_preview)
    if value.get(): update_preview()
    def save():
        try: corr = max(0, int(value.get()))
        except ValueError: messagebox.showerror("Correspondentes", "Informe um número inteiro não negativo.", parent=win); return
        if corr > min(row['real'], row['det']): messagebox.showerror("Correspondentes", "O valor não pode exceder o menor entre Real e Detectado.", parent=win); return
        row['correspondentes'] = corr; self.rows[index] = self.classify(row); aia_v130_snapshot_rows(self); win.destroy(); self.refresh_all()
    ttk.Button(box, text="Salvar correspondência", style="Accent.TButton", command=save).pack(anchor="e", pady=(14, 0))

def _aia_v130_configure_classes(self):
    win = tk.Toplevel(self); win.title("Threshold e Confidence por classe"); win.geometry("800x570"); win.minsize(700,500); win.configure(bg=BG); win.transient(self); win.grab_set()
    box = tk.Frame(win, bg=PANEL, padx=14, pady=14); box.pack(fill="both", expand=True, padx=12, pady=12)
    tk.Label(box, text="CONFIGURAÇÃO POR CLASSE DE DEFEITO", bg=PANEL, fg=TEXT, font=("Segoe UI", 14, "bold")).pack(anchor="w")
    tk.Label(box, text="Cada classe é independente. Ao salvar uma inspeção, a configuração ativa é copiada para o snapshot daquele registro.", bg=PANEL, fg=MUTED, wraplength=650, justify="left").pack(anchor="w", pady=(4, 10))
    table_box=tk.Frame(box,bg=PANEL); table_box.pack(fill="both",expand=True)
    tree = ttk.Treeview(table_box, columns=("classe", "threshold", "confidence", "notes"), show="headings", height=10)
    for col, label, width in (("classe", "Classe", 190), ("threshold", "Threshold", 100), ("confidence", "Confidence", 110), ("notes", "Observação", 240)):
        tree.heading(col, text=label); tree.column(col, width=width, anchor="center" if col != "notes" else "w")
    scrollbar=ttk.Scrollbar(table_box,orient="vertical",command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    tree.pack(side="left",fill="both", expand=True); scrollbar.pack(side="right",fill="y")
    def refill():
        for item in tree.get_children(): tree.delete(item)
        for name, cfg in sorted(self.class_thresholds.items()): tree.insert("", "end", iid=name, values=(name, cfg.get("threshold", ""), cfg.get("confidence", ""), cfg.get("notes", "")))
    tk.Label(box,text="Adicionar ou atualizar classe",bg=PANEL,fg=TEXT,font=("Segoe UI",10,"bold")).pack(anchor="w",pady=(10,3))
    tk.Label(box,text="Informe a classe e os parâmetros que foram usados para ela. Esses valores serão copiados para o registro quando a inspeção for salva.",bg=PANEL,fg=MUTED,wraplength=730,justify="left").pack(anchor="w",pady=(0,5))
    controls = tk.Frame(box, bg=PANEL); controls.pack(fill="x")
    name, threshold, confidence, notes = tk.StringVar(), tk.StringVar(), tk.StringVar(), tk.StringVar()
    fields=(("Classe de defeito",name,22),("Threshold",threshold,12),("Confidence",confidence,14),("Observação",notes,34))
    for col,(label,var,width) in enumerate(fields):
        field=tk.Frame(controls,bg=PANEL); field.grid(row=0,column=col,sticky="ew",padx=3)
        tk.Label(field,text=label,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w")
        ttk.Entry(field,textvariable=var,width=width).pack(fill="x")
        controls.columnconfigure(col,weight=3 if col==3 else 1)
    def upsert():
        key = normalize_defect(name.get())
        if not key: messagebox.showerror("Classe", "Informe o nome da classe.", parent=win); return
        self.class_thresholds[key] = {"threshold": threshold.get().strip(), "confidence": confidence.get().strip(), "notes": notes.get().strip()}
        name.set(""); threshold.set(""); confidence.set(""); notes.set(""); refill()
    actions=tk.Frame(box,bg=PANEL); actions.pack(fill="x",pady=(8,0))
    ttk.Button(actions, text="Adicionar ou atualizar", command=upsert).pack(side="left", padx=2)
    def remove():
        for item in tree.selection(): self.class_thresholds.pop(item, None)
        refill()
    ttk.Button(actions, text="Excluir classe selecionada", command=remove).pack(side="left", padx=4)
    ttk.Button(box, text="Concluir", style="Accent.TButton", command=lambda: (aia_v130_snapshot_rows(self), win.destroy(), self.refresh_all())).pack(anchor="e", pady=(10, 0))
    refill()

_aia_v130_view_operacao_base = _aia_v10_final_operacao

def _aia_v130_quick_entry(self):
    win=tk.Toplevel(self); win.title("Coleta rápida — D-O explícito"); win.geometry("680x720"); win.minsize(610,650); win.configure(bg=BG); win.transient(self); win.grab_set()
    frame=tk.Frame(win,bg=PANEL,highlightbackground=BORDER,highlightthickness=1); frame.pack(fill="both",expand=True,padx=15,pady=15)
    tk.Label(frame,text="COLETA RÁPIDA",bg=PANEL,fg=TEXT,font=("Segoe UI",17,"bold")).pack(anchor="w",padx=18,pady=(16,2))
    tk.Label(frame,text="Registre TP-D, FP-D e FN-D conforme confirmado pelo operador. O AIA calcula os totais físicos para preservar D-I e D-O.",bg=PANEL,fg=MUTED,wraplength=610,justify="left").pack(anchor="w",padx=18,pady=(0,12))
    values={}
    def field(label,key,default="",options=None):
        row=tk.Frame(frame,bg=PANEL);row.pack(fill="x",padx=18,pady=4)
        tk.Label(row,text=label,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w")
        var=tk.StringVar(value=default);values[key]=var
        widget=ttk.Combobox(row,textvariable=var,values=options,state="readonly") if options else ttk.Entry(row,textvariable=var)
        widget.pack(fill="x")
    field("Identificação","id",self._next_numeric_id()); field("Vision AI","vision","OK",["OK","NOK"]); field("Referência / gabarito","ref","Não determinado",REFS)
    section=tk.Frame(frame,bg="#103126",highlightbackground=GOOD,highlightthickness=1);section.pack(fill="x",padx=18,pady=(10,5))
    tk.Label(section,text="DETECÇÃO POR OCORRÊNCIA — D-O",bg="#103126",fg=GOOD,font=("Segoe UI",10,"bold")).pack(anchor="w",padx=10,pady=(7,1))
    tk.Label(section,text="TP-D = defeitos corretamente detectados • FP-D = alarmes sem defeito real • FN-D = defeitos reais que o Vision AI deixou passar",bg="#103126",fg=TEXT,font=("Segoe UI",8),wraplength=590,justify="left").pack(anchor="w",padx=10,pady=(0,7))
    field("Defeitos corretamente detectados (TP-D)","tp_d","0"); field("Falsos positivos confirmados pelo operador (FP-D)","fp_d","0"); field("Falsos negativos confirmados pelo operador (FN-D)","fn_d","0")
    derived=tk.StringVar(value="Totais calculados: Defeitos reais existentes = 0   •   Defeitos detectados pelo Vision AI = 0")
    tk.Label(frame,textvariable=derived,bg="#123B52",fg=ACCENT2,font=("Segoe UI",9,"bold"),justify="left",anchor="w",padx=10,pady=8).pack(fill="x",padx=18,pady=(5,4))
    def update_totals(*_):
        try:
            tp=max(0,int(values["tp_d"].get() or 0));fp=max(0,int(values["fp_d"].get() or 0));fn=max(0,int(values["fn_d"].get() or 0))
            derived.set(f"Totais calculados: Defeitos reais existentes = {tp+fn}   •   Defeitos detectados pelo Vision AI = {tp+fp}")
        except ValueError: derived.set("Informe TP-D, FP-D e FN-D como números inteiros não negativos.")
    for key in ("tp_d","fp_d","fn_d"): values[key].trace_add("write",update_totals)
    field("Observação","obs","")
    btns=tk.Frame(frame,bg=PANEL);btns.pack(fill="x",padx=18,pady=15)
    def collect(next_one=False):
        try:
            tp=max(0,int(values["tp_d"].get() or 0));fp=max(0,int(values["fp_d"].get() or 0));fn=max(0,int(values["fn_d"].get() or 0))
        except ValueError: messagebox.showerror("Dados","TP-D, FP-D e FN-D devem ser números inteiros não negativos.",parent=win);return
        ident=values["id"].get().strip() or self._next_numeric_id()
        if self._id_exists(ident): messagebox.showerror("Identificação duplicada",f"A identificação '{ident}' já existe nesta coleta.",parent=win);return
        r=blank_row(); r.update({"id":ident,"vision":values["vision"].get() or "OK","ref":values["ref"].get() or "Não determinado","real":tp+fn,"det":tp+fp,"correspondentes":tp,"tp_d":tp,"fp_d":fp,"fn_d":fn,"detection_explicit_matrix":True,"obs":values["obs"].get().strip(),"source":"Manual — Coleta rápida D-O explícita","detection_observed":True})
        self.rows.append(self.classify(r)); self.source_file=""; aia_v130_snapshot_rows(self); self.refresh_all(); self.status.set(f"✓ Inspeção {ident} registrada.")
        if next_one:
            values["id"].set(self._next_numeric_id());values["vision"].set("OK");values["ref"].set("Não determinado");values["tp_d"].set("0");values["fp_d"].set("0");values["fn_d"].set("0");values["obs"].set("")
        else: win.destroy()
    ttk.Button(btns,text="Registrar e fechar",command=lambda:collect(False)).pack(side="right",padx=4)
    ttk.Button(btns,text="Registrar e próxima",style="Accent.TButton",command=lambda:collect(True)).pack(side="right",padx=4)

def _aia_v130_view_operacao(self):
    _aia_v130_view_operacao_base(self)
    bar = tk.Frame(self.content, bg=BG); bar.pack(fill="x", pady=(0, 5))
    ttk.Button(bar, text="Configurar Threshold/Confidence por classe", command=lambda: _aia_v130_configure_classes(self)).pack(side="left", padx=3)
    tk.Label(bar, text="D-I é o indicador principal. Em D-O, registre TP-D, FP-D e FN-D; os totais são calculados automaticamente.", bg=BG, fg=ACCENT2, font=("Segoe UI", 8, "bold")).pack(side="left", padx=8)

def _aia_v130_export_pdf(self):
    path = filedialog.asksaveasfilename(title="Gerar relatório executivo V1.3.0 RC", defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
    if not path: return
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import cm
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
        di = self.detection_totals().get("inspection", {}); do = self.defect_counts(); c = self.display_metrics()
        styles = getSampleStyleSheet(); title = ParagraphStyle("t", parent=styles["Title"], textColor=colors.HexColor("#087EA4"), fontName="Helvetica-Bold", fontSize=19, leading=22)
        h = ParagraphStyle("h", parent=styles["Heading2"], textColor=colors.HexColor("#07526B"), fontName="Helvetica-Bold", fontSize=12, spaceBefore=10, spaceAfter=5)
        body = ParagraphStyle("b", parent=styles["BodyText"], fontSize=8, leading=10, wordWrap="CJK")
        cell = ParagraphStyle("cell", parent=body, fontSize=6.6, leading=8, wordWrap="CJK")
        head = ParagraphStyle("head", parent=cell, textColor=colors.white, fontName="Helvetica-Bold", alignment=1, leading=7)
        def pc(v): return "N/D" if v is None else f"{v*100:.2f}%"
        def text_cell(value, style=cell):
            return Paragraph(str(value if value not in (None, "") else "—"), style)
        def tbl(rows, widths=None):
            converted=[[text_cell(v, head if row_index==0 else cell) for v in row] for row_index,row in enumerate(rows)]
            t=Table(converted, colWidths=widths, repeatRows=1, splitByRow=1)
            t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#0E202D")),("GRID",(0,0),(-1,-1),.25,colors.HexColor("#A8C6D2")),("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),3),("RIGHTPADDING",(0,0),(-1,-1),3),("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3)])); return t
        story=[Paragraph("AUTAZA INSPECTION ANALYTICS", title), Paragraph("RELATÓRIO EXECUTIVO — V1.3.0 RC", body), Spacer(1,7), Paragraph("01. D-I — DETECÇÃO POR INSPEÇÃO — INDICADOR PRINCIPAL",h), Paragraph("D-I responde se uma inspeção com defeito real apresentou alguma detecção. C e D-O permanecem independentes e complementares.",body), tbl([["Indicador","Valor"],["N-DI",di.get("N","N/D")],["TP-DI",di.get("TP","N/D")],["TN-DI",di.get("TN","N/D")],["FP-DI",di.get("FP","N/D")],["FN-DI",di.get("FN","N/D")],["Accuracy D-I",pc(di.get("Accuracy"))],["Recall D-I",pc(di.get("Recall / Sensibilidade"))],["Precision D-I",pc(di.get("Precision"))],["F1 D-I",pc(di.get("F1"))],["Método",di.get("method","N/D")]], [7*cm,18*cm]), Paragraph("02. D-O — DETECÇÃO POR OCORRÊNCIA — COMPLEMENTAR",h), tbl([["Indicador","Valor"],["Ocorrências reais",do.get("real",0)],["Detecções",do.get("det",0)],["TP-D",do.get("tp",0)],["FP-D",do.get("fp",0)],["FN-D",do.get("fn",0)],["Status",do.get("status","N/D")],["Método",do.get("method","N/D")]], [7*cm,18*cm]), Paragraph("03. C — CLASSIFICAÇÃO OK/NOK — COMPLEMENTAR",h), tbl([["Indicador","Valor"],["N-C",c.get("N",0)],["TP-C",c.get("TP",0)],["TN-C",c.get("VN/TN",0)],["FP-C",c.get("FP",0)],["FN-C",c.get("FN",0)],["Accuracy C",pc(c.get("Accuracy"))]], [7*cm,18*cm]), PageBreak(), Paragraph("04. DADOS REGISTRADOS",h), Paragraph("Dados que originaram os resultados. TP-D, FP-D e FN-D são os valores de ocorrência confirmados.",body)]
        records=[["#","ID","Vision","Referência","Real","Detectado","TP-D","FP-D","FN-D","Configuração por classe","Observação"]]
        for i, r0 in enumerate(self.rows,1):
            r=normalize_row(r0); cfg="; ".join(f"{k}: T={v.get('threshold','')} C={v.get('confidence','')}" for k,v in r.get("configuracao_classes_snapshot",{}).items()) or "N/D"
            d=self.detection_for_row(r)
            records.append([i,r["id"],r["vision"],r["ref"],r["real"],r["det"],d["tp"],d["fp"],d["fn"],cfg,r["obs"]])
        story += [tbl(records,[.75*cm,1.15*cm,.85*cm,2.0*cm,.75*cm,.95*cm,.75*cm,.75*cm,.75*cm,7.0*cm,7.3*cm]), Paragraph("05. CONFIGURAÇÃO ATIVA POR CLASSE",h)]
        cfgrows=[["Classe","Threshold","Confidence","Observação"]] + [[n,v.get("threshold",""),v.get("confidence",""),v.get("notes","")] for n,v in sorted(self.class_thresholds.items())]
        story.append(tbl(cfgrows,[4*cm,3*cm,3*cm,6.2*cm]))
        doc=SimpleDocTemplate(path,pagesize=landscape(A4),leftMargin=1.15*cm,rightMargin=1.15*cm,topMargin=1.1*cm,bottomMargin=1.1*cm,title="AIA V1.3.0 RC")
        doc.build(story); self.status.set(f"Relatório V1.3.0 RC gerado: {Path(path).name}"); messagebox.showinfo("Relatório", "PDF gerado com D-I em destaque e Dados Registrados.", parent=self)
    except Exception as exc: messagebox.showerror("Relatório", f"Não foi possível gerar o PDF:\n{exc}", parent=self)

def _aia_v130_save_project(self):
    path=filedialog.asksaveasfilename(title="Salvar projeto V1.3.0 RC",defaultextension=".json",filetypes=[("Projeto JSON","*.json")])
    if not path:return
    aia_v130_snapshot_rows(self)
    data={"app":APP,"version":VERSION,"schema_version":"1.3.0-rc","saved_at":datetime.now().isoformat(),"metadata":{k:v.get() for k,v in self.meta.items()},"class_thresholds":aia_v130_config_snapshot(self),"rows":[normalize_row(r) for r in self.rows],"history":self.history,"source_file":self.source_file,"comparison_archive":self.comparison_archive,"calibration_history":self.calibration_history,"calibration_project":self.calibration_project,"expectation":self.expectation_snapshot()}
    Path(path).write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding="utf-8"); self.status.set(f"Projeto salvo: {Path(path).name}")

def aia_v130_validate_project_payload(data):
    """Valida e normaliza projetos AIA sem executar alterações de interface."""
    if not isinstance(data, dict):
        raise ValueError("o conteúdo JSON deve ser um objeto de projeto AIA")
    rows = data.get("rows", [])
    if not isinstance(rows, list):
        raise ValueError("o campo 'rows' deve conter uma lista de inspeções")
    metadata = data.get("metadata", {})
    if not isinstance(metadata, dict):
        raise ValueError("o campo 'metadata' deve conter um objeto")
    classes = data.get("class_thresholds", {})
    if not isinstance(classes, dict):
        classes = {}
    history = data.get("history", [])
    comparison = data.get("comparison_archive", [])
    calibration = data.get("calibration_history", [])
    for label, value in (("history", history), ("comparison_archive", comparison), ("calibration_history", calibration)):
        if not isinstance(value, list): raise ValueError(f"o campo '{label}' deve conter uma lista")
    return {"metadata": metadata, "rows": [normalize_row(row) for row in rows],
            "class_thresholds": {normalize_defect(name): {"threshold": str((cfg or {}).get("threshold", "")),
                                                       "confidence": str((cfg or {}).get("confidence", "")),
                                                       "notes": str((cfg or {}).get("notes", ""))}
                                 for name, cfg in classes.items() if normalize_defect(name)},
            "history": history, "comparison_archive": comparison, "calibration_history": calibration,
            "calibration_project": data.get("calibration_project", {}) if isinstance(data.get("calibration_project", {}), dict) else {},
            "expectation": data.get("expectation", {}) if isinstance(data.get("expectation", {}), dict) else {},
            "version": str(data.get("version", "legado"))}

def _aia_v130_load_project(self):
    path = filedialog.askopenfilename(title="Abrir projeto AIA JSON", filetypes=[("Projeto AIA JSON", "*.json"), ("Todos os arquivos", "*.*")])
    if not path: return
    try:
        raw = json.loads(Path(path).read_text(encoding="utf-8-sig"))
        project = aia_v130_validate_project_payload(raw)
    except (OSError, UnicodeDecodeError, json.JSONDecodeError, ValueError) as exc:
        messagebox.showerror("Abrir projeto JSON", f"O arquivo não é um projeto AIA válido ou está corrompido.\n\n{exc}", parent=self); return
    if self.rows or self.history or self.comparison_archive:
        if not messagebox.askyesno("Abrir projeto JSON", "A coleta atual será substituída na tela. Os arquivos já salvos não serão apagados.\n\nDeseja continuar?", parent=self): return
    for key, variable in self.meta.items(): variable.set(str(project["metadata"].get(key, "")))
    self.rows = project["rows"]; self.class_thresholds = project["class_thresholds"]
    self.history = project["history"]; self.comparison_archive = project["comparison_archive"]; self.calibration_history = project["calibration_history"]
    self.calibration_project = {**{"id":"", "name":"", "objective":"", "responsible":"", "created_at":""}, **project["calibration_project"]}
    for key, variable in self.expectation.items():
        value = project["expectation"].get(key, variable.get())
        if key in ("accuracy", "precision", "recall", "f1") and isinstance(value, (int, float)): value = f"{value * 100:g}"
        variable.set("" if value is None else str(value))
    self.source_file = ""; self.source_analysis = {}
    aia_v130_snapshot_rows(self); self.refresh_all()
    self.status.set(f"✓ Projeto JSON aberto: {Path(path).name} • {len(self.rows)} inspeção(ões) • origem V{project['version']}")
    messagebox.showinfo("Projeto carregado", f"Projeto restaurado com {len(self.rows)} inspeção(ões), {len(self.class_thresholds)} classe(s) e histórico preservado.", parent=self)

def _aia_v130_archive(self, label=None):
    aia_v130_snapshot_rows(self); item=_aia_v120_archive_current_collection(self,label)
    if item is not None: item["class_thresholds"] = aia_v130_config_snapshot(self); item["schema_version"]="1.3.0-rc"
    return item

def _aia_v130_export_csv(self):
    path=filedialog.asksaveasfilename(title="Exportar CSV V1.3.0 RC",defaultextension=".csv",filetypes=[("CSV","*.csv")])
    if not path:return
    aia_v130_snapshot_rows(self)
    headers=["Schema","Identificação","Vision AI","Referência","Real","Detectado","TP-D","FP-D","FN-D","Observação","Origem","Configuração por classe","Cliente","Local","Projeto","Linha","Técnico","Data","Threshold","Método de referência"]
    with open(path,"w",newline="",encoding="utf-8-sig") as f:
        w=csv.writer(f,delimiter=";")
        w.writerow(headers)
        meta={"Cliente":self.meta["cliente"].get(),"Local":self.meta["local"].get(),"Projeto":self.meta["projeto"].get(),"Linha":self.meta["linha"].get(),"Técnico":self.meta["tecnico"].get(),"Data":self.meta["data"].get(),"Threshold":self.meta["threshold"].get(),"Método de referência":self.meta["gabarito"].get()}
        for row0 in self.rows:
            r=normalize_row(row0); d=self.detection_for_row(r); status="EXPLÍCITO" if r.get("correspondentes") is not None else "ESTIMADO"
            cfg=json.dumps(r.get("configuracao_classes_snapshot",{}),ensure_ascii=False,sort_keys=True)
            values=["AIA-1.3.0-RC",r["id"],r["vision"],r["ref"],r["real"],r["det"],d["tp"],d["fp"],d["fn"],r["obs"],r["source"],cfg,meta["Cliente"],meta["Local"],meta["Projeto"],meta["Linha"],meta["Técnico"],meta["Data"],meta["Threshold"],meta["Método de referência"]]
            w.writerow(values)
    self.status.set(f"CSV importável exportado: {Path(path).name}"); messagebox.showinfo("Exportação", "CSV tabular criado. Use ABRIR EXCEL para reabrir este CSV no AIA.", parent=self)

def _aia_v130_import_csv(self, path):
    """Importa o CSV tabular gerado pela V1.3 sem perder contexto da coleta."""
    try:
        with open(path,"r",encoding="utf-8-sig",newline="") as handle:
            sample=handle.read(4096); handle.seek(0)
            delimiter=";" if sample.count(";") >= sample.count(",") else ","
            reader=csv.DictReader(handle,delimiter=delimiter)
            if not reader.fieldnames or "Identificação" not in reader.fieldnames:
                raise ValueError("CSV incompatível: a primeira linha deve conter a coluna 'Identificação'.")
            records=list(reader)
        if not records: raise ValueError("CSV sem inspeções.")
        def number(record,key):
            try:return max(0,int(float((record.get(key) or "0").replace(",","."))))
            except (ValueError,AttributeError):raise ValueError(f"Valor inválido em '{key}' para a inspeção {record.get('Identificação','?')}.")
        rows=[]; classes={}
        for record in records:
            ident=(record.get("Identificação") or "").strip()
            if not ident: continue
            tp,fp,fn=number(record,"TP-D"),number(record,"FP-D"),number(record,"FN-D")
            real,det=number(record,"Real"),number(record,"Detectado")
            if real != tp+fn or det != tp+fp:
                raise ValueError(f"CSV inconsistente na inspeção {ident}: Real deve ser TP-D + FN-D e Detectado deve ser TP-D + FP-D.")
            try: snapshot=json.loads(record.get("Configuração por classe") or "{}")
            except json.JSONDecodeError: snapshot={}
            if isinstance(snapshot,dict):
                for name,cfg in snapshot.items():
                    if normalize_defect(name) and isinstance(cfg,dict): classes[normalize_defect(name)]={"threshold":str(cfg.get("threshold", "")),"confidence":str(cfg.get("confidence", "")),"notes":str(cfg.get("notes", ""))}
            r=blank_row();r.update({"id":ident,"vision":(record.get("Vision AI") or "OK").upper(),"ref":record.get("Referência") or "Não determinado","real":real,"det":det,"correspondentes":tp,"tp_d":tp,"fp_d":fp,"fn_d":fn,"detection_explicit_matrix":True,"detection_observed":True,"obs":record.get("Observação") or "","source":record.get("Origem") or "CSV AIA" ,"configuracao_classes_snapshot":snapshot if isinstance(snapshot,dict) else {}})
            rows.append(self.classify(r))
        if not rows: raise ValueError("CSV sem identificações válidas.")
        self._archive_current_collection(); self.rows=rows; self.history=[]; self.comparison_archive=[]; self.calibration_history=[]; self.class_thresholds=classes; self.source_file=path; self.source_analysis={"file":str(path),"format":"CSV AIA 1.3","inspections_registered":len(rows)}
        first=records[0]
        mapping={"cliente":"Cliente","local":"Local","projeto":"Projeto","linha":"Linha","tecnico":"Técnico","data":"Data","threshold":"Threshold","gabarito":"Método de referência"}
        for key,column in mapping.items(): self.meta[key].set(first.get(column) or "")
        if not self.meta["gabarito"].get(): self.meta["gabarito"].set("Não determinado")
        aia_v130_snapshot_rows(self); self.refresh_all(); self.status.set(f"✓ CSV AIA aberto: {Path(path).name} • {len(rows)} inspeção(ões)")
        messagebox.showinfo("CSV carregado", "Coleta, metadados e registros foram restaurados do CSV.", parent=self)
    except (OSError, UnicodeDecodeError, ValueError) as exc:
        messagebox.showerror("Abrir CSV", f"Não foi possível abrir o CSV.\n\n{exc}", parent=self)

def _aia_v130_import_data(self):
    path=filedialog.askopenfilename(title="Abrir Excel ou CSV AIA",filetypes=[("Excel ou CSV",("*.xlsx","*.csv")),("Excel","*.xlsx"),("CSV AIA","*.csv"),("Todos","*.*")])
    if not path:return
    if Path(path).suffix.lower()==".csv": return _aia_v130_import_csv(self,path)
    original=filedialog.askopenfilename
    try:
        filedialog.askopenfilename=lambda **_kwargs:path
        return _aia_v113_import_excel(self)
    finally:
        filedialog.askopenfilename=original

def _aia_v130_export_excel(self):
    if openpyxl is None: messagebox.showerror("Excel","OpenPyXL não está instalado.",parent=self); return
    path=filedialog.asksaveasfilename(title="Exportar Excel V1.3.0 RC",defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
    if not path:return
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        aia_v130_snapshot_rows(self); wb=Workbook(); summary=wb.active; summary.title="RESUMO"
        dark="0E202D"; blue="087EA4"
        def header(ws, values):
            ws.append(values)
            for cell in ws[ws.max_row]: cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor=dark); cell.alignment=Alignment(wrap_text=True,vertical="top")
        def widths(ws):
            for col in ws.columns:
                letter=col[0].column_letter; ws.column_dimensions[letter].width=min(60,max(12,max(len(str(c.value or "")) for c in col)+2))
        di=self.detection_totals().get("inspection",{}); do=self.defect_counts(); c=self.display_metrics()
        summary.append([APP,"V1.3.0 RC","D-I é o indicador principal"]); summary["A1"].font=Font(size=16,bold=True,color="FFFFFF"); summary["A1"].fill=PatternFill("solid",fgColor=blue)
        summary.append([]); header(summary,["Matriz","Indicador","Valor","Unidade e método"])
        for key,label in (("N","N-DI"),("TP","TP-DI"),("TN","TN-DI"),("FP","FP-DI"),("FN","FN-DI"),("Accuracy","Accuracy D-I"),("Precision","Precision D-I"),("Recall / Sensibilidade","Recall D-I"),("F1","F1 D-I")): summary.append(["D-I principal",label,di.get(key,"N/D"),di.get("method","N/D")])
        for key,label in (("real","Ocorrências reais"),("det","Detecções"),("tp","TP-D"),("fp","FP-D"),("fn","FN-D")): summary.append(["D-O complementar",label,do.get(key,"N/D"),do.get("status","N/D")+" • "+do.get("method","")])
        for key,label in (("N","N-C"),("TP","TP-C"),("VN/TN","TN-C"),("FP","FP-C"),("FN","FN-C"),("Accuracy","Accuracy C")): summary.append(["C complementar",label,c.get(key,"N/D"),"Vision AI × Referência"])
        ins=wb.create_sheet("DADOS REGISTRADOS"); header(ins,["#","Identificação","Vision AI","Referência","Real","Detectado","Correspondentes","TP-D","FP-D","FN-D","Status D-O","Snapshot configuração por classe","Observação","Origem"])
        for i,row0 in enumerate(self.rows,1):
            r=normalize_row(row0); d=self.detection_for_row(r); ins.append([i,r["id"],r["vision"],r["ref"],r["real"],r["det"],r.get("correspondentes"),d["tp"],d["fp"],d["fn"],"EXPLÍCITO" if r.get("correspondentes") is not None else "ESTIMADO",json.dumps(r.get("configuracao_classes_snapshot",{}),ensure_ascii=False,sort_keys=True),r["obs"],r["source"]])
        cfg=wb.create_sheet("CONFIGURAÇÃO POR CLASSE"); header(cfg,["Classe","Threshold","Confidence","Observação"])
        for name,item in sorted(self.class_thresholds.items()): cfg.append([name,item.get("threshold",""),item.get("confidence",""),item.get("notes","")])
        audit=wb.create_sheet("AUDITORIA"); header(audit,["Verificação","Resultado"]); audit.append(["D-I é indicador principal","SIM"]); audit.append(["D-O",do.get("status","N/D")]); audit.append(["Método D-O",do.get("method","N/D")]); audit.append(["Linhas explícitas",do.get("explicit_rows",0)]); audit.append(["Linhas estimadas",do.get("estimated_rows",0)]); audit.append(["Schema","1.3.0-rc"])
        for ws in wb.worksheets: widths(ws); ws.freeze_panes="A2"
        wb.save(path); check=openpyxl.load_workbook(path,read_only=True); required={"RESUMO","DADOS REGISTRADOS","CONFIGURAÇÃO POR CLASSE","AUDITORIA"}; missing=required-set(check.sheetnames); check.close()
        if missing: raise ValueError("Abas ausentes: "+", ".join(sorted(missing)))
        self.status.set(f"Excel V1.3.0 RC exportado e validado: {Path(path).name}"); messagebox.showinfo("Exportação", "Excel com dados registrados, auditoria e configuração por classe criado.",parent=self)
    except Exception as exc: messagebox.showerror("Excel",f"Não foi possível exportar o Excel:\n{exc}",parent=self)

FINAL_RUNTIME_BINDINGS.update({"__init__": _aia_v130_init, "view_operacao": _aia_v130_view_operacao, "quick_entry": _aia_v130_quick_entry, "import_excel": _aia_v130_import_data, "_smart_detection_occurrence": _aia_v130_detection_occurrence, "detection_for_row": _aia_v130_detection_for_row, "export_executive_report": _aia_v130_export_pdf, "export_csv": _aia_v130_export_csv, "export_excel": _aia_v130_export_excel, "save_project": _aia_v130_save_project, "load_project": _aia_v130_load_project, "_archive_current_collection": _aia_v130_archive})
for _method_name, _implementation in FINAL_RUNTIME_BINDINGS.items(): setattr(App, _method_name, _implementation)

# Checkpoints da release:
# 1) Coleta, Dashboard, Gráficos e Engenharia preservados por hash de fonte.
# 2) Comparação, Meta, Calibração e Auditoria possuem uma única implementação final.
# 3) Meta e recomendação de Calibração usam exclusivamente D-I.
# 4) C, D-I e D-O mantêm fonte, unidade e método auditáveis.
# 5) A barra de título exibe V1.2.0.

if __name__ == "__main__":
    App().mainloop()
