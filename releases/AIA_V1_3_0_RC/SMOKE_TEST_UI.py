# Execute com ambiente grafico. O teste abre a janela de forma invisivel e percorre os modulos.
from pathlib import Path
import importlib.util
SRC=Path(__file__).parent / "AIA_V1_3_0_RC.py"
spec=importlib.util.spec_from_file_location("aia_ui", SRC)
mod=importlib.util.module_from_spec(spec); spec.loader.exec_module(mod)
app=mod.App(); app.withdraw()
for key in ["operacao","dashboard","graficos","comparacao","tecnica","defeitos","engenharia","calibracao","expectativa","auditoria"]:
    app.show(key); app.update_idletasks(); app.update()
    print("OK", key)
app.destroy()
