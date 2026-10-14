#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Office -> PDF em lote  (versão 2.0)
===================================
Escolhe uma pasta e converte para PDF todos os ficheiros:
  Word        .doc  .docx
  Excel       .xls  .xlsx      (respeita a configuração de impressão de cada ficheiro)
  PowerPoint  .ppt  .pptx

Motores (usa o Office instalado e, se faltar ou falhar, o LibreOffice):
  1. Microsoft Office - só Windows (requer pywin32 quando se corre o .py)
  2. LibreOffice      - Windows, macOS e Linux (gratuito)

Utilização:
  python office_para_pdf.py                       -> janela para escolher a pasta
  python office_para_pdf.py "C:\\pasta" [opções]   -> sem janela
Opções:  --subpastas   --motor=auto|office|libreoffice   --tipos=word,excel,powerpoint

Os PDF ficam numa subpasta "PDF" dentro da pasta escolhida.
Se existirem ficheiros com o mesmo nome e tipos diferentes (ex.: folha.xls e folha.xlsx),
os PDF ficam como "folha (xls).pdf" e "folha (xlsx).pdf" para não se sobreporem.
"""

import os
import sys
import shutil
import platform
import tempfile
import threading
import subprocess
from pathlib import Path

VERSAO = "2.0"
PASTA_SAIDA = "PDF"

TIPOS = {
    "word": {".doc", ".docx"},
    "excel": {".xls", ".xlsx"},
    "powerpoint": {".ppt", ".pptx"},
}
ROTULOS = {"word": "Word", "excel": "Excel", "powerpoint": "PowerPoint"}
EXT_PARA_TIPO = {ext: tipo for tipo, exts in TIPOS.items() for ext in exts}
PROG_ID = {
    "word": "Word.Application",
    "excel": "Excel.Application",
    "powerpoint": "PowerPoint.Application",
}
MODULO_LO = {
    "word": "writer_pdf_Export",
    "excel": "calc_pdf_Export",
    "powerpoint": "impress_pdf_Export",
}
EH_WINDOWS = platform.system() == "Windows"


# --------------------------------------------------------------------------
# Deteção de motores
# --------------------------------------------------------------------------
def pywin32_disponivel():
    if not EH_WINDOWS:
        return False
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


def encontrar_libreoffice():
    for nome in ("soffice", "libreoffice"):
        caminho = shutil.which(nome)
        if caminho:
            return caminho
    for c in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        if os.path.exists(c):
            return c
    return None


# --------------------------------------------------------------------------
# Recolha de ficheiros e nomes de saída
# --------------------------------------------------------------------------
def listar_ficheiros(pasta: Path, subpastas: bool, tipos):
    exts = set()
    for t in tipos:
        exts |= TIPOS[t]
    it = pasta.rglob("*") if subpastas else pasta.glob("*")
    ficheiros = []
    for f in it:
        if not f.is_file() or f.suffix.lower() not in exts:
            continue
        if f.name.startswith("~$"):                 # ficheiros temporários do Office
            continue
        if PASTA_SAIDA in f.relative_to(pasta).parts[:-1]:
            continue                                  # ignora a própria pasta de saída
        ficheiros.append(f)
    return sorted(ficheiros)


def planear(ficheiros, raiz: Path):
    """Devolve [(origem, tipo, destino)] com nomes de PDF únicos."""
    grupos = {}
    for f in ficheiros:
        chave = (str(f.relative_to(raiz).parent).lower(), f.stem.lower())
        grupos[chave] = grupos.get(chave, 0) + 1
    plano = []
    for f in ficheiros:
        rel = f.relative_to(raiz)
        chave = (str(rel.parent).lower(), f.stem.lower())
        nome = f.stem if grupos[chave] == 1 else f"{f.stem} ({f.suffix[1:].lower()})"
        destino = raiz / PASTA_SAIDA / rel.parent / (nome + ".pdf")
        plano.append((f, EXT_PARA_TIPO[f.suffix.lower()], destino))
    return plano


# --------------------------------------------------------------------------
# Motor 1: Microsoft Office (COM, Windows)
# --------------------------------------------------------------------------
class MotorOffice:
    nome = "Office"

    def __init__(self):
        import pythoncom
        pythoncom.CoInitialize()          # na mesma thread que vai usar o COM
        self._pythoncom = pythoncom
        self._apps = {}

    def _app(self, tipo):
        if tipo in self._apps:
            return self._apps[tipo]
        import win32com.client
        try:
            app = win32com.client.DispatchEx(PROG_ID[tipo])
            try:
                app.AutomationSecurity = 3      # desativa macros durante a conversão
            except Exception:
                pass
            if tipo == "word":
                app.Visible = False
                app.DisplayAlerts = 0
            elif tipo == "excel":
                app.Visible = False
                app.DisplayAlerts = False
                app.AskToUpdateLinks = False
                app.EnableEvents = False
            # PowerPoint não permite esconder a janela da aplicação; abre-se sem janela
        except Exception:
            app = False
        self._apps[tipo] = app
        return app

    def disponivel(self, tipo):
        return bool(self._app(tipo))

    def converter(self, tipo, origem: Path, destino: Path):
        app = self._app(tipo)
        if not app:
            raise RuntimeError(f"{ROTULOS[tipo]} não está disponível")
        destino.parent.mkdir(parents=True, exist_ok=True)
        src, dst = str(origem.resolve()), str(destino.resolve())

        if tipo == "word":
            doc = app.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False,
                                     ConfirmConversions=False)
            try:
                # ExportAsFixedFormat preserva imagens com transparência e símbolos
                doc.ExportAsFixedFormat(
                    OutputFileName=dst, ExportFormat=17, OpenAfterExport=False,
                    OptimizeFor=0, Range=0, Item=0, IncludeDocProps=True,
                    KeepIRM=True, CreateBookmarks=1, DocStructureTags=True,
                    BitmapMissingFonts=False, UseISO19005_1=False)
            finally:
                doc.Close(False)

        elif tipo == "excel":
            wb = app.Workbooks.Open(src, UpdateLinks=0, ReadOnly=True,
                                    IgnoreReadOnlyRecommended=True, AddToMru=False)
            try:
                # Respeita áreas de impressão, orientação, escala e quebras de página
                try:
                    wb.ExportAsFixedFormat(
                        Type=0, Filename=dst, Quality=0, IncludeDocProperties=True,
                        IgnorePrintAreas=False, OpenAfterPublish=False)
                except TypeError:                # chamada alternativa (posicional)
                    wb.ExportAsFixedFormat(0, dst)
            finally:
                wb.Close(SaveChanges=False)

        else:  # powerpoint
            pres = app.Presentations.Open(src, ReadOnly=True, Untitled=False,
                                          WithWindow=False)
            try:
                try:
                    pres.ExportAsFixedFormat(Path=dst, FixedFormatType=2, Intent=2)
                except Exception:
                    pres.SaveAs(dst, 32)         # 32 = ppSaveAsPDF
            finally:
                pres.Close()

    def fechar(self):
        for app in self._apps.values():
            if app:
                try:
                    app.Quit()
                except Exception:
                    pass
        self._apps.clear()
        try:
            self._pythoncom.CoUninitialize()
        except Exception:
            pass


# --------------------------------------------------------------------------
# Motor 2: LibreOffice (headless)
# --------------------------------------------------------------------------
class MotorLibreOffice:
    nome = "LibreOffice"

    def __init__(self, soffice):
        self.soffice = soffice
        self._perfil = tempfile.mkdtemp(prefix="lo_perfil_")   # perfil próprio, reutilizado

    @staticmethod
    def _filtro(tipo):
        # Compressão sem perdas: o JPEG não suporta transparência (fundos pretos)
        return ("pdf:" + MODULO_LO[tipo] + ":{"
                "\"UseLosslessCompression\":{\"type\":\"boolean\",\"value\":\"true\"},"
                "\"ReduceImageResolution\":{\"type\":\"boolean\",\"value\":\"false\"}}")

    def disponivel(self, tipo):
        return True

    def converter(self, tipo, origem: Path, destino: Path, timeout=300):
        destino.parent.mkdir(parents=True, exist_ok=True)
        extra = {"creationflags": subprocess.CREATE_NO_WINDOW} if EH_WINDOWS else {}
        with tempfile.TemporaryDirectory(prefix="lo_saida_") as tmp:
            gerado = Path(tmp) / (origem.stem + ".pdf")
            erro = "sem saída"
            for filtro in (self._filtro(tipo), "pdf"):
                r = subprocess.run(
                    [self.soffice, f"-env:UserInstallation={Path(self._perfil).as_uri()}",
                     "--headless", "--norestore", "--convert-to", filtro,
                     "--outdir", tmp, str(origem)],
                    stdin=subprocess.DEVNULL,       # necessário em executáveis sem consola
                    capture_output=True, text=True, timeout=timeout, **extra)
                if r.returncode == 0 and gerado.exists():
                    if destino.exists():
                        destino.unlink()
                    shutil.move(str(gerado), str(destino))
                    return
                erro = (r.stderr or r.stdout).strip() or erro
            raise RuntimeError(erro)

    def fechar(self):
        shutil.rmtree(self._perfil, ignore_errors=True)


# --------------------------------------------------------------------------
# Orquestração
# --------------------------------------------------------------------------
def converter_pasta(pasta, subpastas, tipos, motor, log, progresso):
    """Converte a pasta. motor: 'auto' | 'office' | 'libreoffice'.
    Devolve (ok, falhas)."""
    pasta = Path(pasta)
    tipos = [t for t in TIPOS if t in set(tipos)]
    ficheiros = listar_ficheiros(pasta, subpastas, tipos)
    if not ficheiros:
        log("Não foram encontrados ficheiros dos tipos escolhidos nesta pasta.")
        return 0, 0
    plano = planear(ficheiros, pasta)
    log(f"{len(plano)} ficheiro(s) encontrado(s).")

    office = lo = None
    try:
        if motor in ("auto", "office") and pywin32_disponivel():
            office = MotorOffice()
        soffice = encontrar_libreoffice() if motor in ("auto", "libreoffice") else None
        if soffice:
            lo = MotorLibreOffice(soffice)

        # Resumo dos motores por tipo
        usados = sorted({t for _, t, _ in plano}, key=list(TIPOS).index)
        for t in usados:
            partes = []
            if office and office.disponivel(t):
                partes.append(f"Microsoft {ROTULOS[t]}")
            if lo:
                partes.append("LibreOffice")
            log(f"Motor para {ROTULOS[t]}: " + (" → ".join(partes) if partes else "NENHUM"))
        if not office and not lo:
            log("\nNão encontrei o Microsoft Office (pywin32) nem o LibreOffice.\n"
                "Instala o Office ou o LibreOffice (https://www.libreoffice.org).")
            return 0, 0
        log("")

        ok = falhas = 0
        for i, (origem, tipo, destino) in enumerate(plano, 1):
            rel = origem.relative_to(pasta)
            motores = []
            if office and office.disponivel(tipo):
                motores.append(office)
            if lo:
                motores.append(lo)
            if not motores:
                log(f"ERRO    {rel}  (sem motor disponível para {ROTULOS[tipo]})")
                falhas += 1
            else:
                ultimo_erro = None
                for m in motores:
                    try:
                        m.converter(tipo, origem, destino)
                        extra = "" if m is motores[0] else f"  [{m.nome}]"
                        log(f"OK      {rel}{extra}")
                        ok += 1
                        ultimo_erro = None
                        break
                    except Exception as e:
                        ultimo_erro = e
                        if m is not motores[-1]:
                            log(f"        {rel}: {m.nome} falhou ({e}); a tentar {motores[-1].nome}…")
                if ultimo_erro is not None:
                    log(f"ERRO    {rel}  ({ultimo_erro})")
                    falhas += 1
            progresso(i, len(plano))

        log(f"\nConcluído: {ok} convertido(s), {falhas} com erro.")
        log(f"PDF guardados em: {pasta / PASTA_SAIDA}")
        return ok, falhas
    finally:
        if office:
            office.fechar()
        if lo:
            lo.fechar()


# --------------------------------------------------------------------------
# Interface gráfica
# --------------------------------------------------------------------------
def abrir_pasta_no_sistema(caminho):
    try:
        if EH_WINDOWS:
            os.startfile(str(caminho))            # noqa: S606
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", str(caminho)])
        else:
            subprocess.Popen(["xdg-open", str(caminho)])
    except Exception:
        pass


def abrir_gui():
    import tkinter as tk
    from tkinter import filedialog, ttk

    root = tk.Tk()
    root.title(f"Office → PDF Criação: Pedro Teixeira (v{VERSAO})")
    root.geometry("720x540")

    pasta_sel = tk.StringVar(value="")
    subpastas = tk.BooleanVar(value=False)
    motor_sel = tk.StringVar(value="Automático")
    tipo_vars = {t: tk.BooleanVar(value=True) for t in TIPOS}

    # --- linha 1: pasta
    topo = ttk.Frame(root, padding=10)
    topo.pack(fill="x")
    ttk.Entry(topo, textvariable=pasta_sel, state="readonly").pack(
        side="left", fill="x", expand=True, padx=(0, 8))
    botao_esc = ttk.Button(topo, text="Escolher pasta…")
    botao_esc.pack(side="left")

    # --- linha 2: tipos
    tipos_f = ttk.LabelFrame(root, text="Converter", padding=(10, 4))
    tipos_f.pack(fill="x", padx=10)
    for t in TIPOS:
        exts = ", ".join(sorted(TIPOS[t]))
        ttk.Checkbutton(tipos_f, text=f"{ROTULOS[t]} ({exts})",
                        variable=tipo_vars[t]).pack(side="left", padx=(0, 14))
    ttk.Checkbutton(tipos_f, text="Incluir subpastas", variable=subpastas).pack(side="right")

    # --- linha 3: motor + ações
    acoes = ttk.Frame(root, padding=(10, 8))
    acoes.pack(fill="x")
    ttk.Label(acoes, text="Motor:").pack(side="left")
    ttk.Combobox(acoes, textvariable=motor_sel, state="readonly", width=18,
                 values=["Automático", "Microsoft Office", "LibreOffice"]).pack(
        side="left", padx=(4, 0))
    botao_abrir = ttk.Button(acoes, text="Abrir pasta PDF", state="disabled")
    botao_abrir.pack(side="right")
    botao_conv = ttk.Button(acoes, text="Converter para PDF", state="disabled")
    botao_conv.pack(side="right", padx=(0, 8))

    barra = ttk.Progressbar(root, mode="determinate")
    barra.pack(fill="x", padx=10)
    caixa = tk.Text(root, height=15, wrap="word", state="disabled")
    caixa.pack(fill="both", expand=True, padx=10, pady=10)

    def log(msg):
        def _f():
            caixa.config(state="normal")
            caixa.insert("end", msg + "\n")
            caixa.see("end")
            caixa.config(state="disabled")
        root.after(0, _f)

    def progresso(i, total):
        root.after(0, lambda: barra.config(maximum=total, value=i))

    def escolher():
        p = filedialog.askdirectory(title="Escolhe a pasta com os ficheiros Office")
        if p:
            pasta_sel.set(p)
            botao_conv.config(state="normal")
            botao_abrir.config(state="disabled")

    def converter():
        if not pasta_sel.get():
            return
        tipos = [t for t, v in tipo_vars.items() if v.get()]
        caixa.config(state="normal")
        caixa.delete("1.0", "end")
        caixa.config(state="disabled")
        if not tipos:
            log("Seleciona pelo menos um tipo de ficheiro (Word, Excel ou PowerPoint).")
            return
        motor = {"Automático": "auto", "Microsoft Office": "office",
                 "LibreOffice": "libreoffice"}[motor_sel.get()]
        botao_conv.config(state="disabled")
        botao_esc.config(state="disabled")
        barra.config(value=0)

        def tarefa():
            try:
                converter_pasta(pasta_sel.get(), subpastas.get(), tipos, motor, log, progresso)
            except Exception as e:
                log(f"Erro inesperado: {e}")
            finally:
                root.after(0, lambda: (botao_conv.config(state="normal"),
                                       botao_esc.config(state="normal"),
                                       botao_abrir.config(state="normal")))

        threading.Thread(target=tarefa, daemon=True).start()

    botao_esc.config(command=escolher)
    botao_conv.config(command=converter)
    botao_abrir.config(command=lambda: abrir_pasta_no_sistema(
        Path(pasta_sel.get()) / PASTA_SAIDA
        if (Path(pasta_sel.get()) / PASTA_SAIDA).exists() else Path(pasta_sel.get())))

    root.mainloop()


# --------------------------------------------------------------------------
# Autoteste e linha de comandos
# --------------------------------------------------------------------------
def autoteste():
    """Verifica se as dependências ficaram incluídas no executável."""
    erros = []
    try:
        import tkinter  # noqa: F401
    except Exception as e:
        erros.append(f"tkinter: {e}")
    if EH_WINDOWS:
        try:
            import pythoncom  # noqa: F401
            import win32com.client  # noqa: F401
        except Exception as e:
            erros.append(f"pywin32: {e}")
    return erros


def _opcao(args, prefixo, padrao=None):
    for a in args:
        if a.startswith(prefixo):
            return a.split("=", 1)[1]
    return padrao


def main_cli(args):
    pasta = next((a for a in args if not a.startswith("--")), None)
    motor = _opcao(args, "--motor=", "auto").lower()
    motor = {"word": "office", "microsoft": "office", "lo": "libreoffice"}.get(motor, motor)
    if motor not in ("auto", "office", "libreoffice"):
        print("Motor inválido. Usa auto, office ou libreoffice.")
        return 2
    tipos = _opcao(args, "--tipos=")
    tipos = list(TIPOS) if not tipos else [t.strip().lower() for t in tipos.split(",")]
    invalidos = [t for t in tipos if t not in TIPOS]
    if invalidos:
        print("Tipos inválidos:", ", ".join(invalidos), "- usa word, excel, powerpoint.")
        return 2
    _, falhas = converter_pasta(pasta, "--subpastas" in args, tipos, motor,
                                print, lambda i, t: None)
    return 1 if falhas else 0


if __name__ == "__main__":
    argv = sys.argv[1:]
    if "--autoteste" in argv:
        i = argv.index("--autoteste")
        destino = argv[i + 1] if len(argv) > i + 1 else "autoteste.txt"
        erros = autoteste()
        with open(destino, "w", encoding="utf-8") as fh:
            fh.write("OK\n" if not erros else "\n".join(erros) + "\n")
        sys.exit(1 if erros else 0)
    if any(not a.startswith("--") for a in argv):
        sys.exit(main_cli(argv))
    abrir_gui()
