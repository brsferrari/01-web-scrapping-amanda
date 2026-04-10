import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as tbs
import os
import json
import threading
import time
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed
import undetected_chromedriver as uc
import winreg
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv, set_key
from api import get_product, get_product_info, inspect_ean

load_dotenv()

# ── Global state ───────────────────────────────────────────────────────────────
driver = None
planilha_path = None
cancel_event = threading.Event()


# ── Helpers ────────────────────────────────────────────────────────────────────
def get_browser_version_from_registry(browser: str = "Brave") -> int | None:
    if browser.lower() == "brave":
        key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\BraveSoftware Brave-Browser"
    else:
        key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"

    for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        try:
            reg_key = winreg.OpenKey(hive, key_path)
            version, _ = winreg.QueryValueEx(reg_key, "DisplayVersion")
            return int(version.split(".")[0])
        except Exception:
            continue

    return None


def log(msg: str) -> None:
    """Append a timestamped line to the log area (thread-safe)."""
    ts = datetime.now().strftime("%H:%M:%S")
    root.after(0, _append_log, f"[{ts}] {msg}")


def _append_log(msg: str) -> None:
    txt_log.configure(state="normal")
    txt_log.insert("end", msg + "\n")
    txt_log.see("end")
    txt_log.configure(state="disabled")


def _set_controls(active: bool) -> None:
    """Disable/enable controls while processing. Must run on the main thread."""
    idle_state = "disabled" if active else "normal"
    for w in (btn_browser, btn_planilha, btn_portal, btn_iniciar):
        w.configure(state=idle_state)
    btn_cancelar.configure(state="normal" if active else "disabled")


def _reset_progress(total: int) -> None:
    progress.configure(maximum=total, value=0)
    lbl_status.configure(text=f"0 / {total}")
    lbl_sucesso.configure(text="0")
    lbl_erros.configure(text="0")
    lbl_eta.configure(text="ETA: --")


def _update_progress(i: int, total: int, s: int, e: int, eta: str) -> None:
    pct = int(i / total * 100) if total else 0
    progress.configure(value=i)
    lbl_status.configure(text=f"{i} / {total}  ({pct}%)")
    lbl_sucesso.configure(text=str(s))
    lbl_erros.configure(text=str(e))
    lbl_eta.configure(text=f"ETA: {eta}")


# ── UI callbacks ───────────────────────────────────────────────────────────────

def selecionar_navegador() -> None:
    path = filedialog.askopenfilename(
        title="Selecione o navegador",
        filetypes=[("Executável", "*.exe")],
    )
    if path:
        var_browser.set(path)
        set_key(".env", "BROWSER_PATH", path)
        log(f"Navegador: {os.path.basename(path)} (salvo no .env)")


def selecionar_planilha() -> None:
    global planilha_path
    path = filedialog.askopenfilename(
        title="Selecione a planilha",
        filetypes=[("Excel", "*.xlsx")],
    )
    if path:
        planilha_path = path
        var_planilha.set(os.path.basename(path))
        log(f"Planilha: {os.path.basename(path)}")


def abrir_portal() -> None:
    path = var_browser.get()
    if not path:
        messagebox.showerror("Erro", "Selecione o navegador primeiro.")
        return
    btn_portal.configure(state="disabled", text="Abrindo…")
    log("Abrindo portal no navegador...")
    threading.Thread(target=_open_portal_thread, args=(path,), daemon=True).start()


def _open_portal_thread(path: str) -> None:
    global driver
    nome = os.path.basename(path).lower()
    browser_name = "Brave" if "brave" in nome else "Chrome"
    version = get_browser_version_from_registry(browser_name)
    try:
        options = uc.ChromeOptions()
        options.binary_location = path
        driver = uc.Chrome(options=options, version_main=version)
        driver.get("https://portal.consultatributaria.com.br/")
        root.after(0, _portal_ok)
    except Exception as exc:
        root.after(0, _portal_err, str(exc))


def _portal_ok() -> None:
    btn_portal.configure(state="normal", text="Abrir Portal")
    log("Portal aberto. Faça login e clique em 'Iniciar Programa'.")
    messagebox.showinfo(
        "Login",
        "Faça login no portal.\nDepois clique em 'Iniciar Programa'.",
    )


def _portal_err(msg: str) -> None:
    btn_portal.configure(state="normal", text="Abrir Portal")
    log(f"Erro ao abrir portal: {msg}")
    messagebox.showerror("Erro ao abrir navegador", msg)


def cancelar() -> None:
    cancel_event.set()
    btn_cancelar.configure(state="disabled")
    log("Cancelamento solicitado...")


def inspecionar_ean() -> None:
    ean = var_ean_inspect.get().strip()
    if not ean:
        messagebox.showerror("Erro", "Digite um EAN para inspecionar.")
        return
    if driver is None:
        messagebox.showerror("Erro", "Abra o portal e faça login primeiro.")
        return

    token = _get_token()
    if not token:
        messagebox.showerror("Erro", "Token não encontrado. Verifique se está logado.")
        return

    btn_inspecionar.configure(state="disabled", text="Consultando…")
    threading.Thread(target=_inspecionar_thread, args=(ean, token), daemon=True).start()


def _inspecionar_thread(ean: str, token: str) -> None:
    try:
        data = inspect_ean(ean, token)
        lines = json.dumps(data, indent=2, ensure_ascii=False).splitlines()
        log(f"──── Inspeção EAN {ean} ────")
        for line in lines:
            log(line)
        log("─────────────────────────────────")
    except Exception as exc:
        log(f"Erro ao inspecionar: {exc}")
    finally:
        root.after(0, lambda: btn_inspecionar.configure(state="normal", text="Inspecionar EAN"))


def _get_token() -> str | None:
    """Lê o token de autenticação do localStorage do browser."""
    try:
        local_storage = driver.execute_script("""
            var items = {};
            for (var i = 0; i < localStorage.length; i++) {
                var key = localStorage.key(i);
                items[key] = localStorage.getItem(key);
            }
            return items;
        """)
    except Exception:
        return None

    for key, value in local_storage.items():
        if "auth" in key.lower():
            try:
                data = json.loads(value)
                if "token" in data:
                    return data["token"]
            except Exception:
                pass

    return None


def iniciar_programa() -> None:
    global driver
    if driver is None:
        messagebox.showerror("Erro", "Abra o portal e faça login primeiro.")
        return
    if not planilha_path:
        messagebox.showerror("Erro", "Selecione a planilha XLSX.")
        return

    token = _get_token()
    if not token:
        messagebox.showerror("Erro", "Token não encontrado. Verifique se está logado.")
        return

    cancel_event.clear()
    root.after(0, _set_controls, True)
    threading.Thread(target=_processar_planilha, args=(token,), daemon=True).start()


# ── Processing thread ─────────────────────────────────────────────────────────
def _processar_planilha(token: str) -> None:
    try:
        df = pd.read_excel(planilha_path)
        coluna_ean = "BAR_CODE"

        if coluna_ean not in df.columns:
            root.after(0, messagebox.showerror, "Erro", "Coluna BAR_CODE não encontrada.")
            return

        df[coluna_ean] = (
            df[coluna_ean]
            .astype(str)
            .str.replace(".0", "", regex=False)
            .str.strip()
        )

        eans = df[coluna_ean].dropna().unique()
        total = len(eans)

        root.after(0, _reset_progress, total)
        log(f"Iniciando: {total} EANs únicos encontrados.")

        resultados: dict[str, tuple[str, str]] = {}
        sucesso = 0
        erros = 0
        start = time.monotonic()

        def fetch(ean: str):
            if cancel_event.is_set():
                return ean, "", "", "", "cancelado"
            try:
                product_id = get_product(ean, token)
                if not product_id:
                    return ean, "", "", "", "nao_encontrado"
                nome, ficha, imagem = get_product_info(product_id, token)
                return ean, nome, ficha, imagem, "ok"
            except Exception as exc:
                return ean, "", "", "", f"erro: {exc}"

        with ThreadPoolExecutor(max_workers=8) as executor:
            futures = {executor.submit(fetch, ean): ean for ean in eans}

            for i, future in enumerate(as_completed(futures), 1):
                if cancel_event.is_set():
                    log("Processamento cancelado pelo usuário.")
                    break

                try:
                    ean, nome, ficha, imagem, status = future.result()
                except Exception as exc:
                    log(f"Erro interno: {exc}")
                    erros += 1
                    continue

                resultados[ean] = (nome, ficha, imagem)

                if status == "ok":
                    sucesso += 1
                    log(f"✓  EAN {ean}: {nome}")
                elif status == "nao_encontrado":
                    erros += 1
                    log(f"✗  EAN {ean}: não encontrado")
                elif status.startswith("erro"):
                    erros += 1
                    log(f"✗  EAN {ean}: {status}")
                elif status == "cancelado":
                    break

                elapsed = time.monotonic() - start
                rate = i / elapsed if elapsed > 0 else 0
                remaining = (total - i) / rate if rate > 0 else 0
                eta = f"{int(remaining // 60)}m {int(remaining % 60)}s" if rate > 0 else "--"

                # Capture loop variables for the closure
                _i, _s, _e, _eta = i, sucesso, erros, eta
                root.after(0, _update_progress, _i, total, _s, _e, _eta)

        if not cancel_event.is_set():
            df["NOMEECOMMERCE"] = df[coluna_ean].map(lambda x: resultados.get(x, ("", "", ""))[0])
            df["FICHA TECNICA"] = df[coluna_ean].map(lambda x: resultados.get(x, ("", "", ""))[1])
            df["IMAGEM"]        = df[coluna_ean].map(lambda x: resultados.get(x, ("", "", ""))[2])
            df.to_excel(planilha_path, index=False)
            log(f"Concluído — {sucesso} sucessos, {erros} erros.")

            if var_baixar_imagens.get():
                _baixar_imagens(resultados)

            summary = f"Planilha salva!\n\n✓ Sucesso: {sucesso}\n✗ Erros: {erros}"
            root.after(0, messagebox.showinfo, "Concluído", summary)

    except Exception as exc:
        log(f"Erro crítico: {exc}")
        root.after(0, messagebox.showerror, "Erro", str(exc))
    finally:
        final_text = "Concluído" if not cancel_event.is_set() else "Cancelado"
        root.after(0, lambda t=final_text: lbl_status.configure(text=t))
        root.after(0, _set_controls, False)


def _baixar_imagens(resultados: dict) -> None:
    """Baixa todas as imagens para uma pasta 'imagens/' ao lado da planilha."""
    pasta = os.path.join(os.path.dirname(planilha_path), "imagens")
    os.makedirs(pasta, exist_ok=True)
    log(f"Baixando imagens para: {pasta}")

    pendentes = [
        (ean, dados[2])
        for ean, dados in resultados.items()
        if dados[2]  # só quem tem URL
    ]

    if not pendentes:
        log("Nenhuma imagem disponível para baixar.")
        return

    baixadas = 0
    falhas = 0

    def download(ean: str, url: str):
        ext = os.path.splitext(url.split("?")[0])[1] or ".jpg"
        destino = os.path.join(pasta, f"{ean}{ext}")
        if os.path.exists(destino):
            return "ja_existe"
        try:
            r = requests.get(url, timeout=30)
            r.raise_for_status()
            with open(destino, "wb") as f:
                f.write(r.content)
            return "ok"
        except Exception as exc:
            return f"erro: {exc}"

    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = {executor.submit(download, ean, url): ean for ean, url in pendentes}
        for future in as_completed(futures):
            ean = futures[future]
            status = future.result()
            if status == "ok":
                baixadas += 1
            elif status == "ja_existe":
                baixadas += 1
            else:
                falhas += 1
                log(f"✗  Imagem EAN {ean}: {status}")

    log(f"Imagens: {baixadas} salvas, {falhas} falhas — pasta: {pasta}")


# ── Build UI ───────────────────────────────────────────────────────────────────
root = tbs.Window(themename="darkly")
root.title("Consulta Tributária — Processador de EAN")
root.geometry("1100x720")
root.minsize(900, 600)

var_browser = tk.StringVar(value=os.getenv("BROWSER_PATH", ""))
var_planilha = tk.StringVar()
var_ean_inspect = tk.StringVar()

# Header
frm_header = tbs.Frame(root, padding=(15, 12, 15, 4))
frm_header.pack(fill="x")
tbs.Label(
    frm_header,
    text="Consulta Tributária — Processador de EAN",
    font=("Segoe UI", 15, "bold"),
    bootstyle="light",
).pack(side="left")

tbs.Separator(root).pack(fill="x", padx=15)

# ── Top row: Config | Actions ─────────────────────────────────────────────────
frm_top = tbs.Frame(root, padding=(15, 8, 15, 4))
frm_top.pack(fill="x")
frm_top.columnconfigure(0, weight=3)
frm_top.columnconfigure(1, weight=2)

# Config panel
frm_config = tbs.LabelFrame(frm_top, text=" Configurações ", padding=14, bootstyle="primary")
frm_config.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
frm_config.columnconfigure(0, weight=1)

tbs.Label(frm_config, text="Executável do navegador:").grid(row=0, column=0, sticky="w")
entry_browser = tbs.Entry(frm_config, textvariable=var_browser, state="readonly")
entry_browser.grid(row=1, column=0, sticky="ew", pady=(2, 4))
btn_browser = tbs.Button(
    frm_config, text="Selecionar Navegador",
    command=selecionar_navegador, bootstyle="secondary-outline",
)
btn_browser.grid(row=2, column=0, sticky="w", pady=(0, 12))

tbs.Label(frm_config, text="Planilha XLSX:").grid(row=3, column=0, sticky="w")
entry_planilha = tbs.Entry(frm_config, textvariable=var_planilha, state="readonly")
entry_planilha.grid(row=4, column=0, sticky="ew", pady=(2, 4))
btn_planilha = tbs.Button(
    frm_config, text="Selecionar Planilha",
    command=selecionar_planilha, bootstyle="secondary-outline",
)
btn_planilha.grid(row=5, column=0, sticky="w", pady=(0, 10))

var_baixar_imagens = tk.BooleanVar(value=False)
tbs.Checkbutton(
    frm_config, text="Baixar imagens para pasta local",
    variable=var_baixar_imagens, bootstyle="primary-round-toggle",
).grid(row=6, column=0, sticky="w")

# Actions panel
frm_actions = tbs.LabelFrame(frm_top, text=" Ações ", padding=14, bootstyle="primary")
frm_actions.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
frm_actions.columnconfigure(0, weight=1)

btn_portal = tbs.Button(
    frm_actions, text="1.  Abrir Portal",
    command=abrir_portal, bootstyle="info",
)
btn_portal.pack(fill="x", pady=4)

btn_iniciar = tbs.Button(
    frm_actions, text="2.  Iniciar Programa",
    command=iniciar_programa, bootstyle="success",
)
btn_iniciar.pack(fill="x", pady=4)

tbs.Separator(frm_actions).pack(fill="x", pady=8)

btn_cancelar = tbs.Button(
    frm_actions, text="Cancelar Processamento",
    command=cancelar, bootstyle="danger-outline", state="disabled",
)
btn_cancelar.pack(fill="x", pady=4)

tbs.Separator(frm_actions).pack(fill="x", pady=8)

tbs.Label(frm_actions, text="Inspecionar EAN:", bootstyle="secondary").pack(anchor="w")
entry_ean_inspect = tbs.Entry(frm_actions, textvariable=var_ean_inspect)
entry_ean_inspect.pack(fill="x", pady=(2, 4))
btn_inspecionar = tbs.Button(
    frm_actions, text="Inspecionar EAN",
    command=inspecionar_ean, bootstyle="warning-outline",
)
btn_inspecionar.pack(fill="x")

# ── Progress panel ────────────────────────────────────────────────────────────
frm_prog = tbs.LabelFrame(root, text=" Progresso ", padding=12, bootstyle="primary")
frm_prog.pack(fill="x", padx=15, pady=(0, 4))

progress = tbs.Progressbar(frm_prog, mode="determinate", bootstyle="success-striped")
progress.pack(fill="x", pady=(0, 8))

frm_stats = tbs.Frame(frm_prog)
frm_stats.pack(fill="x")

lbl_status = tbs.Label(frm_stats, text="Aguardando...", font=("Segoe UI", 10))
lbl_status.pack(side="left", padx=(0, 24))

tbs.Label(frm_stats, text="Sucesso:", bootstyle="success").pack(side="left", padx=(0, 4))
lbl_sucesso = tbs.Label(
    frm_stats, text="0", bootstyle="success", font=("Segoe UI", 10, "bold")
)
lbl_sucesso.pack(side="left", padx=(0, 24))

tbs.Label(frm_stats, text="Erros:", bootstyle="danger").pack(side="left", padx=(0, 4))
lbl_erros = tbs.Label(
    frm_stats, text="0", bootstyle="danger", font=("Segoe UI", 10, "bold")
)
lbl_erros.pack(side="left")

lbl_eta = tbs.Label(frm_stats, text="ETA: --", bootstyle="secondary")
lbl_eta.pack(side="right")

# ── Log panel ─────────────────────────────────────────────────────────────────
frm_log = tbs.LabelFrame(root, text=" Log ", padding=10, bootstyle="primary")
frm_log.pack(fill="both", expand=True, padx=15, pady=(0, 10))

scrollbar = tbs.Scrollbar(frm_log, bootstyle="round-secondary")
scrollbar.pack(side="right", fill="y")

txt_log = tk.Text(
    frm_log,
    state="disabled",
    wrap="word",
    font=("Consolas", 9),
    relief="flat",
    cursor="arrow",
    yscrollcommand=scrollbar.set,
)
txt_log.pack(fill="both", expand=True)
scrollbar.configure(command=txt_log.yview)

# Colour tags for ✓/✗ markers
txt_log.tag_configure("ok", foreground="#00bc8c")
txt_log.tag_configure("err", foreground="#e74c3c")

log("Pronto. Selecione o navegador e a planilha para começar.")

root.mainloop()
