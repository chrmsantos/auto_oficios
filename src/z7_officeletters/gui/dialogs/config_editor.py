"""config.json editor dialog.

Provides a scrollable form to edit the mayor information, the list of
councillor-authors (name, sigla, gender), and the list of drafter names.
On save, the changes are written to ``config.json`` and the runtime state
is hot-reloaded via ``core.config.reload_config`` and
``core.authors.rebuild_tables``.

Public exports:
    show_config_editor: Open the configuration editor dialog.
"""

from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path
from tkinter import messagebox
from typing import Any

import customtkinter as ctk

from z7_officeletters.gui.constants import _C

__all__ = ["show_config_editor"]


def show_config_editor(parent: ctk.CTk, on_save: "Callable[[], None]") -> None:  # type: ignore[name-defined]  # noqa: F821
    """Open the configuration editor dialog.

    Args:
        parent: The root window (used to centre the dialog).
        on_save: Callback called after a successful save so the main window
            can refresh UI elements that depend on the config (e.g. the
            drafter combo-box).
    """
    from typing import Callable  # noqa: PLC0415

    if getattr(sys, "frozen", False):
        cfg_path = Path(sys.executable).parent / "config.json"
        if not cfg_path.exists():
            shutil.copy2(Path(getattr(sys, "_MEIPASS", "")) / "config.json", cfg_path)
    else:
        cfg_path = Path(__file__).parent.parent.parent.parent.parent / "config.json"

    try:
        with cfg_path.open(encoding="utf-8") as fh:
            cfg: dict[str, Any] = json.load(fh)
    except Exception as exc:  # noqa: BLE001
        messagebox.showerror("Erro", f"Não foi possível ler config.json:\n{exc}")
        return

    dlg = ctk.CTkToplevel(parent)
    dlg.title("Configurações")
    dlg.geometry("560x660")
    dlg.resizable(False, True)
    dlg.grab_set()
    dlg.configure(fg_color=_C["bg"])

    dlg.update_idletasks()
    px, py = parent.winfo_x(), parent.winfo_y()
    pw, ph = parent.winfo_width(), parent.winfo_height()
    dlg.geometry(f"560x660+{px + (pw - 560) // 2}+{py + (ph - 660) // 2}")

    # Bottom action bar
    bot = ctk.CTkFrame(dlg, fg_color=_C["card"], corner_radius=0, height=58)
    bot.pack(fill="x", side="bottom")
    bot.pack_propagate(False)
    bot.grid_columnconfigure(0, weight=1)

    scroll = ctk.CTkScrollableFrame(dlg, fg_color=_C["bg"], corner_radius=0)
    scroll.pack(fill="both", expand=True)
    scroll.grid_columnconfigure(0, weight=1)

    def _sec(text: str, row: int) -> None:
        ctk.CTkLabel(
            scroll, text=text,
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=_C["accent"], anchor="w",
        ).grid(row=row, column=0, sticky="w", padx=20, pady=(18, 2))
        ctk.CTkFrame(scroll, height=1, fg_color=_C["border"]).grid(
            row=row + 1, column=0, sticky="ew", padx=20, pady=(0, 10))

    def _lbl(text: str, row: int) -> None:
        ctk.CTkLabel(
            scroll, text=text,
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=_C["text"], anchor="w",
        ).grid(row=row, column=0, sticky="w", padx=20, pady=(0, 4))

    # ── PREFEITO ──────────────────────────────────────────────────────────────
    _sec("PREFEITO", 0)
    _lbl("Nome", 2)
    pref_nome_var = ctk.StringVar(value=cfg.get("prefeito", {}).get("nome", ""))
    ctk.CTkEntry(
        scroll, textvariable=pref_nome_var,
        font=ctk.CTkFont(size=13), height=38,
    ).grid(row=3, column=0, sticky="ew", padx=20, pady=(0, 12))

    _lbl("Endereço  (use \\n para quebra de linha)", 4)
    pref_end_var = ctk.StringVar(
        value=cfg.get("prefeito", {}).get("endereco", "").replace("\n", "\\n")
    )
    ctk.CTkEntry(
        scroll, textvariable=pref_end_var,
        font=ctk.CTkFont(size=13), height=38,
    ).grid(row=5, column=0, sticky="ew", padx=20, pady=(0, 12))

    # ── AUTORES ───────────────────────────────────────────────────────────────
    _sec("AUTORES", 6)

    authors_wrap = ctk.CTkFrame(scroll, fg_color="transparent")
    authors_wrap.grid(row=8, column=0, sticky="ew", padx=20)

    hdr = ctk.CTkFrame(authors_wrap, fg_color="transparent")
    hdr.pack(fill="x", pady=(0, 4))
    ctk.CTkLabel(hdr, text="Nome", font=ctk.CTkFont(size=11, weight="bold"),
                 text_color=_C["dim"]).pack(side="left")
    ctk.CTkLabel(hdr, text="Sigla", font=ctk.CTkFont(size=11, weight="bold"),
                 text_color=_C["dim"], width=76).pack(side="left", padx=(6, 0))
    ctk.CTkLabel(hdr, text="♀", font=ctk.CTkFont(size=11, weight="bold"),
                 text_color=_C["dim"], width=36).pack(side="left", padx=(6, 0))

    rows_frame = ctk.CTkFrame(authors_wrap, fg_color="transparent")
    rows_frame.pack(fill="x")

    feminine_set: set[str] = set(cfg.get("vereadores_feminino", []))
    author_rows: list[dict[str, Any]] = []

    def _add_author_row(nome: str = "", sigla: str = "", fem: bool = False, focus: bool = False) -> None:
        rf = ctk.CTkFrame(rows_frame, fg_color="transparent")
        rf.pack(fill="x", pady=2)
        nv: ctk.StringVar = ctk.StringVar(value=nome)
        sv: ctk.StringVar = ctk.StringVar(value=sigla)
        fv: ctk.BooleanVar = ctk.BooleanVar(value=fem)
        ne = ctk.CTkEntry(rf, textvariable=nv, font=ctk.CTkFont(size=12), height=34)
        ne.pack(side="left", fill="x", expand=True)
        ctk.CTkEntry(rf, textvariable=sv, width=76,
                     font=ctk.CTkFont(size=12), height=34).pack(side="left", padx=(6, 0))
        ctk.CTkCheckBox(rf, text="", variable=fv, width=36, height=34,
                        checkbox_width=18, checkbox_height=18).pack(side="left", padx=(6, 0))
        rd: dict[str, Any] = {"nv": nv, "sv": sv, "fv": fv}

        def _del() -> None:
            author_rows.remove(rd)
            rf.destroy()

        ctk.CTkButton(
            rf, text="✕", width=28, height=34,
            font=ctk.CTkFont(size=11),
            fg_color="transparent", hover_color=_C["error"],
            text_color=_C["dim"], border_width=0,
            command=_del,
        ).pack(side="left", padx=(4, 0))
        author_rows.append(rd)
        if focus:
            ne.focus_set()

    for _n, _s in cfg.get("autores", {}).items():
        _add_author_row(_n, _s, _n in feminine_set)

    ctk.CTkButton(
        scroll, text="＋  Adicionar Autor",
        font=ctk.CTkFont(size=12), height=32, corner_radius=8,
        fg_color=_C["panel"], hover_color=_C["border"],
        text_color=_C["accent"], border_width=1, border_color=_C["border"],
        command=lambda: _add_author_row(focus=True),
    ).grid(row=9, column=0, sticky="w", padx=20, pady=(8, 18))

    # ── REDATORES ─────────────────────────────────────────────────────────────
    _sec("REDATORES", 10)

    redatores_wrap = ctk.CTkFrame(scroll, fg_color="transparent")
    redatores_wrap.grid(row=12, column=0, sticky="ew", padx=20)

    rhdr = ctk.CTkFrame(redatores_wrap, fg_color="transparent")
    rhdr.pack(fill="x", pady=(0, 4))
    ctk.CTkLabel(rhdr, text="Nome", font=ctk.CTkFont(size=11, weight="bold"),
                 text_color=_C["dim"]).pack(side="left")
    ctk.CTkLabel(rhdr, text="Sigla", font=ctk.CTkFont(size=11, weight="bold"),
                 text_color=_C["dim"], width=76).pack(side="left", padx=(6, 0))

    rrows_frame = ctk.CTkFrame(redatores_wrap, fg_color="transparent")
    rrows_frame.pack(fill="x")

    redator_rows: list[dict[str, Any]] = []

    def _add_redator_row(nome: str = "", sigla: str = "", focus: bool = False) -> None:
        rf = ctk.CTkFrame(rrows_frame, fg_color="transparent")
        rf.pack(fill="x", pady=2)
        nv2: ctk.StringVar = ctk.StringVar(value=nome)
        sv2: ctk.StringVar = ctk.StringVar(value=sigla)
        ne2 = ctk.CTkEntry(rf, textvariable=nv2, font=ctk.CTkFont(size=12), height=34)
        ne2.pack(side="left", fill="x", expand=True)
        ctk.CTkEntry(rf, textvariable=sv2, width=76,
                     font=ctk.CTkFont(size=12), height=34).pack(side="left", padx=(6, 0))
        rd2: dict[str, Any] = {"nv": nv2, "sv": sv2}

        def _del_r() -> None:
            redator_rows.remove(rd2)
            rf.destroy()

        ctk.CTkButton(
            rf, text="✕", width=28, height=34,
            font=ctk.CTkFont(size=11),
            fg_color="transparent", hover_color=_C["error"],
            text_color=_C["dim"], border_width=0,
            command=_del_r,
        ).pack(side="left", padx=(4, 0))
        redator_rows.append(rd2)
        if focus:
            ne2.focus_set()

    for _rn, _rs in cfg.get("redatores", {}).items():
        _add_redator_row(_rn, _rs)

    ctk.CTkButton(
        scroll, text="＋  Adicionar Redator",
        font=ctk.CTkFont(size=12), height=32, corner_radius=8,
        fg_color=_C["panel"], hover_color=_C["border"],
        text_color=_C["accent"], border_width=1, border_color=_C["border"],
        command=lambda: _add_redator_row(focus=True),
    ).grid(row=13, column=0, sticky="w", padx=20, pady=(8, 18))

    # ── Save / Cancel ─────────────────────────────────────────────────────────
    def _save() -> None:
        import z7_officeletters.core.authors as _authors  # noqa: PLC0415
        import z7_officeletters.core.config as _config  # noqa: PLC0415

        autores: dict[str, str] = {}
        new_fem: list[str] = []
        for rd in author_rows:
            n: str = rd["nv"].get().strip()
            s: str = rd["sv"].get().strip()
            if n and s:
                autores[n] = s
                if rd["fv"].get():
                    new_fem.append(n)

        redatores: dict[str, str] = {}
        for rd2 in redator_rows:
            n2: str = rd2["nv"].get().strip()
            s2: str = rd2["sv"].get().strip()
            if n2 and s2:
                redatores[n2] = s2

        new_cfg: dict[str, Any] = {
            "_comentario": cfg.get("_comentario", ""),
            "prefeito": {
                "nome": pref_nome_var.get().strip(),
                "endereco": pref_end_var.get().strip().replace("\\n", "\n"),
            },
            "autores": autores,
            "vereadores_feminino": new_fem,
            "redatores": redatores,
        }

        try:
            with cfg_path.open("w", encoding="utf-8") as fh:
                json.dump(new_cfg, fh, ensure_ascii=False, indent=2)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Erro ao Salvar", str(exc), parent=dlg)
            return

        # Hot-reload runtime state
        _config.reload_config()
        _authors.rebuild_tables()
        on_save()

        messagebox.showinfo("Salvo", "Configurações salvas!", parent=dlg)
        dlg.destroy()

    ctk.CTkButton(
        bot, text="Cancelar",
        font=ctk.CTkFont(size=13), height=38, width=110, corner_radius=8,
        fg_color=_C["panel"], hover_color=_C["border"], text_color=_C["dim"],
        command=dlg.destroy,
    ).grid(row=0, column=0, sticky="e", padx=(0, 8), pady=10)

    ctk.CTkButton(
        bot, text="💾  Salvar",
        font=ctk.CTkFont(size=13, weight="bold"), height=38, width=110,
        corner_radius=8, fg_color=_C["accent"], hover_color=_C["accent2"],
        text_color="#ffffff", command=_save,
    ).grid(row=0, column=1, sticky="e", padx=(0, 20), pady=10)
