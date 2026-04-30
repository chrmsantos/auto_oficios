"""Gemini prompt template editor dialog.

Displays the current prompt template in a text editor so the user can
customise it.  The edited template is validated for the ``{texto_mocao}``
placeholder before being saved to disk and hot-reloaded into ``core.ai``.

Public exports:
    show_prompt_editor: Open the prompt editor dialog.
"""

from __future__ import annotations

from tkinter import messagebox

import customtkinter as ctk

from z7_officeletters.gui.constants import _C

__all__ = ["show_prompt_editor"]


def show_prompt_editor(parent: ctk.CTk) -> None:
    """Open the Gemini prompt template editor dialog.

    Args:
        parent: The root window (used to centre the dialog).
    """
    import z7_officeletters.core.ai as _ai  # noqa: PLC0415

    dlg = ctk.CTkToplevel(parent)
    dlg.title("Editor de Prompt IA")
    dlg.geometry("700x560")
    dlg.resizable(True, True)
    dlg.grab_set()
    dlg.configure(fg_color=_C["bg"])

    dlg.update_idletasks()
    px, py = parent.winfo_x(), parent.winfo_y()
    pw, ph = parent.winfo_width(), parent.winfo_height()
    dlg.geometry(f"700x560+{px + (pw - 700) // 2}+{py + (ph - 560) // 2}")

    ctk.CTkLabel(
        dlg, text="PROMPT DA IA",
        font=ctk.CTkFont(size=11, weight="bold"),
        text_color=_C["accent"], anchor="w",
    ).pack(fill="x", padx=20, pady=(18, 2))
    ctk.CTkFrame(dlg, height=1, fg_color=_C["border"]).pack(fill="x", padx=20, pady=(0, 8))

    ctk.CTkLabel(
        dlg,
        text="Use {texto_mocao} como marcador onde o texto da moção será inserido.",
        font=ctk.CTkFont(size=11),
        text_color=_C["dim"],
        anchor="w",
    ).pack(fill="x", padx=20, pady=(0, 8))

    editor = ctk.CTkTextbox(
        dlg,
        font=ctk.CTkFont(family="Consolas", size=12),
        fg_color=_C["panel"],
        text_color=_C["text"],
        corner_radius=10,
        wrap="word",
    )
    editor.pack(fill="both", expand=True, padx=20, pady=(0, 8))
    editor.insert("1.0", _ai.PROMPT_TEMPLATE)

    bot = ctk.CTkFrame(dlg, fg_color=_C["card"], corner_radius=0, height=58)
    bot.pack(fill="x", side="bottom")
    bot.pack_propagate(False)
    bot.grid_columnconfigure(0, weight=1)

    def _restore_default() -> None:
        editor.delete("1.0", "end")
        editor.insert("1.0", _ai.PROMPT_TEMPLATE_PADRAO)

    def _save() -> None:
        new_template: str = editor.get("1.0", "end-1c")
        if "{texto_mocao}" not in new_template:
            messagebox.showwarning(
                "Marcador ausente",
                "O prompt deve conter o marcador {texto_mocao} para que o texto da moção seja inserido.",
                parent=dlg,
            )
            return
        try:
            _ai._prompt_file_path().write_text(new_template, encoding="utf-8")  # noqa: SLF001
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Erro ao Salvar", str(exc), parent=dlg)
            return
        _ai.PROMPT_TEMPLATE = new_template  # type: ignore[assignment]
        messagebox.showinfo("Salvo", "Prompt salvo com sucesso!", parent=dlg)
        dlg.destroy()

    ctk.CTkButton(
        bot, text="↺  Padrão",
        font=ctk.CTkFont(size=13), height=38, width=110, corner_radius=8,
        fg_color=_C["panel"], hover_color=_C["border"], text_color=_C["warn"],
        command=_restore_default,
    ).grid(row=0, column=0, sticky="w", padx=(20, 0), pady=10)

    ctk.CTkButton(
        bot, text="Cancelar",
        font=ctk.CTkFont(size=13), height=38, width=110, corner_radius=8,
        fg_color=_C["panel"], hover_color=_C["border"], text_color=_C["dim"],
        command=dlg.destroy,
    ).grid(row=0, column=1, sticky="e", padx=(0, 8), pady=10)

    ctk.CTkButton(
        bot, text="💾  Salvar",
        font=ctk.CTkFont(size=13, weight="bold"), height=38, width=110,
        corner_radius=8, fg_color=_C["accent"], hover_color=_C["accent2"],
        text_color="#ffffff", command=_save,
    ).grid(row=0, column=2, sticky="e", padx=(0, 20), pady=10)
