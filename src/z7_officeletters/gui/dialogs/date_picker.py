"""Date picker dialog.

Shows a tkcalendar ``Calendar`` widget inside a ``CTkToplevel`` so the user
can select a date by clicking rather than typing.  On confirmation the chosen
date is written back to the provided ``StringVar``.

Public exports:
    show_date_picker: Open the date picker and update the date variable.
"""

from __future__ import annotations

from datetime import datetime

import customtkinter as ctk

from z7_officeletters.gui.constants import _C

__all__ = ["show_date_picker"]


def show_date_picker(parent: ctk.CTk, date_var: ctk.StringVar) -> None:
    """Open a calendar popup and write the selected date to ``date_var``.

    The calendar is pre-set to the date currently stored in ``date_var``
    (expected format ``dd/mm/yyyy``).  If the current value cannot be parsed
    the calendar defaults to today.

    Args:
        parent: The root window (used to centre the popup).
        date_var: StringVar whose value is updated on confirmation.
    """
    from tkcalendar import Calendar  # noqa: PLC0415 — optional dependency

    popup = ctk.CTkToplevel(parent)
    popup.title("Selecionar Data")
    popup.resizable(False, False)
    popup.grab_set()
    popup.configure(fg_color=_C["card"])

    try:
        current = datetime.strptime(date_var.get(), "%d/%m/%Y")
    except ValueError:
        current = datetime.now()

    cal = Calendar(
        popup,
        selectmode="day",
        year=current.year,
        month=current.month,
        day=current.day,
        date_pattern="dd/mm/yyyy",
        background=_C["panel"],
        foreground=_C["text"],
        bordercolor=_C["border"],
        headersbackground=_C["bg"],
        headersforeground=_C["accent"],
        selectbackground=_C["accent"],
        selectforeground="#ffffff",
        normalbackground=_C["panel"],
        normalforeground=_C["text"],
        weekendbackground=_C["panel"],
        weekendforeground=_C["dim"],
        othermonthbackground=_C["bg"],
        othermonthforeground=_C["dim"],
        othermonthwebackground=_C["bg"],
        othermonthweforeground=_C["dim"],
        locale="pt_BR",
    )
    cal.pack(padx=14, pady=(14, 6))

    def _confirm() -> None:
        date_var.set(cal.get_date())
        popup.destroy()

    ctk.CTkButton(
        popup, text="Confirmar",
        font=ctk.CTkFont(size=13, weight="bold"),
        height=38, corner_radius=8,
        fg_color=_C["accent"], hover_color=_C["accent2"],
        command=_confirm,
    ).pack(fill="x", padx=14, pady=(0, 14))

    popup.update_idletasks()
    x = parent.winfo_x() + (parent.winfo_width() - popup.winfo_width()) // 2
    y = parent.winfo_y() + (parent.winfo_height() - popup.winfo_height()) // 2
    popup.geometry(f"+{x}+{y}")
