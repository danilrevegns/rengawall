# -*- coding: utf-8 -*-
import sys
import argparse
import traceback
from typing import List, Tuple, Optional
from rengawall import run_batch, DEFAULT_CONFIG_PATH

def _gui_main(config_path: str) -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    root = tk.Tk()
    root.title("Renga: пол и стены по контуру помещения")
    root.geometry("720x520")

    cfg_var = tk.StringVar(value=config_path or DEFAULT_CONFIG_PATH)
    log_widget = scrolledtext.ScrolledText(root, height=18, state=tk.DISABLED)

    def log(msg: str) -> None:
        log_widget.configure(state=tk.NORMAL)
        log_widget.insert("end", msg + "\n")
        log_widget.see("end")
        log_widget.configure(state=tk.DISABLED)
        root.update_idletasks()

    def browse_cfg():
        p = filedialog.askopenfilename(
            title="Конфигурация JSON",
            filetypes=[("JSON", "*.json"), ("Все", "*.*")],
        )
        if p:
            cfg_var.set(p)

    def run_clicked():
        path = cfg_var.get().strip()
        if not path:
            messagebox.showerror("Ошибка", "Укажите файл конфигурации.")
            return
        try:
            run_batch(path, [], "all", True, log)
            messagebox.showinfo("Готово", "Обработка завершена (см. журнал).")
        except Exception as ex:
            log(str(ex))
            messagebox.showerror("Ошибка", str(ex))

    def run_selection():
        path = cfg_var.get().strip()
        if not path:
            messagebox.showerror("Ошибка", "Укажите файл конфигурации.")
            return
        try:
            run_batch(path, [], "selection", True, log)
            messagebox.showinfo("Готово", "Обработка выбранных помещений завершена.")
        except Exception as ex:
            log(str(ex))
            messagebox.showerror("Ошибка", str(ex))

    frm = tk.Frame(root)
    frm.pack(fill="x", padx=8, pady=6)
    tk.Label(frm, text="Конфиг JSON:").pack(side="left")
    tk.Entry(frm, textvariable=cfg_var, width=56).pack(
        side="left", padx=4, fill="x", expand=True
    )
    tk.Button(frm, text="…", command=browse_cfg).pack(side="left")

    bf = tk.Frame(root)
    bf.pack(fill="x", padx=8, pady=4)
    tk.Button(bf, text="Все помещения", command=run_clicked).pack(
        side="left", padx=4
    )
    tk.Button(bf, text="Только выбранные в Renga", command=run_selection).pack(
        side="left", padx=4
    )

    log_widget.pack(fill="both", expand=True, padx=8, pady=8)
    tk.Label(
        root,
        text="Перед запуском откройте проект в Renga и задайте три свойства помещений.",
        fg="#444",
    ).pack(pady=(0, 6))

    root.mainloop()


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="Пол и отделочные стены по контуру помещения (Renga COM API) - GUI."
    )
    parser.add_argument(
        "--config",
        "-c",
        default=DEFAULT_CONFIG_PATH,
        help="JSON с property_ids (3 шт.) и rules",
    )
    args = parser.parse_args(argv)

    try:
        _gui_main(args.config)
    except Exception as ex:
        print(ex, file=sys.stderr)
        traceback.print_exc()
        return 1
    return 0

if __name__ == "__main__":
    sys.exit(main())
