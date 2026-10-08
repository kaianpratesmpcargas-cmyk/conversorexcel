"""
MP CARGAS — Conversor SSW
Aplicativo Desktop para conversão e análise de tabelas do sistema SSW para Excel.
"""
import sys
import os
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from parser_ssw import parse_ssw_table
from excel_export import export_to_excel
from dashboard import DashboardFrame


class MPCargasApp(tk.Tk):
    """
    Aplicação principal MP CARGAS — Conversor SSW.
    """
    def __init__(self):
        super().__init__()

        self.title("MP CARGAS — Conversor SSW")
        self.geometry("1100x720")
        self.minsize(900, 600)

        # Configurar Estilos Gerais
        self.style = ttk.Style(self)
        try:
            self.style.theme_use('clam')
        except Exception:
            pass

        self._configure_styles()

        # Armazenamento de dados
        self.current_records = []

        # Container Principal
        self.main_container = ttk.Frame(self)
        self.main_container.pack(fill="both", expand=True)

        # Montar Tela Inicial
        self.input_frame = None
        self.dashboard_view = None

        self._show_input_screen()

    def _configure_styles(self):
        self.style.configure(".", font=("Helvetica", 10))
        self.style.configure("Header.TLabel", font=("Helvetica", 16, "bold"), foreground="#1A365D")
        self.style.configure("SubHeader.TLabel", font=("Helvetica", 10, "bold"), foreground="#4A5568")
        self.style.configure("Primary.TButton", font=("Helvetica", 10, "bold"))
        self.style.configure("Treeview.Heading", font=("Helvetica", 9, "bold"))
        self.style.configure("Treeview", rowheight=24, font=("Helvetica", 9))

    def _clear_container(self):
        for widget in self.main_container.winfo_children():
            widget.destroy()

    def _show_input_screen(self):
        self._clear_container()
        self.current_records = []

        self.input_frame = ttk.Frame(self.main_container, padding=20)
        self.input_frame.pack(fill="both", expand=True)

        # Cabeçalho da Aplicação
        header_frame = ttk.Frame(self.input_frame)
        header_frame.pack(fill="x", pady=(0, 15))

        lbl_title = ttk.Label(header_frame, text="MP CARGAS", style="Header.TLabel")
        lbl_title.pack(anchor="center")

        lbl_sub = ttk.Label(header_frame, text="CONVERSOR SSW → EXCEL", style="SubHeader.TLabel")
        lbl_sub.pack(anchor="center", pady=(2, 0))

        separator = ttk.Separator(self.input_frame, orient="horizontal")
        separator.pack(fill="x", pady=(5, 15))

        # Instrução e botões auxiliares
        instruct_frame = ttk.Frame(self.input_frame)
        instruct_frame.pack(fill="x", pady=(0, 5))

        ttk.Label(
            instruct_frame,
            text="Cole aqui a tabela copiada do SSW (formato Markdown ou texto delimitado):",
            font=("Helvetica", 10, "bold")
        ).pack(side="left")

        # Botão Colar da Área de Transferência
        btn_paste = tk.Button(
            instruct_frame,
            text="📋 Colar da Área de Transferência",
            command=self._paste_from_clipboard,
            bg="#EDF2F7",
            fg="#2B6CB0",
            font=("Helvetica", 9, "bold"),
            relief="groove",
            bd=1,
            padx=8,
            pady=2,
            cursor="hand2"
        )
        btn_paste.pack(side="right")

        # Grande Área de Texto com Scrollbar
        text_container = ttk.Frame(self.input_frame)
        text_container.pack(fill="both", expand=True, pady=5)

        self.text_area = tk.Text(
            text_container,
            wrap="none",
            font=("Consolas", 10),
            undo=True,
            relief="solid",
            bd=1,
            bg="#FAFAFA"
        )
        scroll_y = ttk.Scrollbar(text_container, orient="vertical", command=self.text_area.yview)
        scroll_x = ttk.Scrollbar(text_container, orient="horizontal", command=self.text_area.xview)
        self.text_area.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.text_area.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        text_container.grid_rowconfigure(0, weight=1)
        text_container.grid_columnconfigure(0, weight=1)

        # Barra de Botões de Ação
        btn_bar = ttk.Frame(self.input_frame, padding=(0, 15, 0, 0))
        btn_bar.pack(fill="x")

        btn_process = tk.Button(
            btn_bar,
            text="✔ PROCESSAR TABELA",
            command=self._process_text,
            bg="#1A365D",
            fg="white",
            font=("Helvetica", 11, "bold"),
            relief="raised",
            bd=2,
            padx=20,
            pady=8,
            cursor="hand2"
        )
        btn_process.pack(side="left", padx=(0, 10))

        btn_clear = tk.Button(
            btn_bar,
            text="✕ LIMPAR",
            command=self._clear_text_area,
            bg="#E2E8F0",
            fg="#4A5568",
            font=("Helvetica", 10, "bold"),
            relief="groove",
            bd=1,
            padx=15,
            pady=8,
            cursor="hand2"
        )
        btn_clear.pack(side="left")

    def _paste_from_clipboard(self):
        try:
            content = self.clipboard_get()
            self.text_area.delete("1.0", tk.END)
            self.text_area.insert("1.0", content)
        except Exception:
            messagebox.showwarning(
                "Área de Transferência Vazia",
                "Não foi possível obter texto da área de transferência. Você também pode clicar no campo e pressionar Ctrl + V.",
                parent=self
            )

    def _clear_text_area(self):
        self.text_area.delete("1.0", tk.END)

    def _process_text(self):
        raw_text = self.text_area.get("1.0", tk.END).strip()
        if not raw_text:
            messagebox.showwarning(
                "Aviso",
                "Por favor, cole o conteúdo da tabela SSW antes de processar.",
                parent=self
            )
            return

        try:
            records, error_msg = parse_ssw_table(raw_text)
            if error_msg:
                messagebox.showerror(
                    "Erro ao Interpretar Dados",
                    error_msg,
                    parent=self
                )
                return

            self.current_records = records
            self._show_dashboard_screen()

        except Exception as e:
            # Nunca apresentar traceback descontrolado ao usuário
            messagebox.showerror(
                "Erro no Processamento",
                f"Ocorreu uma falha inesperada durante a leitura da tabela:\n{str(e)}\n\nVerifique se o formato colado corresponde ao sistema SSW.",
                parent=self
            )

    def _show_dashboard_screen(self):
        self._clear_container()

        # Exibir Tela de Resumo e Tabela
        self.dashboard_view = DashboardFrame(
            self.main_container,
            self.current_records,
            on_export_callback=self._export_excel,
            on_new_import_callback=self._show_input_screen
        )
        self.dashboard_view.pack(fill="both", expand=True)

    def _export_excel(self):
        if not self.current_records:
            messagebox.showwarning("Aviso", "Não há dados para exportar.", parent=self)
            return

        today_str = datetime.now().strftime("%Y-%m-%d")
        default_filename = f"MP_CARGAS_SSW_{today_str}.xlsx"

        file_path = filedialog.asksaveasfilename(
            parent=self,
            title="Salvar Planilha Excel",
            defaultextension=".xlsx",
            initialfile=default_filename,
            filetypes=[("Planilha Excel (*.xlsx)", "*.xlsx")]
        )

        if not file_path:
            return

        try:
            export_to_excel(self.current_records, file_path)
            messagebox.showinfo(
                "Sucesso",
                f"Planilha Excel gerada com sucesso!\n\nArquivo salvo em:\n{file_path}",
                parent=self
            )
        except Exception as e:
            messagebox.showerror(
                "Erro ao Exportar",
                f"Não foi possível salvar o arquivo Excel:\n{str(e)}",
                parent=self
            )


def main():
    app = MPCargasApp()
    app.mainloop()


if __name__ == "__main__":
    main()
