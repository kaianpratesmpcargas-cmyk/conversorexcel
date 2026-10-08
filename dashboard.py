"""
Módulo de visualização: Dashboard de indicadores e tabela interativa com filtros no Tkinter.
"""
import tkinter as tk
from tkinter import ttk
from typing import List, Dict, Any, Callable


class DashboardFrame(ttk.Frame):
    """
    Frame contendo os cards de resumo e a tabela com pesquisa e filtros.
    """
    def __init__(self, parent, records: List[Dict[str, Any]], on_export_callback: Callable, on_new_import_callback: Callable):
        super().__init__(parent)
        self.records = records
        self.filtered_records = list(records)
        self.on_export_callback = on_export_callback
        self.on_new_import_callback = on_new_import_callback

        self._build_ui()
        self._populate_filters()
        self._apply_filters()

    def _build_ui(self):
        # 1. HEADER / CARDS DE RESUMO
        cards_container = ttk.LabelFrame(self, text=" RESULTADO DO PROCESSAMENTO ", padding=(15, 10))
        cards_container.pack(fill="x", padx=15, pady=(10, 5))

        # Métricas
        total_count = len(self.records)
        atrasados_count = sum(1 for r in self.records if r.get('SITUAÇÃO') == 'ATRASADO')
        no_prazo_count = sum(1 for r in self.records if r.get('SITUAÇÃO') == 'NO PRAZO')
        sem_info_count = sum(1 for r in self.records if r.get('SITUAÇÃO') == 'SEM INFORMAÇÃO')
        peso_total = sum(r.get('PESO', 0.0) or 0.0 for r in self.records)
        frete_total = sum(r.get('FRETE', 0.0) or 0.0 for r in self.records)

        # Configurar grid de cards
        cards_frame = ttk.Frame(cards_container)
        cards_frame.pack(fill="x", expand=True)

        for col_idx in range(6):
            cards_frame.columnconfigure(col_idx, weight=1)

        def create_card(parent, col, title, value, bg_color="#FFFFFF", fg_color="#1A365D", value_color="#1A365D"):
            card = tk.Frame(parent, bg=bg_color, relief="solid", bd=1, padx=10, pady=8)
            card.grid(row=0, column=col, padx=4, pady=4, sticky="nsew")

            lbl_title = tk.Label(card, text=title, font=("Helvetica", 9, "bold"), bg=bg_color, fg=fg_color)
            lbl_title.pack(anchor="center")

            lbl_val = tk.Label(card, text=value, font=("Helvetica", 14, "bold"), bg=bg_color, fg=value_color)
            lbl_val.pack(anchor="center", pady=(4, 0))

        create_card(cards_frame, 0, "TOTAL DE CT-RCs", f"{total_count:,}".replace(',', '.'), "#F0F4F8", "#4A5568", "#1A365D")
        create_card(cards_frame, 1, "ATRASADOS", f"{atrasados_count:,}".replace(',', '.'), "#FFF5F5", "#C53030", "#E53E3E")
        create_card(cards_frame, 2, "NO PRAZO", f"{no_prazo_count:,}".replace(',', '.'), "#F0FFF4", "#276749", "#2E8540")
        create_card(cards_frame, 3, "SEM INFORMAÇÃO", f"{sem_info_count:,}".replace(',', '.'), "#F7FAFC", "#718096", "#4A5568")
        create_card(cards_frame, 4, "PESO TOTAL (KG)", f"{peso_total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'), "#EDF2F7", "#2D3748", "#2B6CB0")
        create_card(cards_frame, 5, "FRETE TOTAL (R$)", f"R$ {frete_total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'), "#EDF2F7", "#2D3748", "#2B6CB0")

        # 2. BARRA DE AÇÕES E FILTROS
        action_bar = ttk.Frame(self, padding=(15, 5))
        action_bar.pack(fill="x")

        # Botões de Ação
        btn_export = tk.Button(
            action_bar,
            text=" EXPORTAR EXCEL (.xlsx) ",
            command=self.on_export_callback,
            bg="#2B6CB0",
            fg="white",
            font=("Helvetica", 10, "bold"),
            relief="raised",
            bd=2,
            padx=12,
            pady=4,
            cursor="hand2"
        )
        btn_export.pack(side="left", padx=(0, 10))

        btn_new = tk.Button(
            action_bar,
            text=" NOVA IMPORTAÇÃO ",
            command=self.on_new_import_callback,
            bg="#4A5568",
            fg="white",
            font=("Helvetica", 10, "bold"),
            relief="raised",
            bd=2,
            padx=12,
            pady=4,
            cursor="hand2"
        )
        btn_new.pack(side="left", padx=(0, 20))

        # Controles de Filtro
        filter_box = ttk.Frame(action_bar)
        filter_box.pack(side="right")

        # Busca Textual (CTRC / NF / Geral)
        ttk.Label(filter_box, text="Pesquisar (CTRC/NF/Geral):", font=("Helvetica", 9, "bold")).pack(side="left", padx=(10, 4))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *args: self._apply_filters())
        self.entry_search = ttk.Entry(filter_box, textvariable=self.search_var, width=22)
        self.entry_search.pack(side="left", padx=(0, 15))

        # Filtro de Situação
        ttk.Label(filter_box, text="Situação:", font=("Helvetica", 9, "bold")).pack(side="left", padx=(0, 4))
        self.situacao_var = tk.StringVar(value="TODOS")
        self.cb_situacao = ttk.Combobox(filter_box, textvariable=self.situacao_var, state="readonly", width=16)
        self.cb_situacao.pack(side="left", padx=(0, 15))
        self.cb_situacao.bind("<<ComboboxSelected>>", lambda e: self._apply_filters())

        # Filtro de Destino
        ttk.Label(filter_box, text="Destino:", font=("Helvetica", 9, "bold")).pack(side="left", padx=(0, 4))
        self.destino_var = tk.StringVar(value="TODOS")
        self.cb_destino = ttk.Combobox(filter_box, textvariable=self.destino_var, state="readonly", width=18)
        self.cb_destino.pack(side="left")
        self.cb_destino.bind("<<ComboboxSelected>>", lambda e: self._apply_filters())

        # 3. TABELA DE DADOS (TREEVIEW)
        table_container = ttk.Frame(self, padding=(15, 5))
        table_container.pack(fill="both", expand=True)

        self.columns = [
            ("CTRC", 140, "center"),
            ("DATA EMISSÃO", 120, "center"),
            ("NOTA FISCAL", 120, "center"),
            ("DIAS DE ATRASO", 120, "center"),
            ("SITUAÇÃO", 140, "center"),
        ]

        col_ids = [c[0] for c in self.columns]
        self.tree = ttk.Treeview(table_container, columns=col_ids, show="headings", selectmode="browse")

        # Configurar Cabeçalhos e Colunas
        for col_name, width, anchor in self.columns:
            self.tree.heading(col_name, text=col_name, anchor="center")
            self.tree.column(col_name, width=width, minwidth=60, anchor=anchor)

        # Barras de Rolagem Vertical e Horizontal
        vsb = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        table_container.grid_rowconfigure(0, weight=1)
        table_container.grid_columnconfigure(0, weight=1)

        # Tags de cores para Situação
        self.tree.tag_configure("ATRASADO", background="#FFECEC", foreground="#9C0006")
        self.tree.tag_configure("NO PRAZO", background="#E8F8EA", foreground="#006100")
        self.tree.tag_configure("SEM INFORMAÇÃO", background="#F5F5F5", foreground="#555555")

        # Rodapé com contagem de linhas filtradas
        self.lbl_status = ttk.Label(self, text="", font=("Helvetica", 9), padding=(15, 2))
        self.lbl_status.pack(anchor="w")

    def _populate_filters(self):
        # Situações
        situacoes = ["TODOS", "ATRASADO", "NO PRAZO", "SEM INFORMAÇÃO"]
        self.cb_situacao['values'] = situacoes

        # Destinos distintos
        destinos = sorted(list(set(r.get('DESTINO', '') for r in self.records if r.get('DESTINO'))))
        self.cb_destino['values'] = ["TODOS"] + destinos

    def _apply_filters(self):
        query = self.search_var.get().strip().lower()
        sel_sit = self.situacao_var.get()
        sel_dest = self.destino_var.get()

        # Limpar tabela
        for item in self.tree.get_children():
            self.tree.delete(item)

        filtered = []
        for r in self.records:
            # Filtro de Situação
            sit = r.get('SITUAÇÃO', '')
            if sel_sit != "TODOS" and sit != sel_sit:
                continue

            # Filtro de Destino
            dest = r.get('DESTINO', '')
            if sel_dest != "TODOS" and dest != sel_dest:
                continue

            # Filtro de Busca Textual
            if query:
                text_target = f"{r.get('CTRC', '')} {r.get('N FISCAL', '')} {r.get('REMETENTE', '')} {r.get('DESTINATÁRIO', '')} {r.get('DESTINO', '')}".lower()
                if query not in text_target:
                    continue

            filtered.append(r)

        # Inserir linhas filtradas na Treeview
        for r in filtered:
            peso_val = r.get('PESO', 0.0) or 0.0
            frete_val = r.get('FRETE', 0.0) or 0.0
            dias_val = r.get('DIAS DE ATRASO')
            dias_str = str(dias_val) if dias_val is not None else ""

            peso_str = f"{peso_val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            frete_str = f"R$ {frete_val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            sit = r.get('SITUAÇÃO', '')

            row_values = [
                r.get('CTRC', ''),
                r.get('DATA EMISSÃO', ''),
                r.get('N FISCAL', ''),
                dias_str,
                sit,
            ]
            self.tree.insert("", "end", values=row_values, tags=(sit,))

        self.filtered_records = filtered
        self.lbl_status.config(text=f"Exibindo {len(filtered)} de {len(self.records)} registros.")
