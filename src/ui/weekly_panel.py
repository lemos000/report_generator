import os
import threading
from tkinter import filedialog, messagebox

import customtkinter as ctk

from src.core.weekly_reader import WeeklyWorkbookReader, WeeklyData
from src.infrastructure.weekly_exporter import WeeklyPdfExporter


class WeeklyPanel(ctk.CTkScrollableFrame):
    """Formulário do relatório SEMANAL.

    Reaproveita a barra de progresso, o log e o botão principal do App (passado como `app`).
    """

    BUTTON_TEXT = "GERAR RELATÓRIO SEMANAL"

    def __init__(self, app, reader: WeeklyWorkbookReader = None, exporter: WeeklyPdfExporter = None):
        super().__init__(app)
        self.app = app
        self.reader = reader or WeeklyWorkbookReader()
        self.exporter = exporter or WeeklyPdfExporter()
        self._cache = None  # (path, mtime, WeeklyData)
        self._build()

    # ------------------------------------------------------------------ UI

    def _build(self):
        app = self.app
        app._create_section_label(self, "Planilha Modelo (XLSX)")
        self.model_entry = app._create_file_field(
            self, "Planilha *", self.select_model, "Ex: Modelo Hotel - ... - Venda Semanal.xlsx")

        app._create_section_label(self, "Informações do Relatório")
        self.hotel_entry = ctk.CTkEntry(self, placeholder_text="Nome da Propriedade (preenchido pela planilha)", font=("Inter", 12))
        self.hotel_entry.pack(pady=5, padx=10, fill="x")
        self.ref_entry = ctk.CTkEntry(self, placeholder_text="Referência (ex: Semana de 05 a 11 de Outubro de 2026)", font=("Inter", 12))
        self.ref_entry.pack(pady=5, padx=10, fill="x")

        app._create_section_label(self, "Conteúdo")
        self.charts_var = ctk.BooleanVar(value=True)
        ctk.CTkSwitch(
            self, text="Incluir gráficos benchmark no final do PDF", variable=self.charts_var,
            font=("Inter", 12), progress_color="#159F92", button_color="#005151"
        ).pack(pady=5, padx=10, anchor="w")
        self.black_cover_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            self, text="Capa preta", variable=self.black_cover_var,
            font=("Inter", 12), fg_color="#159F92", hover_color="#005151"
        ).pack(pady=5, padx=10, anchor="w")

        app._create_section_label(self, "Destino e Saída")
        self.out_entry = app._create_file_field(self, "Pasta Destino *", self.select_output_folder, "Onde o relatório será salvo")
        self.name_entry = ctk.CTkEntry(self, placeholder_text="Nome do Arquivo Final *", font=("Inter", 12))
        self.name_entry.insert(0, "Relatorio_Semanal")
        self.name_entry.pack(pady=5, padx=10, fill="x")

        downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        if os.path.exists(downloads):
            app._fill_entry(self.out_entry, downloads)

    # ------------------------------------------------------------------ seleção de arquivos

    def select_model(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not path:
            return
        self.app._fill_entry(self.model_entry, path)
        self.app.log("Lendo planilha semanal para preencher os campos...")
        threading.Thread(target=self._autofill, args=(path,), daemon=True).start()

    def select_output_folder(self):
        self.app._fill_entry(self.out_entry, filedialog.askdirectory())

    def _load(self, path: str) -> WeeklyData:
        mtime = os.path.getmtime(path)
        if self._cache and self._cache[0] == path and self._cache[1] == mtime:
            return self._cache[2]
        data = self.reader.read(path)
        self._cache = (path, mtime, data)
        return data

    def _autofill(self, path: str):
        try:
            data = self._load(path)
        except Exception as exc:
            self.after(0, lambda: self.app.log(f"Não foi possível ler a planilha: {exc}", is_error=True))
            return

        def apply():
            if data.hotel:
                self.app._fill_entry(self.hotel_entry, data.hotel)
            if data.reference:
                self.app._fill_entry(self.ref_entry, data.reference)
            if data.hotel and data.week_label:
                self.app._fill_entry(self.name_entry, f"Relatório Semanal - {data.hotel} - {data.week_label}")
            self.app.log("Planilha semanal carregada.")
        self.after(0, apply)

    # ------------------------------------------------------------------ processamento

    def start(self):
        path = self.model_entry.get().strip()
        out = self.out_entry.get().strip()
        name = self.name_entry.get().strip()

        missing = []
        if not path: missing.append("Planilha *")
        if not out: missing.append("Pasta Destino *")
        if not name: missing.append("Nome do Arquivo Final *")
        if missing:
            messagebox.showwarning(
                "Campos Obrigatórios",
                "Por favor, preencha os seguintes campos obrigatórios:\n\n" + "\n".join(f"- {f}" for f in missing))
            return

        app = self.app
        app.process_btn.configure(state="disabled", text="PROCESSANDO...")
        app.progress_bar.set(0)
        app.progress_bar.configure(determinate_speed=0.5)
        app.progress_bar.start()

        threading.Thread(
            target=self._process,
            args=(path, out, name, self.hotel_entry.get().strip(), self.ref_entry.get().strip(), self.charts_var.get(),
                  self.black_cover_var.get()),
            daemon=True,
        ).start()

    def _process(self, path, out, name, hotel, reference, include_charts, black_cover=False):
        app = self.app
        try:
            app.log("Lendo tabelas da planilha semanal...")
            data = self._load(path)
            app.progress_bar.set(0.4)

            property_name = hotel or data.hotel or name
            reference = reference or data.reference

            app.log("Gerando PDF do relatório semanal" + (" (com gráficos benchmark)..." if include_charts else "..."))
            output_path = os.path.join(out, name + ".pdf")
            self.exporter.export(data, output_path, property_name, reference,
                                 include_benchmark=include_charts, black_cover=black_cover)

            app.progress_bar.stop()
            app.progress_bar.set(1.0)
            app.log("Relatório semanal gerado com sucesso!")
            if messagebox.askyesno("Sucesso", "Relatório gerado! Deseja abrir a pasta de destino?"):
                if os.path.exists(out):
                    os.startfile(out)
        except Exception as exc:
            app.log(f"Erro no processamento: {exc}", is_error=True)
            messagebox.showerror("Erro Crítico", f"Ocorreu um erro inesperado:\n{exc}")
            app.progress_bar.stop()
            app.progress_bar.set(0)
        finally:
            app.process_btn.configure(state="normal", text=self.BUTTON_TEXT)
