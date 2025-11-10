import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.patches as patches 
from collections import deque, defaultdict
import os
import io
from ttkthemes import ThemedTk

# Mengatur default font untuk Matplotlib agar konsisten
plt.rcParams['font.family'] = 'Segoe UI'

class ProjectSchedulerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Aplikasi Penjadwalan Proyek")
        
        self.root.geometry("1366x768")
        self.root.minsize(1024, 640)

        self.setup_styles()
        self.editing_mode = False
        self.original_activity_name = None

        self.activities_df = pd.DataFrame(columns=['Nama Kegiatan', 'Durasi', 'Dependensi'])

        # PERUBAHAN: Menambah padding utama
        self.main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.main_pane.pack(fill=tk.BOTH, expand=True, padx=15, pady=15) # Padding lebih besar

        self.input_frame = ttk.Frame(self.main_pane, width=400) # Lebar sedikit ditambah
        self.main_pane.add(self.input_frame, weight=1)
        self.create_input_widgets()

        self.output_frame = ttk.Frame(self.main_pane)
        self.main_pane.add(self.output_frame, weight=3)
        self.create_output_widgets()
        
    def setup_styles(self):
        self.style = ttk.Style()
        
        # --- Konfigurasi Font & Padding Global ---
        self.style.configure("TLabel", font=("Segoe UI", 10))
        # PERUBAHAN: Menambah padding internal di semua tombol ttk
        self.style.configure("TButton", font=("Segoe UI", 10), padding=(10, 5))
        
        self.style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), 
                             background="#E8F5E9", foreground="#004D40", padding=(5, 5))
        self.style.configure("TLabelframe.Label", font=("Segoe UI", 11, "bold"), 
                             foreground="#004D40")
        # PERUBAHAN: Menambah padding di tab
        self.style.configure("TNotebook.Tab", font=("Segoe UI", 10), padding=(10, 5))
        self.style.configure("Treeview", rowheight=28) # Membuat baris tabel lebih tinggi


    def create_input_widgets(self):
        self.input_frame.columnconfigure(0, weight=1)
        
        # PERUBAHAN: Menambah padding di LabelFrame
        input_data_frame = ttk.LabelFrame(self.input_frame, text="Input Data Kegiatan", padding=(15, 10))
        input_data_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        input_data_frame.columnconfigure(1, weight=1)

        # PERUBAHAN: Menambah padding di semua label dan entry
        ttk.Label(input_data_frame, text="Nama Kegiatan:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        self.name_entry = ttk.Entry(input_data_frame)
        self.name_entry.grid(row=0, column=1, padx=10, pady=10, sticky=tk.EW)

        ttk.Label(input_data_frame, text="Durasi (hari):").grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
        self.duration_entry = ttk.Entry(input_data_frame)
        self.duration_entry.grid(row=1, column=1, padx=10, pady=10, sticky=tk.EW)

        ttk.Label(input_data_frame, text="Dependensi:").grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)
        self.deps_entry = ttk.Entry(input_data_frame)
        self.deps_entry.grid(row=2, column=1, padx=10, pady=10, sticky=tk.EW)

        self.button_frame = ttk.Frame(input_data_frame)
        self.button_frame.grid(row=3, column=0, columnspan=2, pady=15) # Padding bawah
        self.create_standard_buttons()
        
        file_frame = ttk.LabelFrame(self.input_frame, text="Aksi File", padding=(15, 10))
        file_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=15)
        file_frame.columnconfigure(0, weight=1)
        ttk.Button(file_frame, text="Impor dari Excel", command=self.import_from_excel).grid(sticky="ew", padx=10, pady=5, ipady=5)
        ttk.Button(file_frame, text="Ekspor Hasil ke Excel", command=self.export_results).grid(sticky="ew", padx=10, pady=5, ipady=5)

        process_frame = ttk.LabelFrame(self.input_frame, text="Pengolahan Proyek", padding=(15, 10))
        process_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=15)
        process_frame.columnconfigure(0, weight=1)

        self.run_button = tk.Button(process_frame, text="HITUNG & BUAT DIAGRAM", 
                                     font=("Segoe UI", 11, "bold"), fg="white", 
                                     bg="#28A745",
                                     activebackground="#218838", activeforeground="white",
                                     command=self.run_analysis, relief="flat", borderwidth=0)
        # PERUBAHAN: Padding lebih besar
        self.run_button.grid(sticky="ew", padx=10, pady=10, ipady=10)

    def create_standard_buttons(self):
        for widget in self.button_frame.winfo_children():
            widget.destroy()
        
        # Tombol standar menggunakan ttk.Button
        ttk.Button(self.button_frame, text="Tambah", command=self.add_or_update_activity).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.button_frame, text="Edit", command=self.edit_selected_activity).pack(side=tk.LEFT, padx=5)
        
        # PERUBAHAN: Tombol Hapus menggunakan tk.Button agar bisa diwarnai merah
        self.delete_button = tk.Button(self.button_frame, text="Hapus",
                                       font=("Segoe UI", 10), fg="white", bg="#D9534F",
                                       activebackground="#C9302C", activeforeground="white",
                                       command=self.delete_selected_activity, relief="flat", borderwidth=0, padx=10)
        self.delete_button.pack(side=tk.LEFT, padx=5, ipady=2) # ipady untuk tinggi
        
        ttk.Button(self.button_frame, text="Bersihkan", command=self.clear_all_activities).pack(side=tk.LEFT, padx=5)

    def create_edit_buttons(self):
        for widget in self.button_frame.winfo_children():
            widget.destroy()

        # Tombol Simpan (Primer) pakai tk.Button
        save_button = tk.Button(self.button_frame, text="Simpan Perubahan",
                                 font=("Segoe UI", 10, "bold"), fg="white", 
                                 bg="#28A745",
                                 activebackground="#218838", activeforeground="white",
                                 command=self.add_or_update_activity, relief="flat", borderwidth=0, padx=10)
        save_button.pack(side=tk.LEFT, padx=10, ipady=4) # Padding internal y

        # PERUBAHAN: Tombol Batal pakai ttk.Button standar
        cancel_button = ttk.Button(self.button_frame, text="Batal",
                                   command=self.exit_edit_mode)
        cancel_button.pack(side=tk.LEFT, padx=10)

    def create_output_widgets(self):
        self.output_frame.columnconfigure(0, weight=1)
        self.output_frame.rowconfigure(0, weight=1)
        self.output_frame.rowconfigure(1, weight=3)

        table_frame = ttk.LabelFrame(self.output_frame, text="Tabel Daftar Kegiatan", padding=(10, 5))
        table_frame.grid(row=0, column=0, sticky="nsew", padx=(0,10), pady=10)
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(table_frame, columns=('Nama', 'Durasi', 'Dependensi'), show='headings', style="Treeview")
        self.tree.heading('Nama', text='Nama Kegiatan')
        self.tree.heading('Durasi', text='Durasi')
        self.tree.heading('Dependensi', text='Dependensi')
        
        self.tree.column('Nama', width=250, stretch=True)
        self.tree.column('Durasi', width=80, anchor=tk.CENTER, stretch=False)
        self.tree.column('Dependensi', width=300, stretch=True)
        
        tree_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_scrollbar.grid(row=0, column=1, sticky="ns")

        self.notebook = ttk.Notebook(self.output_frame)
        self.notebook.grid(row=1, column=0, sticky="nsew", padx=(0,10), pady=15)
        
        self.cpm_tab = ttk.Frame(self.notebook)
        self.network_tab = ttk.Frame(self.notebook)
        self.gantt_tab = ttk.Frame(self.notebook)

        self.notebook.add(self.cpm_tab, text=' Hasil Perhitungan CPM ')
        self.notebook.add(self.network_tab, text=' Network Diagram ')
        self.notebook.add(self.gantt_tab, text=' Gantt Chart ')
        
    # ... (Fungsi update_activity_table, add_or_update_activity, ...etc... tidak berubah) ...
    def update_activity_table(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for _, row in self.activities_df.iterrows():
            self.tree.insert('', tk.END, values=list(row))
    
    def add_or_update_activity(self):
        name = self.name_entry.get().strip()
        duration_str = self.duration_entry.get().strip()
        deps = self.deps_entry.get().strip()
        if not name or not duration_str:
            messagebox.showerror("Error", "Nama kegiatan dan durasi tidak boleh kosong.")
            return
        try:
            duration = int(duration_str)
            if duration <= 0: raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Durasi harus berupa angka positif.")
            return
        if self.editing_mode:
            if name != self.original_activity_name and name in self.activities_df['Nama Kegiatan'].values:
                messagebox.showerror("Error", f"Nama kegiatan '{name}' sudah ada.")
                return
            idx = self.activities_df.index[self.activities_df['Nama Kegiatan'] == self.original_activity_name].tolist()[0]
            self.activities_df.at[idx, 'Nama Kegiatan'] = name
            self.activities_df.at[idx, 'Durasi'] = duration
            self.activities_df.at[idx, 'Dependensi'] = deps
            self.exit_edit_mode()
        else:
            if name in self.activities_df['Nama Kegiatan'].values:
                messagebox.showerror("Error", f"Nama kegiatan '{name}' sudah ada.")
                return
            new_activity = {'Nama Kegiatan': name, 'Durasi': duration, 'Dependensi': deps}
            self.activities_df = pd.concat([self.activities_df, pd.DataFrame([new_activity])], ignore_index=True)
        self.update_activity_table()
        self.clear_entries()

    def edit_selected_activity(self):
        selected_items = self.tree.selection()
        if len(selected_items) != 1:
            messagebox.showerror("Error", "Pilih satu kegiatan saja untuk diedit.")
            return
        item_values = self.tree.item(selected_items[0], 'values')
        self.original_activity_name = item_values[0]
        self.clear_entries()
        self.name_entry.insert(0, item_values[0])
        self.duration_entry.insert(0, item_values[1])
        self.deps_entry.insert(0, item_values[2])
        self.enter_edit_mode()

    def enter_edit_mode(self):
        self.editing_mode = True
        self.create_edit_buttons()

    def exit_edit_mode(self):
        self.editing_mode = False
        self.original_activity_name = None
        self.clear_entries()
        self.create_standard_buttons()

    def delete_selected_activity(self):
        if self.editing_mode: self.exit_edit_mode()
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showinfo("Info", "Pilih kegiatan yang akan dihapus.")
            return
        confirm = messagebox.askyesno("Konfirmasi Hapus", f"Anda yakin ingin menghapus {len(selected_items)} kegiatan terpilih?")
        if confirm:
            for item in selected_items:
                item_values = self.tree.item(item, 'values')
                self.activities_df = self.activities_df[self.activities_df['Nama Kegiatan'] != item_values[0]]
            self.update_activity_table()

    def clear_entries(self):
        self.name_entry.delete(0, tk.END)
        self.duration_entry.delete(0, tk.END)
        self.deps_entry.delete(0, tk.END)

    def clear_all_activities(self):
        if self.editing_mode: self.exit_edit_mode()
        if messagebox.askyesno("Konfirmasi", "Anda yakin ingin menghapus semua kegiatan?"):
            self.activities_df = pd.DataFrame(columns=['Nama Kegiatan', 'Durasi', 'Dependensi'])
            self.update_activity_table()
            for tab in [self.cpm_tab, self.network_tab, self.gantt_tab]:
                for widget in tab.winfo_children():
                    widget.destroy()

    def import_from_excel(self):
        if self.editing_mode: self.exit_edit_mode()
        filepath = filedialog.askopenfilename(
            title="Pilih file Excel Anda", filetypes=[("File Excel", "*.xlsx *.xls"), ("Semua File", "*.*")])
        if not filepath: return
        try:
            df = pd.read_excel(filepath)
            df.dropna(how='all', inplace=True)
            required_cols = ['Nama Kegiatan', 'Durasi', 'Dependensi']
            if not all(col in df.columns for col in required_cols):
                messagebox.showerror("Error", f"File Excel harus memiliki kolom: {', '.join(required_cols)}")
                return
            df = df[required_cols].copy()
            df['Dependensi'] = df['Dependensi'].fillna('').astype(str).str.replace('.0', '', regex=False)
            df['Nama Kegiatan'] = df['Nama Kegiatan'].str.strip()
            if df['Nama Kegiatan'].duplicated().any():
                dupes = df[df['Nama Kegiatan'].duplicated()]['Nama Kegiatan'].tolist()
                messagebox.showerror("Error Impor", f"Ada nama kegiatan yang duplikat: {', '.join(dupes)}.")
                return
            self.activities_df = df
            self.update_activity_table()
            messagebox.showinfo("Sukses", f"{len(df)} kegiatan berhasil diimpor.")
        except Exception as e:
            messagebox.showerror("Error Impor", f"Gagal membaca file Excel: {e}")

    def run_analysis(self):
        if self.editing_mode: self.exit_edit_mode()
        if self.activities_df.empty:
            messagebox.showerror("Error", "Tidak ada kegiatan untuk dianalisis.")
            return
        try:
            self.results_df, self.critical_path_names = self.calculate_cpm()
            self.display_cpm_results()
            self.draw_network_diagram()
            self.draw_gantt_chart()
            messagebox.showinfo("Sukses", "Analisis CPM dan pembuatan diagram selesai.")
            self.notebook.select(self.network_tab)
        except Exception as e:
            messagebox.showerror("Analisis Gagal", f"Terjadi kesalahan: {e}")

    def calculate_cpm(self):
        df = self.activities_df.copy()
        df.set_index('Nama Kegiatan', inplace=True)
        all_activities = set(df.index)
        predecessors = {name: [] for name in all_activities}
        successors = {name: [] for name in all_activities}
        for name, row in df.iterrows():
            deps_str = str(row.get('Dependensi', ''))
            if deps_str and deps_str.strip():
                deps = [d.strip() for d in deps_str.split(',')]
                for dep_name in deps:
                    if dep_name not in all_activities:
                        raise ValueError(f"Dependensi '{dep_name}' untuk '{name}' tidak valid.")
                    predecessors[name].append(dep_name)
                    successors[dep_name].append(name)
        in_degree = {u: len(predecessors[u]) for u in all_activities}
        queue = deque([u for u, deg in in_degree.items() if deg == 0])
        topo_order = []
        while queue:
            u = queue.popleft()
            topo_order.append(u)
            for v in successors[u]:
                in_degree[v] -= 1
                if in_degree[v] == 0:
                    queue.append(v)
        if len(topo_order) != len(all_activities):
            raise ValueError("Terdeteksi dependensi sirkular.")
        df['ES'] = 0
        df['EF'] = 0
        for activity in topo_order:
            es = max([df.loc[p, 'EF'] for p in predecessors[activity]], default=0)
            df.loc[activity, 'ES'] = es
            df.loc[activity, 'EF'] = es + df.loc[activity, 'Durasi']
        project_finish_time = df['EF'].max()
        df['LF'] = project_finish_time
        df['LS'] = 0
        for activity in reversed(topo_order):
            lf = min([df.loc[s, 'LS'] for s in successors[activity]], default=project_finish_time)
            df.loc[activity, 'LF'] = lf
            df.loc[activity, 'LS'] = lf - df.loc[activity, 'Durasi']
        df['Slack'] = df['LS'] - df['ES']
        critical_path_names = df[df['Slack'] == 0].index.tolist()
        return df.reset_index(), critical_path_names

    def display_cpm_results(self):
        for widget in self.cpm_tab.winfo_children(): widget.destroy()
        if not hasattr(self, 'results_df'): return
        
        # PERUBAHAN: Padding di frame tab
        frame = ttk.Frame(self.cpm_tab, padding=15)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        
        cols = ['ID Kegiatan', 'Durasi', 'ES', 'EF', 'LS', 'LF', 'Slack']
        df_cols = ['Nama Kegiatan', 'Durasi', 'ES', 'EF', 'LS', 'LF', 'Slack']
        tree = ttk.Treeview(frame, columns=cols, show='headings')
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor=tk.CENTER, stretch=True)
        tree.column('ID Kegiatan', width=150, anchor=tk.W, stretch=True)
        
        # PERUBAHAN: Menambah tag untuk baris non-kritis
        tree.tag_configure('normal', background='white')
        tree.tag_configure('critical', background='#FFEBEE', foreground='#B71C1C', font=('Segoe UI', 9, 'bold'))
        
        for i, row in self.results_df.iterrows():
            values = [int(row[c]) if c != 'Nama Kegiatan' else row[c] for c in df_cols]
            # PERUBAHAN: Memberi tag 'critical' atau 'normal'
            tags = ('critical',) if row['Nama Kegiatan'] in self.critical_path_names else ('normal',)
            tree.insert('', tk.END, values=values, tags=tags)
        
        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=0, column=1, sticky="ns")

    def draw_diagram_on_canvas(self, fig, master_frame):
        for widget in master_frame.winfo_children():
            widget.destroy()
        canvas = FigureCanvasTkAgg(fig, master=master_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def create_diagram_legend(self, parent_frame):
        legend_frame = ttk.LabelFrame(parent_frame, text="Petunjuk Kotak Kegiatan")
        # PERUBAHAN: Padding
        legend_frame.pack(pady=(10, 15), padx=10, fill=tk.X) 
        fig = plt.figure(figsize=(3.2, 1.2))
        ax = fig.add_subplot(111)
        label_text = (f"{'ES':^5}|{'Durasi':^6}|{'EF':^5}\n"
                      f"{'ID Kegiatan':^18}\n"
                      f"{'LS':^5}|{'Slack':^6}|{'LF':^5}")
        ax.text(0.5, 0.5, label_text, ha='center', va='center', fontsize=10,
                bbox=dict(boxstyle="round,pad=0.5", 
                          fc="#E8F5E9", ec="#1E8E3E", lw=1.5),
                # PERUBAHAN: Font konsisten
                fontfamily='Segoe UI', fontweight='bold') 
        ax.axis('off')
        fig.tight_layout(pad=0.1)
        canvas = FigureCanvasTkAgg(fig, master=legend_frame)
        canvas.draw()
        canvas.get_tk_widget().pack()

    def draw_network_diagram(self):
        for widget in self.network_tab.winfo_children():
            widget.destroy()
        if not hasattr(self, 'results_df'): return
        container = ttk.Frame(self.network_tab, padding=10)
        container.pack(fill=tk.BOTH, expand=True)
        self.create_diagram_legend(container)
        diagram_frame = ttk.Frame(container)
        diagram_frame.pack(fill=tk.BOTH, expand=True)

        fig, ax = plt.subplots(figsize=(14, 9))
        G = nx.DiGraph()
        all_activities = set(self.activities_df['Nama Kegiatan'])
        predecessors = {name: [] for name in all_activities}
        for _, row in self.activities_df.iterrows():
            name = row['Nama Kegiatan']
            G.add_node(name)
            deps_str = str(row.get('Dependensi', ''))
            if deps_str and deps_str.strip():
                deps = [d.strip() for d in deps_str.split(',')]
                for dep_name in deps:
                    predecessors[name].append(dep_name)
                    G.add_edge(dep_name, name)
        layers = {node: 0 for node in G.nodes()}
        for node in nx.topological_sort(G):
            if predecessors[node]:
                layers[node] = max(layers[p] for p in predecessors[node]) + 1
        pos = {}
        layer_counts = defaultdict(int)
        for node in G.nodes():
            layer_counts[layers[node]] += 1
        layer_y_offsets = defaultdict(int)
        sorted_nodes_by_layer = sorted(G.nodes(), key=lambda n: (layers[n], n))
        for node in sorted_nodes_by_layer:
            layer = layers[node]
            total_in_layer = layer_counts[layer]
            start_y = (total_in_layer - 1) / 2.0
            pos[node] = (layer, start_y - layer_y_offsets[layer])
            layer_y_offsets[layer] += 1
        
        bbox_props_normal = dict(boxstyle="round,pad=0.7", fc="#E8F5E9", ec="#1E8E3E", lw=1.5)
        bbox_props_critical = dict(boxstyle="round,pad=0.7", fc="#FFEBEE", ec="#B71C1C", lw=2.5)
        
        for node, (x, y) in pos.items():
            row = self.results_df[self.results_df['Nama Kegiatan'] == node].iloc[0]
            label = (f"{int(row['ES']):^5}|{int(row['Durasi']):^6}|{int(row['EF']):^5}\n"
                     f"{node:^18}\n"
                     f"{int(row['LS']):^5}|{int(row['Slack']):^6}|{int(row['LF']):^5}")
            props = bbox_props_critical if node in self.critical_path_names else bbox_props_normal
            ax.text(x, y, label, ha='center', va='center',
                    fontsize=9, bbox=props, 
                    # PERUBAHAN: Font konsisten
                    fontfamily='Monospace', zorder=10) # Monospace lebih baik di sini

        for u, v in G.edges():
            pos_u = pos[u]
            pos_v = pos[v]
            arrow = patches.FancyArrowPatch(
                (pos_u[0] + 0.35, pos_u[1]), 
                (pos_v[0] - 0.35, pos_v[1]),
                arrowstyle='-|>',
                color='gray',
                mutation_scale=20,
                connectionstyle='arc3,rad=0.15',
                zorder=1
            )
            ax.add_patch(arrow)

        max_layer = max(layers.values()) if layers else 0
        ax.set_xlim(-0.5, max_layer + 0.5)
        max_nodes_in_layer = max(layer_counts.values()) if layer_counts else 1
        ax.set_ylim(-(max_nodes_in_layer / 2) - 0.5, (max_nodes_in_layer / 2) + 0.5)
        # PERUBAHAN: Font konsisten
        ax.set_title("Network Diagram Proyek", fontsize=16, fontweight='bold', fontfamily='Segoe UI')
        ax.axis('off')
        plt.tight_layout()
        self.network_fig = fig
        self.draw_diagram_on_canvas(fig, diagram_frame)

    def draw_gantt_chart(self):
        for widget in self.gantt_tab.winfo_children(): widget.destroy()
        if not hasattr(self, 'results_df'): return
        fig, ax = plt.subplots(figsize=(12, 8))
        df = self.results_df.sort_values(by='ES', ascending=False).reset_index(drop=True)
        
        colors = ['#D83B01' if row['Nama Kegiatan'] in self.critical_path_names else '#28A745' for _, row in df.iterrows()]
        
        ax.barh(df['Nama Kegiatan'], df['Durasi'], left=df['ES'], color=colors, edgecolor='black', zorder=3)
        
        # PERUBAHAN: Font konsisten
        ax.set_xlabel("Durasi Proyek (Hari)", fontweight='bold', fontfamily='Segoe UI')
        ax.set_ylabel("Nama Kegiatan", fontweight='bold', fontfamily='Segoe UI')
        ax.set_title("Gantt Chart Proyek", fontsize=15, fontweight='bold', fontfamily='Segoe UI')
        ax.grid(True, which='major', axis='x', linestyle='--', linewidth=0.5, zorder=0)
        ax.invert_yaxis()
        
        # Mengatur font ticks (label sumbu)
        plt.setp(ax.get_xticklabels(), fontfamily='Segoe UI')
        plt.setp(ax.get_yticklabels(), fontfamily='Segoe UI')

        for index, row in df.iterrows():
            ax.text(row['ES'] + row['Durasi']/2, index, str(int(row['Durasi'])), 
                    va='center', ha='center', color='white', 
                    fontweight='bold', fontsize=10, fontfamily='Segoe UI')
        
        plt.tight_layout()
        self.gantt_fig = fig
        self.draw_diagram_on_canvas(fig, self.gantt_tab)

    def export_results(self):
        if not hasattr(self, 'results_df'):
            messagebox.showerror("Error", "Belum ada hasil untuk diekspor.")
            return
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("File Excel", "*.xlsx")], title="Simpan Hasil Analisis")
        if not filepath: return
        try:
            network_image_stream = io.BytesIO()
            self.network_fig.savefig(network_image_stream, format='png', dpi=300, bbox_inches='tight')
            network_image_stream.seek(0)
            gantt_image_stream = io.BytesIO()
            self.gantt_fig.savefig(gantt_image_stream, format='png', dpi=300, bbox_inches='tight')
            gantt_image_stream.seek(0)
            with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
                export_df = self.results_df.copy()
                cols_to_int = ['Durasi', 'ES', 'EF', 'LS', 'LF', 'Slack']
                for col in cols_to_int:
                    export_df[col] = export_df[col].astype(int)
                export_df.to_excel(writer, sheet_name='Hasil CPM', index=False)
                workbook = writer.book
                worksheet_network = workbook.add_worksheet('Network Diagram')
                worksheet_network.insert_image('B2', 'network_diagram.png', {'image_data': network_image_stream, 'x_scale': 0.5, 'y_scale': 0.5})
                worksheet_gantt = workbook.add_worksheet('Gantt Chart')
                worksheet_gantt.insert_image('B2', 'gantt_chart.png', {'image_data': gantt_image_stream, 'x_scale': 0.5, 'y_scale': 0.5})
            messagebox.showinfo("Sukses", f"Hasil berhasil diekspor ke {filepath}")
        except Exception as e:
            messagebox.showerror("Error Ekspor", f"Gagal mengekspor file: {e}")

if __name__ == "__main__":
    root = ThemedTk(theme="arc")
    app = ProjectSchedulerApp(root)
    root.mainloop()