# Aplikasi Penjadwalan Proyek (CPM)

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)
![Tkinter](https://img.shields.io/badge/Tkinter-GUI-green?logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-Data-blue?logo=pandas&logoColor=white)
![Matplotlib](https://img.shields.io/badge/Matplotlib-Visualisasi-orange?logo=matplotlib&logoColor=white)

Sebuah aplikasi desktop sederhana yang dibuat dengan Python dan Tkinter untuk mengelola dan menganalisis jadwal proyek menggunakan **Critical Path Method (CPM)**.

## Tampilan Aplikasi

![Tampilan Aplikasi](assets/screenshot.png)

*(**Catatan:** Anda harus mengambil screenshot aplikasi Anda, menyimpannya di folder (misal: `assets/screenshot.png`), dan mengganti nama file di atas agar gambar muncul)*

## Fitur Utama

* **Manajemen Kegiatan:** Menambah, mengedit, dan menghapus kegiatan proyek (nama, durasi, dependensi).
* **Kalkulasi CPM:** Menghitung **Early Start (ES)**, **Early Finish (EF)**, **Late Start (LS)**, **Late Finish (LF)**, dan **Slack** untuk setiap kegiatan.
* **Identifikasi Jalur Kritis:** Menandai kegiatan yang berada di jalur kritis secara otomatis (Slack = 0).
* **Visualisasi Interaktif:**
    * Membuat **Network Diagram (AON)** untuk memetakan alur kerja dan dependensi.
    * Membuat **Gantt Chart** untuk visualisasi jadwal dan durasi proyek.
* **Impor & Ekspor Data:**
    * Mengimpor daftar kegiatan dari file Excel.
    * Mengekspor hasil kalkulasi CPM, Network Diagram, dan Gantt Chart ke dalam satu file Excel terpisah.

## Teknologi yang Digunakan

* **Python 3:** Bahasa pemrograman utama.
* **Tkinter** (termasuk `ttk` dan `ttkthemes`): Untuk membangun Antarmuka Pengguna Grafis (GUI) desktop.
* **Pandas:** Untuk manajemen dan manipulasi data kegiatan.
* **NetworkX:** Untuk kalkulasi topologi dan dependensi dalam Network Diagram.
* **Matplotlib:** Untuk membuat dan menampilkan Network Diagram dan Gantt Chart.

## Instalasi dan Cara Menjalankan

1.  **Clone repositori ini:**
    ```bash
    git clone [https://github.com/NAMA_ANDA/NAMA_REPO_ANDA.git](https://github.com/NAMA_ANDA/NAMA_REPO_ANDA.git)
    cd NAMA_REPO_ANDA
    ```

2.  **(Rekomendasi) Buat dan aktifkan *virtual environment*:**
    ```bash
    # Buat venv
    python -m venv venv

    # Aktifkan di Windows
    .\venv\Scripts\activate

    # Aktifkan di macOS/Linux
    source venv/bin/activate
    ```

3.  **Install dependensi yang diperlukan:**
    Buat file bernama `requirements.txt` di folder proyek Anda dan isi dengan:
    ```
    pandas
    networkx
    matplotlib
    ttkthemes
    openpyxl
    xlsxwriter
    ```
    Lalu, jalankan perintah ini di terminal Anda:
    ```bash
    pip install -r requirements.txt
    ```

4.  **Jalankan aplikasi:**
    ```bash
    python scheduler_app.py
    ```

## Format Impor Excel

Jika Anda menggunakan fitur "Impor dari Excel", pastikan file `.xlsx` Anda memiliki nama kolom yang **sama persis** (case-sensitive) seperti berikut:

* `Nama Kegiatan`
* `Durasi`
* `Dependensi`

**Contoh:**

| Nama Kegiatan | Durasi | Dependensi |
| :--- | :--- | :--- |
| A | 5 | |
| B | 3 | A |
| C | 4 | A |
| D | 2 | B, C |