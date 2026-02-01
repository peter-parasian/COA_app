## 📋 Ringkasan (Overview)
Aplikasi desktop yang dirancang khusus untuk memproses dataset besar dari catatan produksi Copper Busbar & Strip, menstandarisasi data dimensi, dan secara otomatis menghasilkan dokumen **Certificate of Analysis (COA)**.

Dibangun dengan filosofi **"Pragmatic Human-Readable Code"** dan dioptimalkan secara khusus untuk lingkungan perangkat keras terbatas (RAM 4GB / SSD), dengan memprioritaskan penghematan memori dan kecepatan baca-tulis (I/O).

![Tech Stack](https://img.shields.io/badge/Tech-C%23_.NET_8-blue)
![Platform](https://img.shields.io/badge/Platform-WPF_Desktop-lightgrey)
![Database](https://img.shields.io/badge/Database-SQLite_WAL-green)

## 🚀 Fitur Utama

### 1. ETL Berkinerja Tinggi (Extract, Transform, Load)
* **Pemrosesan Paralel:** Membaca ribuan file Excel produksi lawas secara bersamaan menggunakan `System.Threading.Tasks.Parallel` untuk kecepatan maksimal.
* **Optimasi Memori:** Menggunakan pembacaan streaming (`ExcelDataReader`) dan pemrosesan batch (Transaksi SQL) untuk menjaga penggunaan RAM tetap rendah, bahkan saat memproses ribuan file.
* **Smart Parsing:** Secara otomatis memperbaiki format tanggal yang tidak konsisten dan string dimensi (contoh: "3x10", "3.00 x 10.00") menjadi data terstruktur yang siap pakai.

### 2. Optimasi Database SQLite
* **Mode WAL (Write-Ahead Logging):** Dikonfigurasi untuk operasi baca/tulis konkurensi tinggi yang memanfaatkan kecepatan SSD.
* **Indexing Efisien:** Indeks kustom pada kolom `Size_mm` dan `Prod_date` memungkinkan pencarian data dalam hitungan detik, meskipun data mencakup rentang waktu bertahun-tahun.
* **Pencocokan Cerdas (Zero-Inference):** Algoritma khusus untuk mencocokkan "Nomor Batch" dari lembar log yang terpisah berdasarkan stempel waktu produksi.

### 3. Pelaporan Otomatis (COA)
* **Template Engine:** Menghasilkan file Excel COA resmi menggunakan library `ClosedXML`, menyuntikkan data ke dalam template perusahaan yang sudah diformat sebelumnya.
* **Validasi Standar JIS:** Secara otomatis menghitung toleransi (Ketebalan/Lebar) berdasarkan logika standar **JIS (Japanese Industrial Standards)**.
* **Visual QC:** Secara otomatis menandai nilai yang di luar spesifikasi dengan warna merah langsung di dalam laporan Excel yang dihasilkan.

## 🛠️ Teknologi yang Digunakan

* **Bahasa:** C# 8.0+ (.NET 8 Windows)
* **Framework UI:** WPF (Windows Presentation Foundation) dengan pola MVVM.
* **Library UI:** MahApps.Metro (Gaya tampilan Modern "Metro").
* **Database:** Microsoft.Data.Sqlite (Database relasional tertanam).
* **Penanganan File:**
    * `ExcelDataReader`: Pembacaan stream file Excel biner yang cepat dan ringan.
    * `ClosedXML`: Pembuatan dan manipulasi laporan Excel (.xlsx) modern.

## ⚙️ Arsitektur & Standar Kode

Proyek ini mengikuti filosofi **"Pragmatic Human-Readable"**:
* **Linearitas di atas Abstraksi:** Logika ditulis agar mudah ditelusuri langkah demi langkah tanpa lapisan abstraksi yang berlebihan dan membingungkan.
* **Tipe Eksplisit:** Menggunakan `System.Int32`, `System.String` secara eksplisit untuk memastikan kejelasan dan menghindari ambiguitas tipe data.
* **Keamanan Resource:** Penggunaan blok `using` dan pola `Dispose()` secara agresif untuk mencegah kebocoran memori (memory leaks) pada operasi jangka panjang.

## 📸 Tangkapan Layar (Screenshots)
*(Tambahkan tangkapan layar Menu Utama, Grid Pencarian, dan Contoh Excel COA di sini)*
