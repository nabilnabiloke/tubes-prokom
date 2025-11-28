import pandas as pd
import numpy as np

# -------------------------------------------------------------
# 1. Fungsi Baca File (Excel / Google Sheets) + Error Handling
# -------------------------------------------------------------
def load_input_file(path_or_url):
    try:
        if "docs.google.com" in path_or_url:
            try:
                csv_url = path_or_url.replace("/edit#", "/export?format=csv&")
                df = pd.read_csv(csv_url)
                print("[OK] Data berhasil dibaca dari Google Sheets")
                return df
            except Exception as e:
                raise ValueError(f"Gagal membaca Google Sheets: {e}")
        
        # Jika Excel
        df = pd.read_excel(path_or_url)
        print("[OK] Data berhasil dibaca dari Excel")
        return df
    
    except FileNotFoundError:
        raise FileNotFoundError("ERROR: File tidak ditemukan. Pastikan path benar.")
    except Exception as e:
        raise Exception(f"ERROR membaca file: {e}")


# -------------------------------------------------------------
# 2. Fungsi Validasi Kolom
# -------------------------------------------------------------
def validate_columns(df):
    required = ["nim", "nilai", "jenis_penilaian"]
    missing = [col for col in required if col not in df.columns]

    if missing:
        raise ValueError(f"ERROR: Kolom berikut tidak ditemukan: {missing}\n"
                         f"Pastikan nama kolom sama persis: {required}")

    print("[OK] Kolom valid")
    return True


# -------------------------------------------------------------
# 3. Normalisasi Kolom dan Tipe Data
# -------------------------------------------------------------
def clean_dataframe(df):
    try:
        df["nim"] = df["nim"].astype(str).str.replace(r'\D', '', regex=True)
        df["nilai"] = pd.to_numeric(df["nilai"], errors="coerce")

        if df["nilai"].isna().sum() > 0:
            print("[WARNING] Ada nilai yang tidak valid dan diubah menjadi NaN")

        df["jenis_penilaian"] = df["jenis_penilaian"].str.lower().str.strip()
        print("[OK] Data berhasil dinormalisasi")
        return df
    except Exception as e:
        raise Exception(f"ERROR saat normalisasi data: {e}")


# -------------------------------------------------------------
# 4. Mapping Jenis Penilaian → Template Final
# -------------------------------------------------------------
mapping = {
    "uk1": "uk1",
    "uk 1": "uk1",
    "unit kompetensi 1": "uk1",

    "uk2": "uk2",
    "uk 2": "uk2",

    "uk3": "uk3",
    "uk 3": "uk3",

    "proyek": "hasil_proyek",
    "project": "hasil_proyek",
    "final project": "hasil_proyek",

    "partisipasi": "partisipasi",
    "presensi": "partisipasi",
}

def map_jenis_penilaian(df):
    try:
        df["template_kolom"] = df["jenis_penilaian"].map(mapping)

        if df["template_kolom"].isna().sum() > 0:
            unknown = df[df["template_kolom"].isna()]["jenis_penilaian"].unique()
            raise ValueError(f"ERROR: Jenis penilaian tidak dikenal: {unknown}")

        print("[OK] Jenis penilaian berhasil dipetakan ke template")
        return df
    except Exception as e:
        raise Exception(f"ERROR mapping jenis penilaian: {e}")


# -------------------------------------------------------------
# 5. Hitung Rata-Rata per Mahasiswa per Jenis Penilaian
# -------------------------------------------------------------
def compute_scores(df):
    try:
        grouped = df.groupby(["nim", "template_kolom"])["nilai"].mean().reset_index()
        pivot = grouped.pivot(index="nim", columns="template_kolom", values="nilai")
        print("[OK] Rata-rata nilai berhasil dihitung")
        return pivot
    except Exception as e:
        raise Exception(f"ERROR perhitungan nilai: {e}")


# -------------------------------------------------------------
# 6. Buat Output Final dengan Template
# -------------------------------------------------------------
def assemble_output(final_scores, df_mahasiswa=None):
    try:
        final_scores = final_scores.reset_index()

        output = pd.DataFrame()
        output["nim"] = final_scores["nim"]

        # Kolom default
        for kolom in ["uk1", "uk2", "uk3", "partisipasi", "hasil_proyek"]:
            output[kolom] = final_scores.get(kolom, np.nan)

        # nilai akhir = rata-rata semua
        nilai_columns = ["uk1", "uk2", "uk3", "partisipasi", "hasil_proyek"]
        output["nilai_akhir"] = output[nilai_columns].mean(axis=1)

        # nilai skala
        output["nilai_skala"] = (output["nilai_akhir"] / 100 * 4).round(2)

        # huruf
        def to_grade(x):
            if x >= 85: return "A"
            if x >= 75: return "B"
            if x >= 65: return "C"
            if x >= 50: return "D"
            return "E"

        output["huruf"] = output["nilai_akhir"].apply(to_grade)

        # urutkan berdasarkan nim → otomatis absensi
        output = output.sort_values("nim").reset_index(drop=True)
        output.insert(0, "no", range(1, len(output) + 1))

        print("[OK] Output final berhasil disusun")
        return output

    except Exception as e:
        raise Exception(f"ERROR saat menyusun output: {e}")


# -------------------------------------------------------------
# MAIN PROGRAM
# -------------------------------------------------------------
def process_file(input_path):
    df = load_input_file(input_path)
    validate_columns(df)
    df = clean_dataframe(df)
    df = map_jenis_penilaian(df)

    scores = compute_scores(df)
    final_output = assemble_output(scores)

    print("\n=== OUTPUT FINAL (TERURUT) ===")
    print(final_output)
    return final_output

if __name__ == "__main__":
    print("=== Proses dimulai ===\n")

    # Minta input file langsung dari user
    path = input("Masukkan path file Excel / link Google Sheets: ").strip()

    try:
        output = process_file(path)
        print("\n=== PROSES SELESAI ===")
        print(output)
    except Exception as e:
        print("\n[ERROR]", e)



# -------------------------------------------------------------
# Contoh Pemanggilan
# -------------------------------------------------------------
# output = process_file("nilai_kelas.xlsx")
# output = process_file("https://docs.google.com/spreadsheets/d/xxxxx/edit#gid=0")

