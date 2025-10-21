import base64

# Pastikan path-nya menunjuk ke file Excel yang benar
template_path = r"D:\PROYEK PROGRAMMING\FILE OUTPUT.xlsx"

# Baca file template dan ubah ke base64
with open(template_path, "rb") as f:
    encoded = base64.b64encode(f.read()).decode("utf-8")

# Simpan hasil base64 ke file teks biar mudah disalin nanti
with open("template_base64.txt", "w", encoding="utf-8") as out:
    out.write(encoded)

print("✅ Konversi selesai! File 'template_base64.txt' sudah dibuat di folder ini.")
