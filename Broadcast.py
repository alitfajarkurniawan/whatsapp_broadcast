import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep

# Masukkan path ke ChromeDriver
service = Service(r'chromedriver.exe')  # Ganti dengan path yang sesuai
driver = webdriver.Chrome(service=service)

# Buka WhatsApp Web
driver.get('https://web.whatsapp.com')

# Tunggu hingga login selesai
input("Pindai QR Code di WhatsApp Web dan tekan Enter jika sudah login.")

# Baca data dari file Excel
data = pd.read_excel("mahasiswa.xlsx")

# Pesan template
message_template = (
    "Assalamu’alaikum warahmatullahi wabarakatuh\n"
    "Hi {nama}, Apa kabar?\n"
    "Dengan penuh rasa syukur kepada Allah Subhanahu wa Ta'ala, kami dari Panitia Penerimaan Mahasiswa Baru STAIN Teungku Dirundeng Meulaboh mengundang seluruh calon mahasiswa baru asal Aceh Barat untuk bergabung dalam grup WhatsApp berikut:\n"
    "👉 https://chat.whatsapp.com/Kk6qt21pUar3USUvFSCFew\n\n"
    "Kenapa Harus Bergabung?\n"
    "Grup ini adalah tempat untuk:\n"
    "✅ Mendapatkan informasi resmi terkait pendaftaran mahasiswa baru.\n"
    "✅ Diskusi langsung dengan panitia tentang program studi, beasiswa, dan prosedur penerimaan.\n"
    "✅ Akses cepat ke jadwal penting dan pengumuman lainnya.\n\n"
    "Dengan bergabung di grup ini, Anda akan selalu mendapatkan informasi terbaru dan tidak ketinggalan peluang emas untuk melanjutkan pendidikan di kampus islami yang berkualitas.\n\n"
    "Catatan Penting:\n"
    "📌 Grup ini khusus untuk calon mahasiswa baru asal Aceh Barat.\n"
    "📌 Gunakan nama asli saat bergabung agar mempermudah identifikasi.\n\n"
    "Semoga Allah memudahkan langkah kita semua dalam menuntut ilmu dan menggapai ridha-Nya. Kami sangat berharap kehadiran Anda menjadi bagian dari keluarga besar STAIN Teungku Dirundeng Meulaboh.\n\n"
    "Wassalamu’alaikum warahmatullahi wabarakatuh\n\n"
    "Hormat kami,\n"
    "Panitia Penerimaan Mahasiswa Baru\n"
    "STAIN Teungku Dirundeng Meulaboh"
)

# Loop untuk mengirim pesan ke setiap nomor
for index, row in data.iterrows():
    name = row['NAMA']
    number = str(row['NOMOR HP'])
    # Isi nama ke dalam template pesan
    formatted_message = message_template.format(nama=name).replace("\n", "%0A")  # Ganti '\n' dengan '%0A'

    # Buka chat dengan nomor tujuan
    driver.get(f"https://web.whatsapp.com/send?phone={number}&text={formatted_message}")
    sleep(5)
    
    # Tunggu hingga tombol kirim muncul dan klik
    try:
        send_button = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, '//span[@data-icon="send"]'))
        )
        send_button.click()
        print(f"Pesan berhasil dikirim ke {name} ({number})")
        sleep(3)
    except Exception as e:
        print(f"Gagal mengirim pesan ke {name} ({number}): {e}")

# Tutup browser
driver.quit()
