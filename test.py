import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By

# Setup WebDriver (pastikan Anda memiliki driver yang sesuai, misalnya chromedriver)
driver = webdriver.Firefox()

# Buka halaman web
driver.get('https://www.ojk.go.id/id/Regulasi/Default.aspx')  # Ganti dengan URL halaman yang sesuai

# Mengambil semua elemen <a> di halaman
a_elements = driver.find_elements(By.TAG_NAME, 'a')

# List to store link titles and URLs
links_data = []

# Mengambil teks dan URL dari setiap elemen <a>
for a in a_elements:
    link_text = a.text
    link_url = a.get_attribute('href')
    if link_url:  # Pastikan href ada
        links_data.append({"Title": link_text, "URL": link_url})

# Menyimpan data dalam DataFrame
df_links = pd.DataFrame(links_data)

# Menampilkan DataFrame
print(df_links)

# Menutup browser setelah selesai
driver.quit()
