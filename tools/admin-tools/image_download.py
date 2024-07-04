import requests

def download_images(image_list):
    headers = requests.utils.default_headers()

    headers.update(
        {
            'User-Agent': 'My User Agent 1.0',
        }
    )

    for url, filename in image_list.items():
        try:
            response = requests.get(f"{base_url}{url}", stream=True, headers=headers)
            response.raise_for_status()

            with open(f"{download_folder}{filename}", 'wb') as f:
                for chunk in response.iter_content(chunk_size=1024):
                    if chunk:
                        f.write(chunk)

            print(f"Downloaded {filename} successfully!")
        except requests.exceptions.RequestException as e:
            print(f"Error downloading {url}: {e}")


base_url = ""
download_folder = "./downloaded-images/"
image_list = {
    "https://www.rolscarpets.com/en/wp-content/uploads/sites/2/2023/06/Terrra_Atacama_Iron_LInen_01.webp": "terra-atacama-iron-linen.webp",
    "https://www.rolscarpets.com/en/wp-content/uploads/sites/2/2023/06/Terrra_Atacama_Coconut_White_01.webp": "terra-atacama-coconut-white.webp",
    "https://www.rolscarpets.com/en/wp-content/uploads/sites/2/2023/06/Terrra_Atacama_Oat_Nut_01.webp": "terra-atacama-oat-nut.webp",
    "https://www.rolscarpets.com/en/wp-content/uploads/sites/2/2023/06/Terrra_Atacama_Coconut_Black_01.webp": "terra-atacama-coconut-black.webp",
    "https://www.rolscarpets.com/en/wp-content/uploads/sites/2/2023/04/Terra_Sahara_Nut_01.webp": "terra-sahara-nut.webp",
    "https://www.rolscarpets.com/en/wp-content/uploads/sites/2/2023/04/Terra_Sahara_Oat_01.webp": "terra-sahara-oat.webp",
    "https://www.rolscarpets.com/en/wp-content/uploads/sites/2/2023/04/Terra_Sahara_Coconut_01.webp": "terra-sahara-coconut.webp",
    "https://www.rolscarpets.com/en/wp-content/uploads/sites/2/2023/04/Terra_Sahara_Iron_01.webp": "terra-sahara-iron.webp",
    "https://www.rolscarpets.com/en/wp-content/uploads/sites/2/2023/06/Terra_Serengueti_Coconut_01.webp": "terra-serengueti-coconut.webp",
    "https://www.rolscarpets.com/en/wp-content/uploads/sites/2/2023/06/Terra_Serengueti_Black_White_01.webp": "terra-serengueti-black-white.webp",
    "https://www.rolscarpets.com/en/wp-content/uploads/sites/2/2023/06/Terra_Serengueti_Rose_White_01.webp": "terra-serengueti-rose-white.webp",
    "https://www.rolscarpets.com/en/wp-content/uploads/sites/2/2023/06/Terra_Serengueti_Ivy_White_01.webp": "terra-serengueti-ivy-white.webp",
    "https://www.rolscarpets.com/en/wp-content/uploads/sites/2/2022/06/Terra_Kalahari_Oat_Sea_01.webp": "terra-kalahari-oat-sea.webp",
    "https://www.rolscarpets.com/en/wp-content/uploads/sites/2/2023/11/Terra_Gorafe_Oat_Cobalt_01.webp": "terra-gorafe-oat-cobalt.webp",
    "https://www.rolscarpets.com/en/wp-content/uploads/sites/2/2023/06/Terra_Gorafe_Iron_Linen_01.webp": "terra-gorafe-iron-linen.webp",
    "https://www.rolscarpets.com/en/wp-content/uploads/sites/2/2023/06/Terra_Gorafe_Oat_Nut_01.webp": "terra-gorafe-oat-nut.webp",
    "https://www.rolscarpets.com/en/wp-content/uploads/sites/2/2023/06/Terra_Uyuni_Coconut_01.webp": "terra-uyuni-coconut-white.webp",
    "https://www.rolscarpets.com/en/wp-content/uploads/sites/2/2023/06/Terra_Uyuni_Iron_Linen_01.webp": "terra-uyuni-iron-linen.webp",
    "https://www.rolscarpets.com/en/wp-content/uploads/sites/2/2023/06/Terrra_Gobi_Iron_Linen_01.webp": "terra-gobi-iron-linen.webp",
    "https://www.rolscarpets.com/en/wp-content/uploads/sites/2/2023/06/Terrra_Gobi_Coconut_White_01.webp": "terra-gobi-coconut-white.webp",
}
download_images(image_list)
