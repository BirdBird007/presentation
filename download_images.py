import os
import urllib.request
from PIL import Image

def download_and_convert(url, dest_name):
    print(f"Downloading {url} ...")
    temp_file = "temp_" + dest_name
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36'
    }
    
    try:
        # Fetch the image with User-Agent headers to avoid HTTP 403 Forbidden
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req) as response:
            with open(temp_file, 'wb') as out_file:
                out_file.write(response.read())
        
        # Convert to PNG using Pillow to make sure it's standard and valid
        with Image.open(temp_file) as img:
            img.save(dest_name, "PNG")
        
        # Clean up temporary file
        os.remove(temp_file)
        print(f"[SUCCESS] Saved and converted to {dest_name}")
    except Exception as e:
        print(f"[ERROR] Failed to download or convert {dest_name} from {url}: {e}")
        if os.path.exists(temp_file):
            os.remove(temp_file)

def main():
    images = [
        {
            "url": "https://image.bangkokbiznews.com/uploads/images/contents/w1024/2023/08/0VKYtVYJMd7B6BykqtTD.webp",
            "dest": "manit_profile.png"
        },
        {
            "url": "https://image.bangkokbiznews.com/uploads/images/contents/w1024/2023/08/J2LP0qT4Q7Kkg2PbQJhM.webp",
            "dest": "rena_profile.png"
        },
        {
            "url": "https://image.bangkokbiznews.com/uploads/images/contents/w1024/2023/08/22NHoKbGdtN3hgCtCLCf.webp",
            "dest": "luxury_villa_lake.png"
        },
        {
            "url": "https://image.bangkokbiznews.com/uploads/images/md/2023/08/kwmcRGujNUKZM65EEgRT.webp",
            "dest": "mineral_water_bottle.png"
        },
        {
            "url": "https://assets.brandinside.asia/uploads/2020/03/banner-home-01.jpg",
            "dest": "modern_retail_store.png"
        }
    ]
    
    for img in images:
        download_and_convert(img["url"], img["dest"])

if __name__ == "__main__":
    main()
