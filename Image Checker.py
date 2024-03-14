import requests
from PIL import Image
from io import BytesIO

# Replace this with your image URL
image_url = "https://freeforlifegroup.therapydatasolutions.com/getImageKey.php?id=16&key=849de66040dd76db73e7488e888a7dbeb2bcdc09786c29622a7cf922b615393f"

headers = {
    'Accept': '*/*',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
}

response = requests.get(image_url, headers=headers)

# Check if the request was successful
if response.status_code == 200:
    try:
        # Try to open the image
        img = Image.open(BytesIO(response.content))
        img.show()  # Display the image

        # Optionally, save to a file for inspection
        with open('downloaded_image.jpg', 'wb') as f:
            f.write(response.content)
        print("Image fetched and displayed successfully.")
    except Exception as e:
        print(f"Error opening image: {e}")
else:
    print(f"Failed to fetch the image. Status code: {response.status_code}")
