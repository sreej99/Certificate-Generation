import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os

# === CONFIGURATIONS ===
EXCEL_FILE = r'C:\Users\admin\Desktop\Certficate volunteers\Volunteers.xlsx'
SHEET_NAME = 'Sheet1'
NAME_COLUMN = 'Name'
CERTIFICATE_TEMPLATE = r'C:\Users\admin\Desktop\Certficate volunteers\Final certificate.png'
OUTPUT_FOLDER = r'C:\Users\admin\DesktopCertficate volunteers\certificates'

FONT_PATH = r"C:\Users\admin\Desktop\Certficate volunteers\PlayfairDisplay-Italic-VariableFont_wght.ttf"
FONT_COLOR = (0, 0, 0)
MAX_FONT_SIZE = 100
MIN_FONT_SIZE = 40

TEXT_CENTER_X = 1058
TEXT_CENTER_Y = 560
TEXT_BOX_WIDTH = 1094

# === VERIFY FONT EXISTS ===
if not os.path.isfile(FONT_PATH):
    print("❌ Font file NOT found. Check the path and filename:")
    print(FONT_PATH)
    exit()
else:
    print("✅ Font file exists!")

# === LOAD EXCEL ===
df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)

# === PREPARE OUTPUT DIRECTORY ===
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# === FUNCTION TO FIT TEXT TO WIDTH ===
def fit_text(draw, text, font_path, max_width, max_font_size, min_font_size):
    font_size = max_font_size
    while font_size >= min_font_size:
        font = ImageFont.truetype(font_path, font_size)
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        if text_width <= max_width:
            return font
        font_size -= 1
    return ImageFont.truetype(font_path, min_font_size)

# === MAIN LOOP ===
for index, row in df.iterrows():
    name = str(row[NAME_COLUMN]).strip()

    # Open certificate image
    base = Image.open(CERTIFICATE_TEMPLATE).convert('RGB')
    draw = ImageDraw.Draw(base)

    # Fit font to name
    font = fit_text(draw, name, FONT_PATH, TEXT_BOX_WIDTH, MAX_FONT_SIZE, MIN_FONT_SIZE)

    # Calculate text position (centered)
    bbox = draw.textbbox((0, 0), name, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    position = (TEXT_CENTER_X - text_width // 2, TEXT_CENTER_Y - text_height // 2)

    # Draw name on certificate
    draw.text(position, name, fill=FONT_COLOR, font=font)

    # Save certificate
    output_path = os.path.join(OUTPUT_FOLDER, f'{name}.png')
    base.save(output_path)
    print(f"✅ Saved: {output_path}")

print("🎉 All volunteer certificates generated successfully!")


