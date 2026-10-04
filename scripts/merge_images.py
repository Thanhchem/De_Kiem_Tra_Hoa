import os
import sys
from PIL import Image

def merge_images_to_pdf():
    images = [
        r"C:/Users/Admin/.gemini/antigravity/brain/ef627512-8794-4964-a5d1-2dcb2dd52240/.user_uploaded/media_1791080730724.jpg", # Page 1
        r"C:/Users/Admin/.gemini/antigravity/brain/ef627512-8794-4964-a5d1-2dcb2dd52240/.user_uploaded/media_1791080730689.jpg", # Page 2
        r"C:/Users/Admin/.gemini/antigravity/brain/ef627512-8794-4964-a5d1-2dcb2dd52240/.user_uploaded/media_1791080730704.jpg"  # Page 3
    ]
    
    output_pdf = r"c:\Antigravity_Thanh\De_Kiem_Tra_Hoa_11_Ma209_Goc.pdf"
    
    pil_images = []
    for img_path in images:
        if os.path.exists(img_path):
            img = Image.open(img_path)
            if img.mode != 'RGB':
                img = img.convert('RGB')
            pil_images.append(img)
        else:
            print(f"Error: {img_path} not found")
            return
            
    if pil_images:
        pil_images[0].save(output_pdf, save_all=True, append_images=pil_images[1:], resolution=100.0)
        print(f"Saved merged PDF to: {output_pdf}")

if __name__ == "__main__":
    merge_images_to_pdf()
