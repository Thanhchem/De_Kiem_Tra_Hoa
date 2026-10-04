import os
from PIL import Image

def crop_figures():
    os.makedirs(r"c:\Antigravity_Thanh\images", exist_ok=True)
    
    # P1: media_1791080730724.jpg
    p1 = Image.open(r"C:/Users/Admin/.gemini/antigravity/brain/ef627512-8794-4964-a5d1-2dcb2dd52240/.user_uploaded/media_1791080730724.jpg")
    # Câu 6 figure: around y = 515 to 555, x = 200 to 350
    # Câu 7 figure: right side of Câu 7, around y = 580 to 700, x = 360 to 520
    
    # P2: media_1791080730689.jpg
    p2 = Image.open(r"C:/Users/Admin/.gemini/antigravity/brain/ef627512-8794-4964-a5d1-2dcb2dd52240/.user_uploaded/media_1791080730689.jpg")
    # Câu 11 figure: around y = 120 to 180, x = 220 to 420
    
    # Save crops with generous borders to inspect
    crop_c6 = p1.crop((200, 515, 340, 560))
    crop_c6.save(r"c:\Antigravity_Thanh\images\cau6_crop.png")
    
    crop_c7 = p1.crop((350, 580, 530, 700))
    crop_c7.save(r"c:\Antigravity_Thanh\images\catechin_crop.png")
    
    crop_c11 = p2.crop((220, 115, 430, 185))
    crop_c11.save(r"c:\Antigravity_Thanh\images\geraniol_crop.png")
    print("Cropped successfully!")

if __name__ == "__main__":
    crop_figures()
