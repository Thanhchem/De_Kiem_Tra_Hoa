# -*- coding: utf-8 -*-
"""
Script vẽ lại các hình ảnh sơ đồ, phổ MS, phổ IR và công thức phân tử Geraniol cho đề THPT Bùi Thị Xuân - Mã 1102.
Hình vẽ vector / 300 DPI sắc nét, chuẩn sạch sẽ, loại bỏ hoàn toàn nét bút bi của học sinh.
"""

import os
import matplotlib.pyplot as plt
import numpy as np

OUT_DIR = os.path.abspath("San_Pham/images_ck1_hoa11_buithixuan")
os.makedirs(OUT_DIR, exist_ok=True)

# Cấu hình font và kiểu hiển thị
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['font.size'] = 10

# 1. Vẽ Phổ Khối lượng Câu 7 (Hợp chất X có M = 44)
def draw_ms_cau7():
    fig, ax = plt.subplots(figsize=(6, 3.5), dpi=300)
    
    # Dữ liệu peak m/z và cường độ (%)
    peaks = {
        15: 12,
        28: 15,
        29: 20,
        30: 100,
        42: 8,
        43: 15,
        44: 35
    }
    
    for mz, intensity in peaks.items():
        ax.plot([mz, mz], [0, intensity], color='black', linewidth=1.5)
        if mz in [28, 30, 44]:
            ax.text(mz, intensity + 3, str(mz), ha='center', va='bottom', fontsize=9, fontweight='bold')

    ax.set_xlim(10, 80)
    ax.set_ylim(0, 115)
    ax.set_xlabel(r'$m/z$', fontsize=11, fontweight='bold')
    ax.set_ylabel('Cường độ tương đối (%)', fontsize=10, fontweight='bold')
    ax.set_title('Phổ khối lượng của hợp chất hữu cơ X', fontsize=11, pad=10, fontweight='bold')
    ax.grid(True, linestyle='--', alpha=0.4)
    ax.set_xticks([15, 30, 45, 60, 75])
    ax.set_yticks([0, 20, 40, 60, 80, 100])
    
    plt.tight_layout()
    path = os.path.join(OUT_DIR, "ms_cau7_clean.png")
    plt.savefig(path, dpi=300)
    plt.close()
    print(f"Đã vẽ: {path}")

# 2. Vẽ Công thức cấu tạo Geraniol Câu 4 Phần III
def draw_geraniol():
    fig, ax = plt.subplots(figsize=(5, 3), dpi=300)
    
    # Tọa độ các đỉnh của mạch khung liên kết Geraniol
    # (CH3)2C=CH-CH2-CH2-C(CH3)=CH-CH2-OH
    # Đỉnh:
    # C1 (nhánh methyl trên): (0.5, 2.5)
    # C2 (nhánh methyl dưới): (0.5, 0.5)
    # C3 (C=): (1.5, 1.5)
    # C4 (=CH): (2.5, 1.5)
    # C5 (CH2): (3.5, 2.3)
    # C6 (CH2): (4.5, 1.5)
    # C7 (C(CH3)=): (5.5, 2.3)
    # C_branch (nhánh CH3 trên C7): (5.5, 3.4)
    # C8 (=CH): (6.5, 1.5)
    # C9 (CH2): (7.5, 2.3)
    # O (OH): (8.5, 1.5)
    
    # Liên kết đôi C3=C4
    ax.plot([0.5, 1.5], [2.3, 1.5], 'k-', lw=2) # CH3 trên
    ax.plot([0.5, 1.5], [0.7, 1.5], 'k-', lw=2) # CH3 dưới
    ax.text(0.4, 2.4, r'$\mathbf{CH_3}$', ha='right', va='center', fontsize=11)
    ax.text(0.4, 0.6, r'$\mathbf{H_3C}$', ha='right', va='center', fontsize=11)
    
    # Nối đôi C3=C4
    ax.plot([1.5, 2.7], [1.55, 1.55], 'k-', lw=2.2)
    ax.plot([1.5, 2.7], [1.45, 1.45], 'k-', lw=2.2)
    
    # C4-C5-C6-C7
    ax.plot([2.7, 3.8], [1.5, 2.3], 'k-', lw=2)
    ax.plot([3.8, 4.9], [2.3, 1.5], 'k-', lw=2)
    ax.plot([4.9, 6.0], [1.5, 2.3], 'k-', lw=2)
    
    # Nhánh CH3 trên C7
    ax.plot([6.0, 6.0], [2.3, 3.3], 'k-', lw=2)
    ax.text(6.0, 3.5, r'$\mathbf{CH_3}$', ha='center', va='bottom', fontsize=11)
    
    # Nối đôi C7=C8
    ax.plot([6.0, 7.1], [2.33, 1.53], 'k-', lw=2.2)
    ax.plot([5.95, 7.05], [2.22, 1.42], 'k-', lw=2.2)
    
    # C8-C9-OH
    ax.plot([7.1, 8.2], [1.5, 2.3], 'k-', lw=2)
    ax.plot([8.2, 9.1], [2.3, 1.7], 'k-', lw=2)
    ax.text(9.2, 1.6, r'$\mathbf{OH}$', ha='left', va='center', fontsize=11)
    
    ax.set_xlim(-0.2, 10.2)
    ax.set_ylim(0.0, 4.2)
    ax.axis('off')
    
    plt.tight_layout()
    path = os.path.join(OUT_DIR, "geraniol_clean.png")
    plt.savefig(path, dpi=300)
    plt.close()
    print(f"Đã vẽ: {path}")

# 3. Vẽ Phổ Khối lượng Câu 3 Tự luận (Acetic acid, M = 60)
def draw_ms_acetic_acid():
    fig, ax = plt.subplots(figsize=(6, 3.5), dpi=300)
    
    peaks = {
        15: 18,
        29: 8,
        43: 82,
        45: 100,
        60: 68
    }
    
    for mz, intensity in peaks.items():
        ax.plot([mz, mz], [0, intensity], color='black', linewidth=1.5)
        ax.text(mz, intensity + 3, str(mz), ha='center', va='bottom', fontsize=9, fontweight='bold')

    ax.set_xlim(5, 75)
    ax.set_ylim(0, 115)
    ax.set_xlabel(r'$m/z$', fontsize=11, fontweight='bold')
    ax.set_ylabel('Cường độ tương đối (%)', fontsize=10, fontweight='bold')
    ax.set_title('Phổ khối lượng của Acetic acid', fontsize=11, pad=10, fontweight='bold')
    ax.grid(True, linestyle='--', alpha=0.4)
    ax.set_xticks([10, 20, 30, 40, 50, 60, 70])
    ax.set_yticks([0, 20, 40, 60, 80, 100])
    
    plt.tight_layout()
    path = os.path.join(OUT_DIR, "ms_acetic_acid_clean.png")
    plt.savefig(path, dpi=300)
    plt.close()
    print(f"Đã vẽ: {path}")

if __name__ == "__main__":
    draw_ms_cau7()
    draw_geraniol()
    draw_ms_acetic_acid()
