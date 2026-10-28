# -*- coding: utf-8 -*-
"""
Script parser và chuẩn hóa dữ liệu 10 Đề ôn tập Chương 1, 2, 3 Hóa học 12.
"""

import re

def normalize_chem_text(text):
    """
    Chuẩn hóa các công thức và thuật ngữ Hóa học:
    - amilase -> amylase
    - glucose, fructose, saccharose, maltose, cellulose, amylose, amylopectin
    - Chuyển công thức CTPT/CTCT thành dạng có ~...~ cho chỉ số dưới và ^...^ cho chỉ số trên
    """
    s = text.strip()
    
    # 1. Sửa lỗi chính tả thuật ngữ
    s = re.sub(r'\bamilase\b', 'amylase', s, flags=re.IGNORECASE)
    s = re.sub(r'\bα\s*-\s*', 'α-', s)
    s = re.sub(r'\bβ\s*-\s*', 'β-', s)
    s = re.sub(r'α-1,4-glycoside\b', 'α-1,4-glycosidic', s)
    s = re.sub(r'β-1,4-glycoside\b', 'β-1,4-glycosidic', s)
    s = re.sub(r'α-1,6-glycoside\b', 'α-1,6-glycosidic', s)
    s = re.sub(r'β-1,6-glycoside\b', 'β-1,6-glycosidic', s)
    
    # Độ C
    s = re.sub(r'(\d+)\s*°C', r'\1 °C', s)
    s = re.sub(r'(\d+)\s*oC\b', r'\1 °C', s)
    s = re.sub(r't°', 't°', s)

    # 2. Xử lý các ion
    s = re.sub(r'\b(Ca|Mg|Cu|Fe|Zn|Ba)\s*2\+', r'\1^2+^', s)
    s = re.sub(r'\b(Fe|Al|Cr)\s*3\+', r'\1^3+^', s)
    s = re.sub(r'\b(Na|K|Ag|H|NH4)\+', r'\1^+^', s)
    s = re.sub(r'\b(OH|Cl|Br|NO3|HCO3)-', r'\1^-^', s)
    s = re.sub(r'\b(SO4|CO3)2-', r'\1^2-^', s)

    # 3. Chuẩn hóa công thức hóa học hữu cơ và vô cơ thông dụng
    chem_tokens = [
        ("C17H33COOK", "C~17~H~33~COOK"),
        ("(C17H33COO)2Ca", "(C~17~H~33~COO)~2~Ca"),
        ("CH3COONa", "CH~3~COONa"),
        ("CH3[CH2]11C6H4SO3Na", "CH~3~[CH~2~]~11~C~6~H~4~SO~3~Na"),
        ("CH3COOCH=CH2", "CH~3~COOCH=CH~2~"),
        ("CH3NHCH3", "CH~3~NHCH~3~"),
        ("-COOH", "-COOH"),
        ("-NH2", "-NH~2~"),
        ("-CHO", "-CHO"),
        ("-OH", "-OH"),
        ("Cu(OH)2", "Cu(OH)~2~"),
        ("H2N-CH2-COOH", "H~2~N-CH~2~-COOH"),
        ("CH3COOH", "CH~3~COOH"),
        ("C2H5OH", "C~2~H~5~OH"),
        ("CH3NH2", "CH~3~NH~2~"),
        ("(C17H35COO)3C3H5", "(C~17~H~35~COO)~3~C~3~H~5~"),
        ("C12H22O11", "C~12~H~22~O~11~"),
        ("C6H12O6", "C~6~H~12~O~6~"),
        ("(C6H10O5)n", "(C~6~H~10~O~5~)~n~"),
        ("C3H6O2", "C~3~H~6~O~2~"),
        ("H2N-CH2-CO-NH-CH(CH3)COOH", "H~2~N-CH~2~-CO-NH-CH(CH~3~)COOH"),
        ("C6H5NH2", "C~6~H~5~NH~2~"),
        ("(CH3)2NH", "(CH~3~)~2~NH"),
        ("H2N-(CH2)6-NH2", "H~2~N-(CH~2~)~6~-NH~2~"),
        ("CH3-CH(NH2)-COOH", "CH~3~-CH(NH~2~)-COOH"),
        ("CH3-CH2-NH-CH3", "CH~3~-CH~2~-NH-CH~3~"),
        ("CH3[CH2]11OSO3Na", "CH~3~[CH~2~]~11~OSO~3~Na"),
        ("C12H25OSO3Na", "C~12~H~25~OSO~3~Na"),
        ("CO2", "CO~2~"),
        ("H2O", "H~2~O"),
        ("O2", "O~2~"),
        ("H2", "H~2~"),
        ("N2", "N~2~"),
        ("NH3", "NH~3~"),
        ("HNO3", "HNO~3~"),
        ("H2SO4", "H~2~SO~4~"),
        ("HNO2", "HNO~2~"),
        ("AgNO3", "AgNO~3~"),
        ("FeCl3", "FeCl~3~"),
        ("Cu2O", "Cu~2~O"),
        ("(C17H31COO)3C3H5", "(C~17~H~31~COO)~3~C~3~H~5~"),
        ("C57H98O6", "C~57~H~98~O~6~"),
        ("CH3COOC6H5", "CH~3~COOC~6~H~5~"),
        ("C6H5ONa", "C~6~H~5~ONa"),
        ("(C17H33COO)3C3H5", "(C~17~H~33~COO)~3~C~3~H~5~"),
        ("(C15H31COO)3C3H5", "(C~15~H~31~COO)~3~C~3~H~5~"),
        ("C15H31COONa", "C~15~H~31~COONa"),
        ("C17H35COONa", "C~17~H~35~COONa"),
        ("C17H33COONa", "C~17~H~33~COONa"),
        ("C17H31COONa", "C~17~H~31~COONa"),
        ("C6H5N2Cl", "C~6~H~5~N~2~Cl"),
        ("CH3COOCH2CH3", "CH~3~COOCH~2~CH~3~"),
        ("HCOOCH3", "HCOOCH~3~"),
        ("CH2=CHCOOCH3", "CH~2~=CHCOOCH~3~"),
        ("CH2=C(CH3)COOCH3", "CH~2~=C(CH~3~)COOCH~3~"),
        ("CH3COOCH2CH2CH(CH3)2", "CH~3~COOCH~2~CH~2~CH(CH~3~)~2~"),
        ("HCOOCH=CH2", "HCOOCH=CH~2~"),
        ("C7H14O2", "C~7~H~14~O~2~"),
        ("C4H8O2", "C~4~H~8~O~2~"),
        ("C3H7NO2", "C~3~H~7~NO~2~"),
        ("C2H5NO2", "C~2~H~5~NO~2~"),
        ("C5H11NO2", "C~5~H~11~NO~2~"),
        ("C4H9NO2", "C~4~H~9~NO~2~"),
        ("C2H7N", "C~2~H~7~N"),
        ("CH5N", "CH~5~N"),
        ("C3H9N", "C~3~H~9~N"),
        ("C6H7N", "C~6~H~7~N"),
        ("CaCO3", "CaCO~3~"),
        ("Pb(NO3)2", "Pb(NO~3~)~2~"),
        ("Br2", "Br~2~"),
        ("Cl2", "Cl~2~"),
        ("HCl", "HCl"),
        ("NaOH", "NaOH"),
        ("KOH", "KOH"),
        ("NaCl", "NaCl")
    ]
    for old, new in chem_tokens:
        # Thay thế có điều kiện biên từ hoặc ranh giới ký tự
        s = re.sub(r'(?<![A-Za-z0-9~^])' + re.escape(old) + r'(?![A-Za-z0-9~^])', new, s)

    return s
