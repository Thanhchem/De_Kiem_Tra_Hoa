# -*- coding: utf-8 -*-
"""
Dữ liệu Tổng hợp Đề cương Ôn tập Giữa kì I - KHTN 6
Trường THCS Kim Long - Năm học 2026 - 2027
Biên soạn chuẩn hóa theo "FORM ĐỀ KIỂM TRA" - Thầy TRẦN VĂN THẠNH
"""

import sys
import os

sys.path.append(os.path.dirname(__file__))

from khtn6_kimlong_data_common import (
    INFO, PHAN_CHUNG_THEORY, PHAN_CHUNG_MCQ, PHAN_CHUNG_TF,
    PHAN_CHUNG_SHORT, PHAN_CHUNG_ESSAY
)
from khtn6_kimlong_data_vatli import (
    VATLI_THEORY, VATLI_MCQ, VATLI_TF, VATLI_FILL, VATLI_ESSAY
)
from khtn6_kimlong_data_hoa import (
    HOA_THEORY, HOA_MCQ, HOA_TF, HOA_SHORT, HOA_ESSAY
)
from khtn6_kimlong_data_sinh import (
    SINH_THEORY, SINH_MCQ, SINH_TF, SINH_SHORT, SINH_ESSAY
)

__all__ = [
    "INFO",
    "PHAN_CHUNG_THEORY", "PHAN_CHUNG_MCQ", "PHAN_CHUNG_TF", "PHAN_CHUNG_SHORT", "PHAN_CHUNG_ESSAY",
    "VATLI_THEORY", "VATLI_MCQ", "VATLI_TF", "VATLI_FILL", "VATLI_ESSAY",
    "HOA_THEORY", "HOA_MCQ", "HOA_TF", "HOA_SHORT", "HOA_ESSAY",
    "SINH_THEORY", "SINH_MCQ", "SINH_TF", "SINH_SHORT", "SINH_ESSAY"
]
