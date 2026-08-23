CONTRACT_TYPES = [
    "Oyuncu Sözleşmesi",
    "Senarist Sözleşmesi",
    "Yönetmen Sözleşmesi",
]

PROJECT_TYPES = ["Ana Akım TV", "Dijital", "Sinema"]

NEGOTIATION_LEVELS = ["Zayıf", "Orta", "Güçlü"]

PROFILE_STATUS = {
    ("Oyuncu Sözleşmesi", "Ana Akım TV"): "ACTIVE",
    ("Oyuncu Sözleşmesi", "Dijital"): "BASIC",
    ("Oyuncu Sözleşmesi", "Sinema"): "BASIC",
    ("Senarist Sözleşmesi", "Ana Akım TV"): "BASIC",
    ("Senarist Sözleşmesi", "Dijital"): "BASIC",
    ("Senarist Sözleşmesi", "Sinema"): "BASIC",
    ("Yönetmen Sözleşmesi", "Ana Akım TV"): "BASIC",
    ("Yönetmen Sözleşmesi", "Dijital"): "BASIC",
    ("Yönetmen Sözleşmesi", "Sinema"): "BASIC",
}
