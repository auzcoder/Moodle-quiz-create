

# v4 Tartib raqam qo'shildi va uni olib tashlab tayyorlaydi
import pandas as pd
from docx import Document


def read_docx_table(docx_path):
    document = Document(docx_path)
    data = []

    table = document.tables[0]  # Birinchi jadvalni olish
    for i, row in enumerate(table.rows):
        text = [cell.text.strip() for cell in row.cells]
        data.append(text)

    # Pandas DataFrame yaratish
    df = pd.DataFrame(data[1:], columns=data[0])
    return df


def format_question(row):
    # Null yoki bo'sh qiymatlar mavjudligini tekshirish
    if row.isnull().any() or row.str.strip().eq('').any():
        return None

    correct_option = f"= {row.iloc[1]}"
    wrong_options = [f"~ {row.iloc[i]}" for i in range(2, len(row))]

    return f":: {row.iloc[0]} {{\n{correct_option}\n{'\n'.join(wrong_options)}\n}}"
print("tayyor")


def main(docx_path, output_txt_path):
    df = read_docx_table(docx_path)

    # Tartib raqami ustunini olib tashlash
    df = df.iloc[:, 1:]

    # Har bir savolni formatlash
    formatted_questions = df.apply(format_question, axis=1)

    # Null yoki bo'sh qiymatlar bo'lmagan savollarni filterlash
    formatted_questions = formatted_questions.dropna()

    # Natijani faylga yozish
    with open(output_txt_path, 'w') as f:
        for question in formatted_questions:
            if question is not None:
                f.write(question + '\n\n')


# docx fayl manzilini va chiqadigan txt fayl manzilini kiriting
docx_path = '/Users/auzcoder/Desktop/testuchun/pythonProject/Moodle-quiz-create/test2.docx'  # DOCX fayl manzili
output_txt_path = '/Users/auzcoder/Desktop/testuchun/pythonProject/Moodle-quiz-create/output.txt'  # TXT fayl manzili

main(docx_path, output_txt_path)
