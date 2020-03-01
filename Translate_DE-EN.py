from Read_write_workbook import read_from_xlsx, write_to_xlsx
from googletrans import Translator


def translate_de_to_en():
    filename = input('Excel table name to be read: ')
    column = input('Column letter to be read: ')
    start_row = int(input('First row number to be read: '))
    end_row = int(input('Last row number to be read: '))

    translated_text = []
    translator = Translator()

    sentences_to_be_translated = read_from_xlsx(filename, start_row, end_row, column)

    for sentence_to_be_translated in sentences_to_be_translated:
        translated_text.append(translator.translate(sentence_to_be_translated, dest='en').text + " (" + sentence_to_be_translated + ")")

    write_to_xlsx(filename, start_row, column, translated_text)


if __name__ == "__main__":
    translate_de_to_en()
