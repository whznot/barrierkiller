import os
import pandas as pd
from gtts import gTTS
import subprocess

df = pd.read_excel("phrases.xlsx", sheet_name="phrases")

columns_order = []

for col in df.columns:
    if col == "de" or col.endswith("_de"):
        columns_order.append((col, "de"))
    elif col == "ru" or col.endswith("_ru"):
        columns_order.append((col, "ru"))
    else:
        continue


def text_to_speech(text, lang, filename='speech_output.mp3'):
    tts = gTTS(text=text, lang=lang, slow=False)
    tts.save(filename)


if __name__ == "__main__":
    os.makedirs("audio", exist_ok=True)

    MAX_ROWS = 20
    generated_files = []
    silence_file = os.path.join("audio", "silence2s.mp3")
    
    for i, row_data in df.iterrows():
        if i > MAX_ROWS:
            break

        for col_name, lang_code in columns_order:
            text_value = row_data[col_name]
            if pd.isna(text_value):
                continue

            text_str = str(text_value).strip()
            
            filename = f"row_{i}_{col_name}.mp3"
            filepath = os.path.join("audio", filename)

            text_to_speech(text_str, lang_code, filename=filepath)

            generated_files.append(filepath)

            if os.path.exists(silence_file):
                generated_files.append(silence_file)

    if not generated_files:
        exit()

    list_txt = os.path.join("audio", "list.txt")
    with open(list_txt, "w", encoding="utf-8") as f:
        for file_path in generated_files:
            f.write(f"file '{os.path.abspath(file_path)}'\n")

    final_output = os.path.join("audio", "final_combined.mp3")

    cmd = [
        "ffmpeg",
        "-f", "concat",
        "-safe", "0",
        "-i", list_txt,
        "-c", "copy",
        final_output
    ]
    subprocess.run(cmd, check=True)

    print(f"Готово! Итоговый файл: {final_output}")