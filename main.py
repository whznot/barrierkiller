import os
import openpyxl
from gtts import gTTS
from pydub import AudioSegment
from pydub.silence import silence


def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)


def generate_audio(text, lang, save_path):
    tts = gTTS(text=text, lang=lang)
    tts.save(save_path)


def add_audio_to_combined(audio_path, combined_audio, pause_duration):
    audio = AudioSegment.from_file(audio_path)
    pause = AudioSegment.silent(duration=pause_duration)
    return combined_audio + audio + pause


base_audio_dir = "audio"
ensure_dir(base_audio_dir)

words_dirs = {
    "de": os.path.join(base_audio_dir, "words", "de"),
    "ru": os.path.join(base_audio_dir, "words", "ru")
}

sentences_dirs = {
    "de": {
        "B1": os.path.join(base_audio_dir, "sentences", "de", "b1"),
        "B2": os.path.join(base_audio_dir, "sentences", "de", "b2"),
        "C1": os.path.join(base_audio_dir, "sentences", "de", "c1")
    },
    "ru": {
        "B1": os.path.join(base_audio_dir, "sentences", "ru", "b1"),
        "B2": os.path.join(base_audio_dir, "sentences", "ru", "b2"),
        "C1": os.path.join(base_audio_dir, "sentences", "ru", "c1")
    }
}

for lang_dirs in [words_dirs, sentences_dirs["de"], sentences_dirs["ru"]]:
    for path in lang_dirs.values():
        ensure_dir(path)

wb = openpyxl.load_workbook("phrases.xlsx")
sheet = wb.active

data = []
for row in sheet.iter_rows(min_row=2, values_only=True):
    german_word, russian_word, b1_german, b1_russian, b2_german, b2_russian, c1_german, c1_russian = row
    data.append({
        "german_word": german_word,
        "russian_word": russian_word,
        "examples": {
            "B1": {"de": b1_german, "ru": b1_russian},
            "B2": {"de": b2_german, "ru": b2_russian},
            "C1": {"de": c1_german, "ru": c1_russian}
        }
    })

combined_audio = AudioSegment.empty()
pause_duration = 2000
total_duration = 0

stop_generation = False

for idx, item in enumerate(data):
    if stop_generation:
        break
    
    for lang, text in [("de", item["german_word"]), ("ru", item["russian_word"])]:
        word_path = os.path.join(words_dirs[lang], f"{idx}.mp3")
        generate_audio(text, lang, word_path)
        combined_audio = add_audio_to_combined(
            word_path, combined_audio, pause_duration)
        total_duration += pause_duration

    for level, sentences in item["examples"].items():
        for lang, text in [("de", sentences["de"]), ("ru", sentences["ru"])]:
            sentence_path = os.path.join(
                sentences_dirs[lang][level], f"{idx}.mp3")
            generate_audio(text, lang, sentence_path)
            combined_audio = add_audio_to_combined(
                sentence_path, combined_audio, pause_duration)
            total_duration += pause_duration

            if total_duration / 1000 > 60:
                print("The limit of 1 minute has been reached. Stopping generation.")
                stop_generation = True
                break
        
        if stop_generation:
            break

combined_audio.export("output_audio.mp3", format="mp3")
print("The audio file was successfully created.")
