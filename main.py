import os
import pandas as pd
import shutil
from transliterate import translit
import editdistance

df = pd.read_excel("BazaIDCulture.xlsx")
df["Project_name"] = df["Project_name"].fillna('').astype(str)
df["Files_name"] = df["Files_name"].fillna('').astype(str)

result_folder = "result-folder"
renaming_dir = os.path.join(result_folder, "renaming-files")
clarification_needed_dir = os.path.join(result_folder, "clarification-needed")
no_match_found_dir = os.path.join(result_folder, "no-match-found")


def transliterate_ukrainian_to_english(ukrainian_text):
    return translit(ukrainian_text, 'uk', reversed=True)


def sanitize_filename(filename):
    return "".join(c for c in filename if c.isalnum() or c in (' ', '.', '_')).rstrip()


def find_files_in_directory(directory):
    file_paths = []
    for root, _, files in os.walk(directory):
        for file in files:
            file_paths.append(os.path.join(root, file))
    return file_paths


files = find_files_in_directory("Ark")

for file_path in files:
    file_name_with_ext = os.path.basename(file_path).lower()
    file_name, file_extension = os.path.splitext(file_name_with_ext)
    best_dist = 0
    best_name = ""
    new_file_name = ""

    for index, row in df.iloc[:7500].iterrows():
        program_name_ukr = row["Project_name"]
        desired_file_name = row["Files_name"]

        transliterated_name = transliterate_ukrainian_to_english(program_name_ukr)

        cur_distance = editdistance.eval(file_name, transliterated_name)
        norm_distance = round(1 - (cur_distance / max(len(file_name), len(transliterated_name))), 3)
        if norm_distance > best_dist:
            best_dist = norm_distance
            best_name = transliterated_name
            new_file_name = desired_file_name


    if best_dist >= 0.65:
        new_file_name_with_ext = new_file_name + file_extension
        print(f"Best match for '{file_name}' ---> '{best_name}' (probability {best_dist})")
        shutil.copy(file_path, os.path.join(renaming_dir, new_file_name_with_ext))
    elif 0.45 < best_dist < 0.65:
        new_file_name_with_ext = sanitize_filename(file_name) + "_" + new_file_name + "_" + sanitize_filename(best_name) + file_extension
        print(f"Сheck match for '{file_name}' ---> '{best_name}' (probability {best_dist})")
        shutil.copy(file_path, os.path.join(clarification_needed_dir, new_file_name_with_ext))
    else:
        print(f"No match found for '{file_name}'")
        shutil.copy(file_path, os.path.join(no_match_found_dir, file_name_with_ext))
