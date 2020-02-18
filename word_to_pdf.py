import os
import json
from copy import deepcopy
from docx import Document
from source_code import docx_fill_template, word_to_pdf


with open('settings.json') as settings_json:
    settings = json.load(settings_json)

with open('profile.json', 'r') as profile_json:
    profiles_array = json.load(profile_json)

template_file = open(settings['target_template'], 'rb')
doc = Document(template_file)

for item in profiles_array:
    cur_doc = deepcopy(doc)
    full_path_docx = settings['output_location'] + item["output_name"] + ".docx"
    full_path_pdf  = settings['output_location'] + item["output_name"] + ".pdf"

    docx_fill_template(cur_doc, item["profile"])
    cur_doc.save(full_path_docx)

    if settings['create_pdf'] == True:
        word_to_pdf(full_path_docx, full_path_pdf)

    if settings["create_word_doc"] == False:
        os.remove(full_path_docx)

template_file.close()