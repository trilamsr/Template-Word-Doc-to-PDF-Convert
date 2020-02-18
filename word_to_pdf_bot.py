import os
import json
from docx import Document
from source_code import docx_fill_template, word_to_pdf



with open('settings.json') as settings_json:
    settings = json.load(settings_json)

with open('profile.json', 'r') as profile_json:
    profile = json.load(profile_json)

template_file = open(settings['template_location'], 'rb')
doc = Document(template_file)
docx_fill_template(doc, profile)
doc.save(settings['target_location'])

if settings['convert_to_pdf'] == True:
    word_to_pdf(settings['target_location'], settings['target_location'][:-4]+'pdf')

if settings["keep_word_copy"] == False:
    os.remove(settings['target_location'])