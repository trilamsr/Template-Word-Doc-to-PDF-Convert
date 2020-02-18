import os
import sys
import comtypes.client

def word_to_pdf(input, output):
    path_in = os.path.abspath(input)
    path_out = os.path.abspath(output)
    word_doc = comtypes.client.CreateObject('Word.Application')
    doc = word_doc.Documents.Open(path_in)
    doc.SaveAs(path_out, FileFormat=17)
    doc.Close()
    word_doc.Quit()

def append_paragraph(doc, ret):
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    ret.append(paragraph)

def search_replace_content(paragraphs, data):
    for p in paragraphs:
        for key, val in data.items():
            # Placeholder format: ${PlaceholderName}
            key_name = '${{{}}}'.format(key)
            if key_name in p.text:
                inline = p.runs
                started = False
                key_index = 0
                found_runs = list()
                found_all = False
                replace_done = False
                for i in range(len(inline)):
                    if key_name in inline[i].text and not started:
                        found_runs.append((i, inline[i].text.find(key_name), len(key_name)))
                        text = inline[i].text.replace(key_name, str(val))
                        inline[i].text = text
                        replace_done = True
                        found_all = True
                        break

                    if key_name[key_index] not in inline[i].text and not started:
                        continue

                    if key_name[key_index] in inline[i].text and inline[i].text[-1] in key_name and not started:
                        start_index = inline[i].text.find(key_name[key_index])
                        check_length = len(inline[i].text)
                        for text_index in range(start_index, check_length):
                            if inline[i].text[text_index] != key_name[key_index]:
                                break
                        if key_index == 0:
                            started = True
                        chars_found = check_length - start_index
                        key_index += chars_found
                        found_runs.append((i, start_index, chars_found))
                        if key_index != len(key_name):
                            continue
                        else:
                            found_all = True
                            break

                    if key_name[key_index] in inline[i].text and started and not found_all:
                        # check sequence
                        chars_found = 0
                        check_length = len(inline[i].text)
                        for text_index in range(0, check_length):
                            if inline[i].text[text_index] == key_name[key_index]:
                                key_index += 1
                                chars_found += 1
                            else:
                                break
                        found_runs.append((i, 0, chars_found))
                        if key_index == len(key_name):
                            found_all = True
                            break

                if found_all and not replace_done:
                    for i, item in enumerate(found_runs):
                        index, start, length = [t for t in item]
                        if i == 0:
                            text = inline[index].text.replace(inline[index].text[start:start + length], str(val))
                            inline[index].text = text
                        else:
                            text = inline[index].text.replace(inline[index].text[start:start + length], '')
                            inline[index].text = text

def docx_fill_template(doc, data):
    paragraphs = list(doc.paragraphs)
    append_paragraph(doc, paragraphs)
    search_replace_content(paragraphs, data)