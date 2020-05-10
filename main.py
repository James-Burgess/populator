import PySimpleGUI as sg

import jinja2schema
from docx import Document
from docxtpl import DocxTemplate
import time
from datetime import datetime

def error_msg(err):
    sg.Popup('Error', err)


def get_template():
    layout = [[sg.Text('Please Select Template document')],
                [sg.Input(), sg.FileBrowse()],
                [sg.OK(), sg.Cancel()]]

    window = sg.Window('Select template', layout)

    event, values = window.Read()
    window.close()

    if values[0]:
        return values[0]
    else:
        error_msg('You did not select a document')
        raise


def read_vars(file_path):
    f = open(file_path, 'rb')
    document = Document(f)
    f.close()

    fullText = []
    for para in document.paragraphs:
        fullText.append(para.text)
    temp = '\n'.join(fullText)

    variables = jinja2schema.infer(temp)
    return variables.keys()


def get_data(keys):
    if not keys:
        error_msg('Could not get data, must be a bad file. shutting down')
        raise

    layout = [[sg.Text('Ready to Input information')],]

    for key in keys:
        layout.append([sg.Text(key), sg.InputText()])
    layout.append([sg.Button('Ok'), sg.Button('Cancel')])


    # Create the Window
    window = sg.Window('Input data to populate').Layout([[sg.Column(layout, size=(400,400), scrollable=True)]])

    event, values = window.read()
    print(f"event {event}, vales {values}")

    if event == 'Ok':
        window.close()
        ret = {}
        for k, v in zip(keys, values):
            ret[k] = str(v) or "{{ " +k+" }}"

        print(ret)
    elif event in (None, 'Cancel'):
        error_msg('Could not get data, must be a bad file. shutting down')
        raise
    else:
        error_msg('Could not get data, must be a bad file. shutting down')
        raise
    print(ret)
    return ret


def save_file(template, pop_context):
    layout = [[sg.Text('Please Select save location')],
                [sg.Input(), sg.FolderBrowse()],
                [sg.Text("Filename:"), sg.InputText('populated')],
                [sg.OK(), sg.Cancel()]]

    window = sg.Window('Select save', layout)
    event, values = window.Read()

    if values[0]:
        save = values[0]
        doc = DocxTemplate(template)
        context = pop_context
        doc.render(context)
        save_path = f"{save}/{datetime.now().strftime('%m%d_%H%M%S')}_{values[1]}.docx"
        doc.save(save_path)
    else:
        error_msg('You did not select a document')
        raise

    window.close()

    sg.Popup('Success', f"Successfully saved populated document to {save_path}")


def main():
    sg.theme('Dark Blue 3')   # Add a touch of color

    try:
        template = get_template()
        keys = read_vars(template)
        print(keys)
        data = get_data(keys)
        save_file(template, data)
    except:
        time.sleep(5)
        return


if __name__ == '__main__':
    main()
