'''
Script to convert the WORD export from NBS Chorus to Keynotes for Revit
created by matt.g.jones
version 0.1

The Keynote output needs 3no. columns TAB separated
[UNIQUE NBS REF][NBS CLAUSE / SECTION TITLE][PARENT REF]
'''

from docx import Document  # Import 'Document' library from python-docx module to read the WORD doc
from appJar import gui  # Import appJar to manage user input / dialogue
import os  # Import OS


def clean_text(text_in):
    '''
    Takes the input text and converts it to uppercase, before splitting
    it into the NBS ref and NBS Title
    '''
    # Convert to uppercase
    text_in = text_in.upper()

    # Clean up any hard returns and replace them with a space
    text_in = str(text_in).replace('\n', ' ')

    # Split the NBS ref from the front of the text string
    nbs_ref = text_in.split(' ')[0]

    # Take the rest of the string as the NBS Title
    nbs_text = ' '.join(text_in.split(' ')[1:])

    if list(text_in)[len(nbs_ref) + 2] == ' ':
        nbs_ref = nbs_ref + list(text_in)[len(nbs_ref) + 1]
        nbs_text = nbs_text[2:]

    # Return the two values
    return [nbs_ref, nbs_text]


def get_nbs_section_titles(nbs_sectiontitles_path):
    '''
    Pulls in the section titles from an external text file
    '''
    with open(nbs_sectiontitles_path) as f:
        nbs_section_titles = f.readlines()

    # Strip whitespace etc
    nbs_section_titles = [x.strip() for x in nbs_section_titles]

    # Convert into Keynote format
    nbs_titles = []
    for nbs_title in nbs_section_titles:
        temp = clean_text(nbs_title)
        nbs_titles.append([temp[0], temp[1], ''])

    return nbs_titles


def process_doc(file_input, nbs_sectiontitles_path):
    '''
    Takes the docx (file_input) and pulls out the 2no. paragraph styles that
    provide the NBS Section and NBS Clauses. The output is a TAB-separated text file
    '''

    # Setup empty list to hold the data
    nbs_clauses = []

    # Open the docx file from the given location
    doc = Document(file_input)

    # Process through the paragraphs reading the key paragraph styles ONLY
    for paragraph in doc.paragraphs:
        if paragraph.style.name == 'chorus-section-header':
            # Clean the text and split it into two values using the function
            nbs_section = clean_text(paragraph.text)

            # Append the 3no. values to the list
            nbs_clauses.append([nbs_section[0], nbs_section[1], nbs_section[0][0]])

        if paragraph.style.name == 'chorus-clause-title':
            # Clean the text and split it into two values using the function
            nbs_clause = clean_text(paragraph.text)

            # We need to prefix the NBS Clause Ref with the NBS Section Ref
            nbs_ref = str(nbs_section[0] + '/' + nbs_clause[0])

            # Append the 3no. values to the list
            nbs_clauses.append([nbs_ref, nbs_clause[1], nbs_section[0]])

    # Check if any clauses have been found (if not perhaps not an NBS Chorus document!)
    if len(nbs_clauses) <= 1:
        show_error_message('The document provided does not contain the NBS Chorus Information in the correct format, please check the file is using the standard NBS Chorus Template and try again')
        app.stop()
    else:
        # Add titles
        nbs_clauses = nbs_clauses + get_nbs_section_titles(nbs_sectiontitles_path)

        # Sort the data correctly
        nbs_clauses.sort()

        # Create the file_output path from the file_input path by replacing the file suffix
        file_output = file_input.replace('docx', 'txt')

        # Create the output file and iterate over list of clauses
        with open(file_output, 'w') as textfile:
            for nbs_clause in nbs_clauses:
                # Write each clause as a new line with the 3no. values separated by a TAB
                textfile.write(nbs_clause[0] + '\t' + nbs_clause[1] + '\t' + nbs_clause[2] + '\n')

        # Notify the user and close the app
        close_app()


def setup_dialogue(app, nbs_sectiontitles_path, app_width, app_height):
    app.setGuiPadding(10, 0)
    app.setPadding([0, 0])
    app.setInPadding([0, 0])
    app.setSticky('ew')

    row = 0
    app.addMessage(
        'Info',
        'This app will take the NBS Chorus WORD file (docx) and convert it to a Revit Keynote file. The Keynote file will be saved next to the NBS Chorus Word file.',
        row)
    app.setMessageWidth('Info', app_width - 50)

    row += 1
    app.addEmptyLabel("0")  # Padding
    app.addLabel('FilePath', 'Locate the NBS Chorus WORD file:', row, 0)

    row += 1
    app.addEntry('path', row, 0)
    app.setEntryDefault('path', '...')

    row += 1
    app.addButton('Find file', get_file_input, row, 0)

    row += 1
    app.addButtons(["Submit", "Cancel"], press, row, 0)
    app.enableEnter(press)

    row += 1
    app.addWebLink("mattgjones.com", "http://mattgjones.com")

    app.go()


def get_file_input(btnName):
    if btnName == 'Find file':
        file_input = app.openBox(
            title='Select the NBS Chorus Word document - a .docx filetype',
            dirName=os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop'),
            fileTypes=[('document', 'docx')],
            asFile=False,
            parent=None,
            multiple=False,
            mode='r'
        )
        app.setEntry('path', file_input)


def press(btnName):
    if btnName == "Cancel":
        nbs_filepath = 'Null'
        app.stop()
    if btnName == "Submit":
        nbs_filepath = app.getEntry('path')
        if nbs_filepath == '':
            show_error_message('Please provide a filepath to the NBS Chorus document')
        else:
            process_doc(nbs_filepath, nbs_sectiontitles_path)


def show_error_message(message):
    '''
    Dialogue for reporting errors
    '''
    app.infoBox('Error in processing', message, parent=None)


def close_app():
    '''
    Dialogue for confirming the process and closing the app
    '''
    app.infoBox('Conversion completed',
                'The process has been completed and the Keynote file saved',
                parent=None)
    app.stop()


# Execution starts here
if __name__ == "__main__":
    # Setup fileinput path for CAWS sections
    nbs_sectiontitles_path = 'NBS_Sections_Titles.txt'

    # Setup GUI variables
    app_width = 500
    app_height = 300
    app_size = str(app_width) + 'x' + str(app_height)
    this_path = os.path.dirname(os.path.abspath(__file__))
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    app = gui('NBS Chorus to Revit Keynote Converter - by mattgjones.com', app_size)
    app.setFont(size=9)
    app.setButtonFont(size=9)
    app.setIcon(this_path + "\\icon.ico")

    setup_dialogue(app, nbs_sectiontitles_path, app_width, app_height)