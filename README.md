# NBS Chorus to Revit Keynote

The python script takes a given NBS Chorus Word doc (using the default NBS Teamplate) and then converts the date into a Revit Keynote file. It does this by using the paragraph styles for the section and clause titles. THese are then concatenanted with the section / clause numbers.

## Standalone App

If you want a standalone version (no python installed) then download these three files and run the exe
* CreateKeynotes.exe - the main executable created by pyinstaller
* icon.ico - the icon for the app
* NBS_Section_Titles.txt - the CAWS master sections

## Running the Script

If you have Python installed simply download the python file and ensure the following packages are present on your computer:
* Appjar - for gui
* docx - for reading the dox format

## Known Limitations

Currently this has been setup for CAWS specifications ONLY.
When I get some time I will upate this to allow for UniClass organisation as well.
