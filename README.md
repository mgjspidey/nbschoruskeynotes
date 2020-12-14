# NBS Chorus to Revit Keynote

The python script takes a given NBS Chorus Word doc (using the default NBS Teamplate) and then converts the date into a Revit Keynote file. It does this by using the paragraph styles for the section and clause titles. THese are then concatenanted with the section / clause numbers.

## Known Limitations

Currently this has been setup for CAWS specifications ONLY.
When I get some time I will upate this to allow for UniClass organisation as well

## Running the Script

If you have Python installed simply download and ensure the following packages are present
* Appjar - for gui
* docx - for reading the dox format

Alternatively you can download the packaged version and run the enclosed exe (Packaged using pyinstaller)
