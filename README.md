# manuskript-exporter

A Python application for Manuskript data export from outsides.

## Features

- Creates a document containing the story world descriptions. 
  The heading levels reflect the hierarchy in *Manuskript*. 
- Creates a document containing the character data.
  The first level headings show the characters' names. 
  The character information is structured on the second level.
- Creates a document containing all chapters and scenes.
- Creates synopses on all levels (up to 6) of the Manuskript *Outline*:
    - Full chapter summaries in a document per chapter level.
    - Short chapter summaries in a document per chapter level.
    - Full scene summaries.
    - Short scene summaries.
    - Scene titles.
- You can select the output format (md, odt, docx, html). 

## Requirements

- A Python installation (version 3.6 or newer).
- [pypandoc](https://github.com/JessicaTegner/pypandoc) installed via *PyPi*.
- [mskmd.py](https://github.com/peter88213/manuskript_md) in the same directory as *manuskript-exporter.pyw*.

### Note for Linux users

Please make sure that your Python3 installation has the *tkinter* module. 
On Ubuntu, for example, it is not available out of the box and must be installed via a 
separate package named **python3-tk**. 

## Download

Save the file [manuskript-exporter.pyw](https://raw.githubusercontent.com/peter88213/manuskript-exporter/mskexporter/main/mskexporter.pyw).

## Usage


The file *mskmd.py* must be placed besides *manuskript-exporter.pyw*. 

1. Open your *Manuskript* project with **File > Open** or **Ctrl-O**. Select the *.msk* file.
2. Select the output document type.
3. Create the documents selecting the type with the **Convert** menu. 


## License

Published under the [MIT License](https://opensource.org/licenses/mit-license.php)
