# ordina-cv #

## Setup ##

The script can be run in python3 with the docx2txt package installed.

To set up this environment with anaconda:
```bash
conda create -n cv_tool_env python=3.7 
activate cv_tool_env
pip install -r requirements.txt 
```

## How to run ##

* On a single CV:
```bash
python scrape_cv.py [input_file.docx]
```
This will create input_file.json and img_input_file directory with the scraped images.

* On a directory containing CVs:
```bash
python scrape_cv.py [input_dir]
```
This will do the same for every docx file in input_dir.
