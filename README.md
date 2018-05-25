# ordina-cv #

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
