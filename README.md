# Word-Parser

Ever had a bunch of MS Word files you wanted to format in a particular way? Sat down and manually edited them for hours wishing there was an application to do it all at once?

Word-Parser is a Python application which does just that! Simply define your formatting rules, add your MS Word files to a folder and get all the edited files instantly.

**Note**: This application was created as a solution to an issue I faced, so the formatting rules are specified to my use-case.

### Built With

#### Technologies
* [Python 3.9](https://www.python.org/downloads/release/python-390/)
* [Docker](https://www.docker.com/)

#### Libraries
* [docx2python](https://pypi.org/project/docx2python/)


## How to use

Format your MS Word files in three simple steps:

* Copy your MS Word files into the `source-files` directory.
* Run the command: `docker build -t word-parser .`
* Next, run the command: `docker run -it word-parser /bin/bash`
* Within the container, run `python main.py`
* Copy the formatted files from the container using `docker cp <container_id>:/app/output <your_local_path>\word-parser\output`
* View your formatted files in the `output` directory


## Limitations

If you were tricked into thinking this was a perfect project, I wouldn't blame you. Here are some of the limitations:

* Accepts only .docx documents.
* Formats documents into .txt files


Please make sure to update tests as appropriate.

## License

[MIT](https://choosealicense.com/licenses/mit/)