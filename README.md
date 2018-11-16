# Enumerate Sheets (enumsheets.py)

Counts the total number of dxf files which have specially crafted
title block. Makes copies of those files with title blocks updated
with the values for sheet number, total number of sheets, date,
scale, address. Extracts titles from those title blocks, creates
contents and saves it to Excel file.

This script was created to ease the final stage in production
process of design drawings for the needs of http://artidea.gallery
interior design studio.

This script was tested only with dxf files created and modified
with [LibreCAD](https://librecad.org). Never tested it with files
created using another CAD software.

In order to be processed by the script a dxf file must contain a
specially prepared block (see 'examples' directory). That block:

* must be an INSERT in that dxf file;
* must contain some unique marker (_artidea.gallery_ in my case).

If such a block is found, this script tries to find corresponding
title block fields. If title block is just inserted from my template,
those fields would contain the following ready for replacement
placeholder markers:

- **X** - for the '_Sheet number_' field;
- **XX** - for the '_Number of sheets_' field;
- **TitleField** - for the '_Sheet title_' field;
- **AddressField** - for the '_Address_' field;
- **1:50** - for the '_Scale_' field;
- **0000-00-00** - for the '_Sheet date_' field.

In my workflow, during sheet editing the '_TitleField_' is usually manually
replaced by some text starting from words 'План' or 'Развёртка' (in Russian).
Yours is obviously different. You can adjust corresponding regular expression
in configuration file. Also you probably need to adjust regular expression
for the '_Address_' field.

This script should be able to process not only freshly inserted from
template title block, but also modified title block whose field markers
are already replaced with real values. As far as '_Sheet number_' and
'_Number of sheets_' may contain values of the same pattern, some dumb logic
is used to guess which value belongs to what field.


## The idea behind

I use LibreCAD for creation of my drawings. Since LibreCAD cannot
manage multiple drawings as one single multipage document, I have to
do that myself. Each drawing is a single dxf file. My usual projects
contain of 50-60 such files. When working on a project the final
total number of those drawings is unknown. Even after all of them are
ready, some may be removed or added or reordered. So, there is no way
to know beforehand the exact total number of sheets and their serial
numbering. When all sheets are ready I had to do manual insertion of
sheet numbers and total number of sheets to each sheet title block.
I had to copy every sheet title and create contents file manually.
That required a couple of hours of boring repetitive manual work. Even worse,
when after doing all that job you suddenly realize that you need a couple
of new drawings more somewhere in the middle. Or if during numbering you
make a mistake. That's horrible. This script automates the job and
produces results in seconds.


## My usual workflow

An initial drawing is drawn. When it is ready additional blocks
are inserted into the drawing file, including the block with sheet
border and the title block (see 'templates' directory for my templates
examples). I edit _TitleField_ in inserted title block and optionally
set the _Scale_ field. Before saving the prepared sheet I also do print
preview in LibreCAD, set printing scale according to my sheet's scale, and
also center the drawing on the page. After all that I save the drawing and
close it. My final dxf files names usually include a number, e.g.
sheet-001.dxf, sheet-002.dxf, and so on to ease their ordering.

When all sheets are ready I make a copy of the script's _config.ini_
in my working directory, edit it in order to add the address lines
(_address_value_ parameter) and optionally set other configuration
parameters. Then I run the script with that _config.ini_ and have a bunch
of copies of my sheets with title blocks filled with the correct values,
and also an Excel file with all titles exracted from those sheets.
By using another tools, those prepared and numbered sheets are converted
to one big pdf file ready to be printed or given to a client.


## Initial setup

Create virtual environment in a directory of your choice and activate it
by running:

```
virtualenv -p /usr/bin/python3 venv
. venv/bin/activate
```

Install required dependencies:

```
pip install -r requirements.txt
```


## Usage

```
enumsheets.py [-h] [-c CONFIG] dxf_file [dxf_file ...]
```


### Helper enumsheets.sh script

There is also **enumsheets.sh** helper script. It activates
Python virtual environment and after that runs the **enumsheets.py**
script with the same arguments provided on the command line.
If Python virtual environment was not created yet this script
creates one in the same directory where both scripts are located.


## How you can use it

First, take a look at 'templates' and 'examples' directories. Modify
templates according to your needs and set a unique marker. Change
the **marker** parameter in the configuration file according to
what you set in your template. Also edit patterns parameters in
the configuration file in order to make the script to be able to
find your fields inside your title block.

**To be clear, sheet templates in 'templates' directory are not to
be used as base for new drawings. Only as imported block to your
existing drawing. That is important. Otherwise, the script will
not be able to find the title block inside your drawings.**

_Have fun!_

