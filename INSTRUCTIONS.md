# Moodle Course Content Mapper --- Instructions

Applies to **v0.4.x** of the mapper.

This guide describes the operational workflow for creating a course
content map.

For project purpose, features and technical overview, see `README.md`.

### Terminology

The mapper uses **course item** as an umbrella term for elements represented
in the Moodle course. Where possible, the generated outputs distinguish
**resources/content** (for example files, URLs, pages and books) from
**learning activities** (for example forums, quizzes and assignments).

The source Moodle Course Auditor data still uses internal fields such as
`activity_type` and `activity_name`; these are retained for compatibility.


## 1. Generate the prerequisite course data

Moodle Course Content Mapper expects output from **Moodle Course
Audit**.

Start with a Moodle `.mbz` backup and run Moodle Course Audit with
resource extraction enabled.

The processed course must contain:

``` text
course-name/
├── audit/
└── extracted_files/
```

Both folders are required. The mapper does not currently read the `.mbz`
file directly.

## 2. Copy the processed course into the mapper

From the Moodle Course Content Mapper repository, place the complete
processed course under `input/`.

Example:

``` text
moodle-course-content-mapper/
├── content_mapper.py
├── requirements.txt
├── input/
│   └── literature-review/
│       ├── audit/
│       └── extracted_files/
└── output/
```

Keep the generated folder structure intact.

## 3. Create a Python virtual environment

From the repository root:

``` bash
python3 -m venv .venv
```

Activate it on macOS/Linux:

``` bash
source .venv/bin/activate
```

Your terminal prompt should normally show `(.venv)` when the environment
is active.

## 4. Install dependencies

``` bash
python3 -m pip install -r requirements.txt
```

The Word output requires `python-docx`.

## 5. Run the mapper

For a standard map:

``` bash
python3 content_mapper.py input/literature-review
```

For a portable map with copied resources:

``` bash
python3 content_mapper.py input/literature-review --bundle
```

The bundled form is recommended when the outputs will be reviewed, moved
or shared as a package.

## 6. Check the terminal summary

A successful run reports the course title, numbers of
sections/course items/placements, mapping rows, local resource links,
external links, unresolved issues, HTML map and Word working copy.

Pay particular attention to the unresolved issue count.

## 7. Review the outputs

The output folder will normally contain:

``` text
output/literature-review/
├── index.html
├── content_map.docx
├── content_map.csv
├── mapping_report.md
├── unresolved_items.csv
└── resources/          # when --bundle is used
```

Open the HTML map on macOS:

``` bash
open output/literature-review/index.html
```

Open the Word working copy:

``` bash
open output/literature-review/content_map.docx
```

## 8. Use the HTML and Word outputs differently

Use `index.html` to browse, search and inspect the mapped **current Moodle course structure**.

Use `content_map.docx` as an editable working document. Academics and
learning designers can move headings, course items and linked content references, add notes,
propose new groupings and develop a revised course structure.

The Word document is a working representation; editing it does not
change the Moodle backup, audit data or extracted source resources.

When `--bundle` is used, keep `content_map.docx` together with the generated
`resources/` folder if you want its local resource hyperlinks to continue
working after the output is moved or shared.

## 9. Review unresolved items

Check:

``` text
mapping_report.md
unresolved_items.csv
```

The mapper deliberately avoids guessing when a resource match is
ambiguous.

An unresolved item therefore does not necessarily mean that the resource
is absent. It means that the mapper could not establish a sufficiently
reliable link automatically.

## 10. Useful command options

Include hidden Moodle course items and sections:

``` bash
python3 content_mapper.py input/literature-review --include-hidden
```

Override the detected course title:

``` bash
python3 content_mapper.py input/literature-review --title "Reviewing the Literature"
```

Choose an output location:

``` bash
python3 content_mapper.py input/literature-review --output-dir output/reviewing-the-literature
```

Show all available options:

``` bash
python3 content_mapper.py --help
```

## 11. Recommended quality check before sharing

Before distributing a map:

1.  Check the unresolved count.
2.  Review `unresolved_items.csv`.
3.  Open `index.html` and test a sample of local and external links.
4.  Open `content_map.docx` and test a sample of hyperlinks.
5.  If sharing the map outside the original working directory, use
    `--bundle` and keep the entire output folder together.

## 12. What the mapper does not do

The current version does not automatically redesign the course according
to a pedagogical model.

It preserves and exposes the existing structure so that restructuring
decisions can be made transparently by academics and learning designers.

This distinction is intentional: source mapping remains evidence-based,
while redesign remains a separate human-led stage.
