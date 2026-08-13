# Moodle Course Content Mapper

Create an interactive, browsable and editable representation of an
existing Moodle course from the processed outputs of [Moodle Course
Auditor](https://github.com/cloudpedagogy/cloudpedagogy-moodle-course-auditor).

Moodle Course Content Mapper preserves the current Moodle section and
**course-item structure** and connects course items to extracted course
resources. **Course item** is used as an umbrella term for elements
represented in Moodle; where possible, the mapper distinguishes
**resources/content** (such as files, URLs, pages and books) from
**learning activities** (such as forums, quizzes and assignments). Its
**primary output is an interactive HTML course map** for exploring the
existing course, with navigation, filtering and clickable links to
mapped resources.

It also generates an **editable Microsoft Word version** of the map for
human-led course review and redesign. In Word, headings, course items,
explanatory text and linked resource references can be moved, renamed,
annotated and reorganised to sketch a proposed new structure while
retaining links back to the source resources.

The mapper is intentionally conservative: it represents what is present
in the source course and avoids automatically deciding how the course
should be redesigned.

## Example course map

![Interactive Moodle course content map showing the existing section,
course-item and resource structure](img/screenshot.png)

*Example of the primary HTML output generated from a Moodle course. The
map represents the current Moodle structure and provides interactive
navigation, filtering and clickable links to mapped course resources.*

## Where the data comes from

Moodle Course Content Mapper is a **downstream companion tool** to
[CloudPedagogy Moodle Course
Auditor](https://github.com/cloudpedagogy/cloudpedagogy-moodle-course-auditor).

Moodle Course Auditor analyses Moodle backup (`.mbz`) files and produces
structured audit data plus extracted course resources. The analysis is
**non-destructive**: it works on a static backup file and does not
connect to, alter or write back to the live Moodle course.

Moodle backups are normally associated with backup and restore
workflows, but they also contain a rich source of information about a
course, including its sections, course items, resources, files, links,
visibility, structure and other metadata. Moodle Course Auditor exposes
that information in reusable CSV/JSON reports and, when file extraction
is enabled, recovers the associated course resources.

The Content Mapper then uses those outputs to create a human-readable
representation of the course:

``` text
Live Moodle course
        ↓
Moodle backup (.mbz)
        ↓
Moodle Course Auditor
(non-destructive analysis of the static backup)
        ↓
audit/
extracted_files/
        ↓
Moodle Course Content Mapper
        ↓
index.html          → interactive view of the current course
content_map.docx    → editable working copy for redesign
content_map.csv     → structured mapping data
mapping_report.md   → QA and provenance
```

This separation is deliberate:

-   **Moodle Course Auditor** answers: *What is in this Moodle backup?*
-   **Moodle Course Content Mapper** answers: *How can that existing
    course structure and its resources be presented clearly for review
    and redesign?*

## Course review and redesign

The mapper is designed to provide a bridge between **Moodle course
audit** and **human-led course redesign**.

``` text
Existing Moodle course
        ↓
Moodle Course Audit
        ↓
Evidence about structure + extracted resources
        ↓
Moodle Course Content Mapper
        ↓
Interactive HTML map of the current course
        ↓
Editable Word redesign workspace
        ↓
Human-reviewed proposed course structure
```

The HTML output provides a clear view of the **current Moodle course**:
sections, course items and associated resources can be explored without
changing Moodle itself. It also acts as a lightweight **course-review
layer**, using available Course Auditor metadata to add search, dynamic
filters and optional review information without changing the underlying
course.

The HTML remains deliberately data-driven. Filters and metadata are
shown only when the source audit provides useful values. Missing
metadata is not guessed or treated as an error, so courses with
different combinations of resources and activities can be processed by
the same mapper.

The Word output supports the next stage. Because it is editable,
academics and learning designers can rearrange headings, course items,
explanatory text and linked resources to explore alternative structures.
The links to the bundled source files remain available as reference
points during that process.

The tool therefore **supports course redesign; it does not automatically
redesign the course**. Pedagogical interpretation and decisions remain
with the course team.

A useful way to think about the two main outputs is:

-   **HTML = understand, search, filter and review the existing Moodle
    course.**
-   **Word = prototype and discuss a possible redesigned structure.**

## Terminology

The mapper uses **course item** as the user-facing umbrella term for
elements represented in the Moodle course. Moodle distinguishes between
resources/content and learning activities, so the mapper presents that
distinction where the source activity type allows it:

-   **Resources/content** --- for example files, URLs, pages, books,
    folders and labels.
-   **Activities** --- for example forums, quizzes, assignments,
    lessons, H5P, SCORM and LTI items.

The underlying Moodle Course Auditor fields such as `activity_type` and
`activity_name` remain unchanged for compatibility. This is therefore a
user-facing terminology improvement rather than a change to the audit
data model.

## Features

-   Preserves Moodle section order and course-item order.
-   Uses **course item** as the user-facing umbrella term while
    retaining the Course Auditor data model.
-   Distinguishes resources/content from learning activities where the
    Moodle item type allows this.
-   Uses existing Moodle labels as subheadings.
-   Maps resources and URLs to their corresponding Moodle course items.
-   Creates clickable links to extracted PDFs, Word files, spreadsheets
    and other resources.
-   Creates an editable Microsoft Word working copy for course review
    and redesign.
-   Creates a browsable HTML course map with full course-item labels.
-   Provides text search across mapped course items and resources.
-   Generates **dynamic review filters** from the metadata actually
    present in the course, such as item category/type, provider,
    local/external location, file type, modification year, link QA and
    visibility.
-   Shows optional per-item **Review metadata** where supported by the
    Course Auditor data.
-   Can expose explicit duplicate candidates when the Course Auditor
    provides activity-level duplicate identifiers.
-   Suppresses unavailable or non-useful filters rather than assuming
    that every Moodle course contains the same resource types or
    metadata.
-   Does not infer missing dates, duplicate status or other review
    metadata.
-   Uses Moodle labels as subheadings but does not count those labels as
    displayed course items.
-   Supports a portable bundle containing copied resources.
-   Uses conservative filename matching and reports ambiguous or
    unresolved items rather than guessing.
-   Can optionally include hidden Moodle content.

## Robustness and portability

The mapper is designed to work across different Moodle Course Auditor
outputs rather than being tailored to a single course.

The core course map depends on the required section, activity and
placement data. Analytical metadata is treated as optional enrichment.
If a particular course does not contain a provider, file type,
modification year, duplicate flag or other optional value, the relevant
filter/detail is omitted and the rest of the map continues to render.

This approach keeps the mapper conservative:

-   source order is preserved;
-   resource associations use accuracy-first matching;
-   ambiguous mappings are reported rather than guessed;
-   item-level dates are shown only when supported by item-level data;
-   duplicate status is shown only when supported by explicit
    identifiers;
-   optional metadata never becomes a requirement for rendering the map.

## Requirements

-   Python 3
-   `python-docx`
-   A processed Moodle Course Audit course folder containing `audit/`
    and `extracted_files/`

## Installation

Clone or download the repository, then from the repository root create a
virtual environment:

``` bash
python3 -m venv .venv
source .venv/bin/activate
```

Install dependencies:

``` bash
python3 -m pip install -r requirements.txt
```

## Input structure

Place the complete processed course folder inside `input/`.

Example:

``` text
input/
└── literature-review/
    ├── audit/
    │   ├── sections.csv
    │   ├── activities.csv
    │   ├── content_placement_inventory.csv
    │   └── ...
    └── extracted_files/
        ├── files_by_moodle_context/
        ├── resource_bundle/
        ├── resource_bundle_by_type/
        ├── resource_manifest.csv
        └── ...
```

Do not manually reorganise the `audit/` or `extracted_files/` folders
before mapping.

## Usage

Basic mapping:

``` bash
python3 content_mapper.py input/literature-review
```

Recommended portable mapping:

``` bash
python3 content_mapper.py input/literature-review --bundle
```

The `--bundle` option copies linked extracted resources into the output
`resources/` directory. This is recommended when the HTML/Word map will
be moved or shared together with its resources.

Other options:

``` bash
python3 content_mapper.py input/literature-review --include-hidden
python3 content_mapper.py input/literature-review --title "Reviewing the Literature"
python3 content_mapper.py input/literature-review --output-dir output/custom-name
python3 content_mapper.py --help
```

## Outputs

A typical bundled output is:

``` text
output/
└── literature-review/
    ├── index.html
    ├── content_map.docx
    ├── content_map.csv
    ├── mapping_report.md
    ├── unresolved_items.csv
    └── resources/
```

### `index.html` --- primary interactive output

The main output of the mapper. It provides an interactive, browsable
representation of the **current Moodle course structure**, including its
sections, course items and mapped resources.

The HTML interface includes course navigation, text search and
**data-driven review filters**. Mapped local resources and external URLs
are clickable. When `--bundle` is used, links to local files point to
the portable copies in `resources/`.

Depending on the metadata available for a particular course, the HTML
may offer filters for:

-   course-item category (resource/content or activity);
-   Moodle item type;
-   provider or hosting source;
-   local/bundled versus external links;
-   file type;
-   modification year;
-   link/mapping QA status;
-   visibility; and
-   explicit duplicate candidates.

These controls are **dynamic rather than hard-coded**. A filter is only
shown when the mapped course contains useful values for it. For example,
a course with no relevant video/file-type metadata will not gain a
video-specific option merely because another course contains videos.
Likewise, missing modification dates or duplicate information are not
inferred.

### What the review filters mean

The filters are **dynamic and data-driven**. They are generated from the
metadata actually present in the processed course rather than from a
fixed list of expected Moodle content. A filter is shown only when the
source data contains useful values for it. For example, if a course has
no relevant video or file-type metadata, no video-specific option is
created; if item-level modification dates are unavailable, the
modification-year filter is omitted.

  -----------------------------------------------------------------------
  Filter                              What it filters
  ----------------------------------- -----------------------------------
  **Category**                        Broad course-item family, such as
                                      resource/content or learning
                                      activity.

  **Item type**                       Moodle item type, such as File,
                                      URL, Forum, Quiz, Assignment, LTI
                                      or another type present in the
                                      course.

  **Provider**                        Identified provider or hosting
                                      source, such as Moodle or Panopto,
                                      when supplied by the audit data.

  **Link location**                   Whether a mapped link is
                                      bundled/local or points to an
                                      external resource.

  **File type**                       File extension of mapped resources,
                                      such as PDF, DOCX, XLSX or MP4,
                                      when present.

  **Modified year**                   Item-level modification year, but
                                      only when supported by the source
                                      audit data.

  **Link QA**                         Mapping/link status, such as
                                      linked, unresolved, ambiguous or no
                                      separate resource.

  **Visibility**                      Moodle visibility state when
                                      available, for example visible or
                                      hidden.

  **Duplicate candidates**            Restricts the view to items
                                      explicitly identified as possible
                                      duplicates by Course Auditor data.
  -----------------------------------------------------------------------

Filters can be combined with each other and with text search to narrow
the course map during review. They affect only the displayed static map;
they do not change Moodle or the underlying audit data.

Where useful metadata exists, each course item can expose a collapsible
**Review metadata** panel. This allows the HTML output to act as an
analysis/review layer over the current Moodle structure while remaining
a static, portable representation.

The bundled HTML and local resources can be reviewed without Moodle
access and can work offline. External services and URLs, such as Panopto
or SharePoint, still require an internet connection.

This output is intended primarily for exploring and reviewing the
existing course before or during redesign.

### `content_map.docx` --- editable redesign workspace

Editable Microsoft Word representation of the course map for the next
stage of review and redesign.

Unlike the HTML map, which is primarily for exploring the current
course, the Word version can be actively rearranged. Course teams can
move and rename headings, course items and explanatory text; move linked
resource references into a proposed new sequence; add notes; and use the
document as a working outline for a redesigned course.

Clickable links to the underlying bundled resources are retained, so a
reviewer can continue opening the relevant PDFs, Word documents,
spreadsheets and other mapped resources while developing the proposed
structure.

### `content_map.csv`

Machine-readable mapping of the course structure and resources for
further analysis.

### `mapping_report.md`

QA and provenance report summarising the mapping process and
resource-link resolution.

### `unresolved_items.csv`

Items that could not be linked reliably. These are surfaced for review
rather than guessed.

### `resources/`

Created when `--bundle` is used. Contains portable copies of the linked
extracted resources.

## Resource matching

Resource resolution follows an accuracy-first sequence:

1.  Exact filename match.
2.  Case-insensitive exact filename match.
3.  Conservative normalised filename match.
4.  Very-high-confidence fuzzy filename match using the same file
    extension.

Ambiguous matches are deliberately left unresolved.

## Design principle

The mapper separates **evidence mapping** from **pedagogical redesign**.

It answers:

> What is in the Moodle course, where is it currently located, and which
> resources belong to it?

It does not currently answer:

> Where should this content move under a particular pedagogical model?

The Word working copy provides a bridge to that next stage by giving
academics and learning designers an editable representation that can be
manually reorganised without changing the source Moodle backup or
extracted resources.

## Recommended workflow

1.  Back up the Moodle course as `.mbz`.
2.  Process the backup using Moodle Course Audit with resource
    extraction enabled.
3.  Copy the resulting processed course folder into `input/`.
4.  Run Moodle Course Content Mapper with `--bundle`.
5.  Review `mapping_report.md` and `unresolved_items.csv`.
6.  Browse the source structure in `index.html`.
7.  Use `content_map.docx` as an editable working document for course
    review or restructuring.

See [INSTRUCTIONS.md](INSTRUCTIONS.md) for the step-by-step operational
workflow.

## Status

Early development / experimental.

Current mapper release documented here: **v0.5.x**.

The current focus is reliable structural mapping, resource linking and
editable course-review outputs. Future development may introduce
configurable redesign templates or pedagogical mapping while retaining
the original Moodle structure as source evidence.

## Licence

MIT
