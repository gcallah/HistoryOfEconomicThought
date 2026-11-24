# HistoryOfEconomicThought

Class material from my history of economic thought course.

Course site: https://gcallah.github.io/HistoryOfEconomicThought/

## Building the site

The site is generated from course material kept outside this repo, in
`~/Documents/CoursesTaught/HistoryOfThought`: lecture decks (`.pptx`), handouts
(`.docx`, `.rtf`), spreadsheets (`.xlsx`) and quizzes (`.docx`).

Run `make html` to regenerate `index.html`, `lectures/`, `handouts/`,
`spreadsheets/` and `quizzes/`. Override the source location with
`make html SRC=...`. Requires `pandoc` and `python-pptx` (`pip install python-pptx`).

What goes on the site, and in what order, is set by the tables at the top of
`bin/build.py`: `DECKS`, `HANDOUTS`, `SPREADSHEETS`, `QUIZZES` and the weekly
`TOPICS` schedule. Colours, fonts and widths are set once, in `:root` at the top
of `style.css`.
