# Build the static site from course material kept outside this repo.
# Override the source location on the command line if it moves, e.g.
#   make html SRC=/path/to/HistoryOfThought
# Needs python-pptx (pip install python-pptx) and pandoc.
SRC ?= $(HOME)/Documents/CoursesTaught/HistoryOfThought

FORCE:

html: FORCE
	python3 bin/build.py --src "$(SRC)"

clean: FORCE
	rm -rf lectures handouts spreadsheets quizzes index.html
