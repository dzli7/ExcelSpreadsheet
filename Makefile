# PYLINT = pylint
# PYLINTFLAGS = -rn

PYTHONFILES := $(wildcard *.py)

pylint: $(patsubst %.py,%.pylint,$(PYTHONFILES))

%.pylint:
		pylint -rn $*.py

.PHONY: clean
clean:
		rm -rf *.o

.PHONY: lint
lint:
		@pylint --rcfile=.pylintrc sheets 