#!/bin/sh
#
# enumsheets.sh - Helper script to run enumsheets.py script.
#
# Copyright (C) 2018 Alexander Pravdin <aledin@mail.ru>
#
# License: MIT
#

PROGDIR=$(dirname ${0})

if [ -e ${PROGDIR}/venv/bin/activate ]; then
    . ${PROGDIR}/venv/bin/activate || exit 1
else
    echo "Initializing python virtual environment..."
    virtualenv -p /usr/bin/python3 ${PROGDIR}/venv || exit 1
    . ${PROGDIR}/venv/bin/activate || exit 1
    pip install -r ${PROGDIR}/requirements.txt || exit 1
    echo "Python virtual environment initialized"
    echo
fi

${PROGDIR}/enumsheets.py ${@}

deactivate

