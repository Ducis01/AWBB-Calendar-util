#!/bin/bash

# Run for every files

for file in ./Files/*; do

    if [[ $file == *".xls" ]]
    then

        python3 reg_calendar.py -s salles_clubs_AWBB.xlsx -c ${file}
    fi

done
