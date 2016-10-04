#!/usr/bin/env python3

"""
    Create Calendar for TeamStuff App with AWBB calendar datas
    Author: Nicolas Van Wallendael <nicolasvanwallendael@icloud.com>
    Copyright (C) 2016, Nicolas Van Wallendael
"""



from pyexcel_xls import get_data
import json
import argparse
import csv


# Parser

parser = argparse.ArgumentParser()


# Input - Files

parser.add_argument("-s", "--salleXLS", default="",
                    help="The .XLS file containing the salle infos assignations")

parser.add_argument("-c", "--calendarXLS", default="",
                    help="The .XLS file containing the salle infos assignations")


# Cosmetic Options

parser.add_argument("-v", "--verbosity", action="count", default=0,
                    help="Verbosity lvl (0*v=mute, 1*v=visual, 2*v=text, 3*v=graph, 4*v=text+graph, 5*v=visual+graph)")

args = parser.parse_args()



# Global

# ["Matricule club", "Nom club", "Courriel secr\u00e9taire", "Nom de la salle", "Adresse de la salle", "T\u00e9l\u00e9phone"]
# ['Jour', 'Date', 'Heure', 'Domicile', 'Visiteur', 'Score Domicile', 'Score Visiteur', 'Salle', 'Numero']

CLUB    = "Nom club"

SALLE   = "Nom de la salle"

ADRES   = "Adresse de la salle"

UBW     = "United Basket Woluwé"


#print(get_data(args.salleXLS))


sallesL = [ values for key, values in get_data(args.salleXLS).items()][0] # only get first SHEET
calendL = [ values for key, values in get_data(args.calendarXLS).items()][0] # only get first SHEET

salles = {}

for [_, club, _, salle, adrs] in sallesL :
    
    club = club.strip(" ") if club is not '' else club_old
    
    #print(club, salle, adrs)

    if club not in salles:
        
        salles[club] = {salle : adrs}
    
    else :
        
        # Check for duplicate SALLE name for same club
        if salle in salles[club]:
            exit(-1)
        else :
            salles[club][salle] = adrs

    club_old = club


#print(salles)



# Out :

# Type,Date,Time,Arrival,Duration,Opposition,Venue Name,Venue Address,Special Info
# Home,2016-09-17,14:15,45,120,Castors Braine A,SALLE MALOU,"""Rue Joseph Aernaut 9, 1200 Woluwe-St-Lambert""",

#print(calendL)

calend = [["Type", "Date", "Time", "Arrival", "Duration", "Opposition", "Venue Name", "Venue Address", "Special Info"]]

for [_, date, time, home, away, _, _, salle, _] in calendL :
    
    if not (UBW in home or UBW in away) or home == "" or away == "" :
        continue

    print("HOME = " + home , " --- AWAY = " + away)
    
    home = home.strip(" ")
    home = home[:-4] if home[-1] == ")" else home
    away = away.strip(" ")
    away = away[:-4] if away[-1] == ")" else away
    
    


    
    type = "Home" if UBW in home else "Away"
    arri = 45     if UBW in home else 60
    opp  = away   if UBW in home else home

    print(type, date, time, arri, 120, opp, salle)

    calend.append( [type, date, time, arri, 120, opp, salle, salles[home][salle], "" ] )

print(calend)

# OUT

with open(args.calendarXLS[:-4] + ".csv", "w") as f:
    writer = csv.writer(f)
    writer.writerows(calend)










