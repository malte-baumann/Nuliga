import data
import excel
from datetime import datetime
import re

def roman_to_int(s):
    roman_values = {'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100, 'D': 500, 'M': 1000}
    result = 0
    for i in range(len(s)):
        if i > 0 and roman_values[s[i]] > roman_values[s[i - 1]]:
            result += roman_values[s[i]] - 2 * roman_values[s[i - 1]]
        else:
            result += roman_values[s[i]]
    return result

def sort_team(s):
    parts = s.split()
    if len(parts) > 2 and parts[-1].isdigit():
        return (parts[0], parts[1], int(parts[-1]))
    elif len(parts) > 2 and re.match(r'[IVXLCDM]+', parts[-1]):
        return (parts[0], parts[1], roman_to_int(parts[-1]))
    else:
        return (parts[0], parts[1], 0)

if __name__ == "__main__":
    search_term = input("Suche nach Verein: ")
    club_site = data.search_club(search_term.lower())
    id = data.get_club_id(club_site)

    now = datetime.now()
    year = now.year
    if (now.month < 6):
        year = now.year -1 

    saison = f"{year}/{str(year + 1)[2:]}" 
    
    runde = "vorrunde"
    if (now.month < 6):
        runde = "rueckrunde"

    damen = data.get_rangliste(id, runde, "Damen", saison)
    herren = data.get_rangliste(id, runde, "Herren", saison)
    rangliste = {"damen": damen, "herren": herren}

    spiele, alle_mannschaften = data.get_alle_spiele(id, search_term)
    teams = [s for s in alle_mannschaften if search_term.lower() in s.lower()]
    sorted_teams = sorted(teams, key=sort_team)

    excel.create_sheet(rangliste, alle_mannschaften, sorted_teams, spiele, runde, f"spieltermine_{search_term.replace(" ", "-")}.xlsx")