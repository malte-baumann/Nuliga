import requests
from bs4 import BeautifulSoup
from urllib.parse import urlparse, parse_qs
import exception

NULIGA_URL = "https://hbv-badminton.liga.nu"

def search_club(search_for):
    print(f"Suche nach {search_for} in NULIGA")
    url = "https://hbv-badminton.liga.nu/cgi-bin/WebObjects/nuLigaBADDE.woa/wa/clubSearch"
    payload = {'searchFor': search_for, 'federations': "HBV", 'federation': "HBV"}
    
    try:
        response = requests.post(url, params=payload)
        response.raise_for_status()
        return response.text 
    except requests.exceptions.RequestException as e:
        print(f"An error occurred: {e}")
        raise RuntimeError(e)

# sucht nach allen Begegnungen des Vereins
def get_alle_spiele(club_id, search_term, saison):
    print("Hole alle Spiele")
    spielbetrieb_url = "/cgi-bin/WebObjects/nuLigaBADDE.woa/wa/clubMeetings"
    params = {"club": club_id, "searchType": 0, "selectedTeamId": "WONoSelectionSTring", "searchMeetings": "Suchen"}
    # "searchTimeRange": "13-2607"
    response = requests.post(NULIGA_URL + spielbetrieb_url, params=params)
    response.raise_for_status()

    soup = BeautifulSoup(response.content, 'html.parser')
    
    # Korrekte Saison (Zeitraum) auswählen
    sel = soup.find("select", attrs={"name": "searchTimeRange"})
    opt = sel.find("option", string=saison)

    searchTime = opt["value"] if opt else None
    if (searchTime is None):
        raise exception.SaisonNotFound(f"Saison {saison} in /clubMeetings nicht gefunden")
    
    params["searchTimeRange"] = searchTime
    response = requests.post(NULIGA_URL + spielbetrieb_url, params=params)
    response.raise_for_status()
    soup = BeautifulSoup(response.content, 'html.parser')

    
    spiel_tabelle = soup.find(class_="result-set")
    trs = spiel_tabelle.find_all('tr')

    if (len(trs) < 2): raise exception.EmptyResult("Keine Begegnungen gefunden!")

    liste = {}
    last_date = ""
    mannschaften = set()
    for tr in trs[1:]:
        tds = tr.find_all('td')
        date = tds[1].text.strip()
        time = tds[2].text.strip().split()[0]
        liga = tds[5].text.strip()
        heim = tds[6].text.strip()
        gast = tds[7].text.strip()


        if (date != ""):
            last_date = date
            liste[date] = liste.get(date, [])

        if (liga == "Jugend-WI" or liga == "SchMini-Wi"):
            # Der Mannschaftsname der Jugend / Schüler ist der gleiche wie bei den Aktiven
            continue

        isHeim = False
        mannschaft = gast
        gegner = heim
        if search_term.lower() in heim.lower():
            isHeim = True
            mannschaft = heim
            gegner = gast
        
        mannschaften.add(mannschaft)
        
        if gegner == "spielfrei" or mannschaft == "spielfrei":
            continue

        liste[last_date].append({"time": time, "heim": isHeim, "mannschaft": mannschaft, "gegner": gegner})

    return liste, mannschaften


def get_rangliste(club_id, runde, gender, season):
    rangliste_url = "/cgi-bin/WebObjects/nuLigaBADDE.woa/wa/clubPools"
    params = {"club": club_id,"displayTyp": runde, "contestType": gender, "seasonName": season}
    print("Get Rangliste for: ", club_id, runde, gender, season)

    response = requests.get(NULIGA_URL + rangliste_url, params=params)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, 'html.parser')
    tabelle = soup.find(class_="result-set")
    rows = tabelle.find_all('tr')
    
    if (len(rows) < 2): raise exception.EmptyResult("Keine Spieler:Innen gefunden")

    spielerinnen = rows[1:]
    rangliste = []
    for spieler in spielerinnen:
        data = spieler.find_all('td')
        spieler_daten = { "rang": data[0].text, "mannschaft": data[1].text, "name": data[3].text.strip() }
        rangliste.append(spieler_daten)
    
    return rangliste

def get_club_id(club_site_html):
    print("Hole Club-ID")
    soup = BeautifulSoup(club_site_html, 'html.parser')
    content_row = soup.find(id='content-row2')
    first_table = content_row.find('table')
    if (not first_table):
        raise RuntimeError("Verein konnte nicht gefunden werden!")
    link = first_table.find_all('a', string="Spielbetrieb und Ergebnisse")
    if (link):
        link_url = link[0]['href']
        if (not link_url):
            raise RuntimeError("Kein Link gefunden!", link)
        parsed_url = urlparse(link_url)
        search_params = parse_qs(parsed_url.query)
        club_id = search_params['club']
        if (not club_id):
            raise RuntimeError(f"Club ID in URL {parsed_url} nicht gefunden!")

        return club_id[0]
    else:
        raise RuntimeError("Kein Spielbetrieb und Ergebnisse Link gefunden!")
