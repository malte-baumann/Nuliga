from flask import Flask, request, jsonify, send_file, abort, current_app
from flask_cors import CORS
import data
import excel
from datetime import datetime
import re
from io import BytesIO

app = Flask(__name__)

CORS(appresources={r"/*": {"origins": ["https://nuliga.vercel.app"]}})

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

@app.route('/club', methods=['GET'])
def get_club_data():
    search_term = request.args.get('search_term')
    id = 0
    try:
        club_site = data.search_club(search_term.lower())
        id = data.get_club_id(club_site)
    except RuntimeError:
        return abort(404)

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

    filename = f"spieltermine_{search_term.replace(' ', '-')}.xlsx"
    excel_data = excel.create_sheet(rangliste, alle_mannschaften, sorted_teams, spiele, runde, filename)

    output = BytesIO()
    excel_data.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/')
def index():
    return current_app.send_static_file('index.html')

@app.errorhandler(404)
def not_found(_error):
    return "Verein nicht gefunden", 404
