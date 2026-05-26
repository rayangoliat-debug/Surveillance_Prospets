import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import time
import os
import re
import hashlib
import json
import sys
import smtplib
import csv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ==================== CONFIGURATION EMAIL ====================
EMAIL_EXPEDITEUR = "rayangoliat@gmail.com"
EMAIL_DESTINATAIRE = "fatima1.goliat@gmail.com"
MOT_DE_PASSE = "goliatcont@iner202509"

# ==================== CONFIGURATION GOOGLE SHEETS ====================
SHEET_SOURCE_ID = "16r5SyKtSrC8_JG5atMkYPXM4QynQReFGdNjQWOKJQ4E"
SHEET_DEST_ID = "1XBidRt-lJX9zXD3ZCCWZ-A1xKiW9NVrM5sVRPM5wce4"
ONGLET_DEST = "Extraction"
ONGLET_DEPT = "DEPARTEMENTS"
FICHIER_SUIVI = "derniere_ligne.txt"

DELAI_ENTRE_ECRITURES = 0.5
MAX_TENTATIVES = 3

# ==================== ALERTE EMAIL ====================
def envoyer_alerte_email(ligne_arret, erreur=None):
    try:
        sujet = f"🔴 Alerte Prospects - Arrêt à la ligne {ligne_arret}"
        corps = f"""
📊 ALERTE PROGRAMME PROSPECTS

Le script s'est arrêté à la ligne: {ligne_arret}

Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

"""
        if erreur:
            corps += f"Erreur: {str(erreur)[:200]}\n"
        corps += "\nVérifier les logs sur GitHub Actions."
        
        msg = MIMEMultipart()
        msg['From'] = EMAIL_EXPEDITEUR
        msg['To'] = EMAIL_DESTINATAIRE
        msg['Subject'] = sujet
        msg.attach(MIMEText(corps, 'plain'))
        
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(EMAIL_EXPEDITEUR, MOT_DE_PASSE)
            server.send_message(msg)
        print(f"✅ Email envoyé: arrêt à la ligne {ligne_arret}", flush=True)
        return True
    except Exception as e:
        print(f"⚠️ Erreur envoi email: {e}", flush=True)
        return False

# ==================== EXPORT CSV LOCAL ====================
def sauvegarder_csv_local(data, filename="prospects_export.csv"):
    file_exists = os.path.isfile(filename)
    
    with open(filename, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        
        if not file_exists:
            writer.writerow(['Timestamp', 'Email', 'Statut', 'Téléphone', 'Nom', 
                           'Département', 'Produit', 'Date', 'Plateforme', 'Source'])
        
        writer.writerows(data)
    
    print(f"   💾 {len(data)} prospects ajoutés à {filename}", flush=True)

# ==================== SUIVI DANS GOOGLE SHEETS ====================
def ecrire_suivi_dans_sheet(client, progression, total_nouvelles):
    try:
        spreadsheet = client.open_by_key(SHEET_DEST_ID)
        
        try:
            ws = spreadsheet.worksheet("Suivi")
        except:
            ws = spreadsheet.add_worksheet("Suivi", rows=100, cols=10)
            ws.append_row(["Timestamp", "Total", "Détail", "Fichier CSV"])
        
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        filename = f"prospects_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        
        details = "; ".join([f"{o}: {i.get('nouvelles', 0)}" for o, i in progression.items() if i.get('nouvelles', 0) > 0])
        
        ws.append_row([timestamp, total_nouvelles, details, filename])
        print(f"   📝 Progression écrite dans Google Sheet (onglet Suivi)", flush=True)
        return True
    except Exception as e:
        print(f"   ⚠️ Erreur écriture suivi: {e}", flush=True)
        return False

# ==================== GESTION DERNIÈRE LIGNE ====================
def generer_resume_import(progression, total_nouvelles, filename):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    resume = []
    resume.append(f"\n{'='*60}")
    resume.append(f"📊 IMPORT - {timestamp}")
    resume.append(f"📁 Fichier CSV: {filename}")
    resume.append(f"📈 TOTAL: {total_nouvelles} nouveaux prospects")
    resume.append("📋 Détail par onglet:")
    
    for onglet, info in progression.items():
        nouvelles = info.get("nouvelles", 0)
        if nouvelles > 0:
            resume.append(f"   - {onglet}: +{nouvelles} lignes")
        else:
            resume.append(f"   - {onglet}: 0 nouvelle ligne")
    
    resume.append(f"{'='*60}")
    return "\n".join(resume)

def sauvegarder_progression(progression, total_nouvelles=0):
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"prospects_export_{timestamp}.csv"
    
    resume = generer_resume_import(progression, total_nouvelles, filename)
    with open(FICHIER_SUIVI, 'a') as f:
        f.write(resume)
    
    progression_simple = {}
    for onglet, info in progression.items():
        progression_simple[onglet] = {
            "ligne_traitee": info.get("ligne_traitee", 0),
            "total": info.get("total", 0)
        }
    
    with open(FICHIER_SUIVI + ".json", 'w') as f:
        json.dump(progression_simple, f, indent=2)
    
    print(f"\n📝 Progression sauvegardée: {total_nouvelles} nouvelles lignes", flush=True)
    return filename

def charger_progression():
    if os.path.exists(FICHIER_SUIVI + ".json"):
        with open(FICHIER_SUIVI + ".json", 'r') as f:
            return json.load(f)
    return {}

def afficher_progression(progression):
    print("\n📊 PROGRESSION PAR ONGLET:")
    for onglet, info in progression.items():
        ligne = info.get("ligne_traitee", 0)
        total = info.get("total", 0)
        pourcent = round(ligne / max(total, 1) * 100, 1)
        print(f"   📄 {onglet}: {ligne}/{total} ({pourcent}%)")
    print()

# ==================== CONNEXION ====================
def connecter():
    print("🔐 Connexion à Google Sheets...", flush=True)
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_json:
        raise Exception("❌ Variable GOOGLE_CREDENTIALS non définie")
    
    creds_dict = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    print("✅ Connexion établie !", flush=True)
    return gspread.authorize(creds)

# ==================== NETTOYAGE ====================
def clean_phone(phone):
    if not phone or phone == "":
        return ""
    phone = str(phone).replace("p:", "").replace("p", "").strip()
    phone = re.sub(r'[^0-9+]', '', phone)
    if phone.startswith(("+33", "33", "+32", "32")):
        return phone
    if phone.startswith("0"):
        return phone
    return "0" + phone

def clean_postal(postal):
    if not postal or postal == "":
        return ""
    return str(postal).replace("z:", "").replace("z", "").strip()

def clean_product(product):
    if not product or product == "":
        return ""
    resultat = str(product)
    resultat = resultat.replace("GOLIAT", "").replace("goliat", "")
    resultat = re.sub(r'\s+', ' ', resultat).strip()
    return resultat

def traiter_opportunite_ct(valeur):
    """
    Traite la colonne Opportunité (CT) :
    - Retire tous les tirets '-'
    - Retire le 's' à la fin du mot "container" ou "containers"
    - Met le 'F' de 'Frigorifique' en minuscule (frigorifique)
    """
    if not valeur or valeur == "":
        return ""
    
    valeur = str(valeur)
    # Retire tous les tirets
    valeur = valeur.replace('-', '')
    
    # Retire le 's' à la fin du mot "container" ou "containers"
    # gère "container" -> "container", "containers" -> "container"
    valeur = re.sub(r'containers?\b', 'container', valeur, flags=re.IGNORECASE)
    
    # Met 'Frigorifique' avec un f minuscule
    valeur = re.sub(r'frigorifique', 'frigorifique', valeur, flags=re.IGNORECASE)
    
    return valeur

def extract_date(date_str):
    if not date_str or date_str == "":
        return ""
    try:
        return str(date_str)[:10]
    except:
        return ""

def get_etiquette(platform):
    if not platform:
        return "AUTRE"
    platform = str(platform).upper().strip()
    if platform in ["FB", "FACEBOOK", "META"]:
        return "FACEBOOK"
    if platform in ["IG", "INSTAGRAM"]:
        return "INSTAGRAM"
    return "AUTRE"

def get_etape(etiquette):
    return "prospect" if etiquette in ["FACEBOOK", "INSTAGRAM"] else ""

def get_source_id(etiquette):
    return "Meta" if etiquette in ["FACEBOOK", "INSTAGRAM"] else ""

# ==================== DÉPARTEMENT ====================
def get_departement_defaut(code_postal):
    if not code_postal or code_postal == "":
        return "Inconnu"
    code_postal = str(code_postal).strip()
    code_2 = code_postal[:2]
    code_3 = code_postal[:3]
    
    if code_3 == "971":
        return "971-Guadeloupe"
    elif code_3 == "972":
        return "972-Martinique"
    elif code_3 == "973":
        return "973-Guyane"
    elif code_3 == "974":
        return "974-La Réunion"
    elif code_3 == "976":
        return "976-Mayotte"
    
    if code_postal[:2] == "20" and len(code_postal) >= 3:
        if code_postal[2].isalpha():
            if code_postal[2] == 'A' or code_postal[2] == 'a':
                return "2A-Corse-du-Sud"
            elif code_postal[2] == 'B' or code_postal[2] == 'b':
                return "2B-Haute-Corse"
    
    try:
        code_int = int(code_2)
        if 1 <= code_int <= 95:
            dept_france = {
                "01": "01-Ain", "02": "02-Aisne", "03": "03-Allier", "04": "04-Alpes-de-Haute-Provence",
                "05": "05-Hautes-Alpes", "06": "06-Alpes-Maritimes", "07": "07-Ardèche", "08": "08-Ardennes",
                "09": "09-Ariège", "10": "10-Aube", "11": "11-Aude", "12": "12-Aveyron", "13": "13-Bouches-du-Rhône",
                "14": "14-Calvados", "15": "15-Cantal", "16": "16-Charente", "17": "17-Charente-Maritime",
                "18": "18-Cher", "19": "19-Corrèze", "21": "21-Côte-d'Or", "22": "22-Côtes-d'Armor",
                "23": "23-Creuse", "24": "24-Dordogne", "25": "25-Doubs", "26": "26-Drôme", "27": "27-Eure",
                "28": "28-Eure-et-Loir", "29": "29-Finistère", "30": "30-Gard", "31": "31-Haute-Garonne",
                "32": "32-Gers", "33": "33-Gironde", "34": "34-Hérault", "35": "35-Ille-et-Vilaine",
                "36": "36-Indre", "37": "37-Indre-et-Loire", "38": "38-Isère", "39": "39-Jura", "40": "40-Landes",
                "41": "41-Loir-et-Cher", "42": "42-Loire", "43": "43-Haute-Loire", "44": "44-Loire-Atlantique",
                "45": "45-Loiret", "46": "46-Lot", "47": "47-Lot-et-Garonne", "48": "48-Lozère",
                "49": "49-Maine-et-Loire", "50": "50-Manche", "51": "51-Marne", "52": "52-Haute-Marne",
                "53": "53-Mayenne", "54": "54-Meurthe-et-Moselle", "55": "55-Meuse", "56": "56-Morbihan",
                "57": "57-Moselle", "58": "58-Nièvre", "59": "59-Nord", "60": "60-Oise", "61": "61-Orne",
                "62": "62-Pas-de-Calais", "63": "63-Puy-de-Dôme", "64": "64-Pyrénées-Atlantiques",
                "65": "65-Hautes-Pyrénées", "66": "66-Pyrénées-Orientales", "67": "67-Bas-Rhin",
                "68": "68-Haut-Rhin", "69": "69-Rhône", "70": "70-Haute-Saône", "71": "71-Saône-et-Loire",
                "72": "72-Sarthe", "73": "73-Savoie", "74": "74-Haute-Savoie", "75": "75-Paris",
                "76": "76-Seine-Maritime", "77": "77-Seine-et-Marne", "78": "78-Yvelines", "79": "79-Deux-Sèvres",
                "80": "80-Somme", "81": "81-Tarn", "82": "82-Tarn-et-Garonne", "83": "83-Var", "84": "84-Vaucluse",
                "85": "85-Vendée", "86": "86-Vienne", "87": "87-Haute-Vienne", "88": "88-Vosges", "89": "89-Yonne",
                "90": "90-Territoire-de-Belfort", "91": "91-Essonne", "92": "92-Hauts-de-Seine",
                "93": "93-Seine-Saint-Denis", "94": "94-Val-de-Marne", "95": "95-Val-d'Oise"
            }
            code_str = f"{code_int:02d}"
            if code_str in dept_france:
                return dept_france[code_str]
    except:
        pass
    
    try:
        code_int_belge = int(code_postal[:4]) if len(code_postal) >= 4 else 0
        if 1000 <= code_int_belge <= 9999:
            if code_int_belge >= 1000 and code_int_belge <= 1999:
                return "1000-Bruxelles"
            elif code_int_belge >= 2000 and code_int_belge <= 2999:
                return "2000-Anvers"
            elif code_int_belge >= 3000 and code_int_belge <= 3999:
                return "3000-Liège"
            elif code_int_belge >= 4000 and code_int_belge <= 4999:
                return "4000-Luxembourg"
            elif code_int_belge >= 5000 and code_int_belge <= 5999:
                return "5000-Namur"
            elif code_int_belge >= 6000 and code_int_belge <= 6999:
                return "6000-Hainaut"
            elif code_int_belge >= 7000 and code_int_belge <= 7999:
                return "7000-Limbourg"
            elif code_int_belge >= 8000 and code_int_belge <= 8999:
                return "8000-Flandre-Occidentale"
            elif code_int_belge >= 9000 and code_int_belge <= 9999:
                return "9000-Flandre-Orientale"
    except:
        pass
    
    return "Inconnu"

def charger_departements(client):
    try:
        sheet = client.open_by_key(SHEET_DEST_ID)
        ws_dept = sheet.worksheet(ONGLET_DEPT)
        data = ws_dept.get_all_values()
        dept_dict = {}
        for row in data[1:]:
            if len(row) >= 2:
                dept_dict[str(row[0]).strip()] = str(row[1]).strip()
        return dept_dict
    except Exception as e:
        print(f"Erreur chargement départements: {e}")
        return {}

def get_department(postal_code, dept_dict):
    if not postal_code or postal_code == "":
        return "Inconnu"
    code_postal = str(postal_code).strip()
    code = code_postal[:2]
    if len(code_postal) >= 3 and code_postal[:2] == "20":
        if code_postal[2].isalpha():
            code = code_postal[:3]
    departement = dept_dict.get(code)
    if departement:
        return f"{code}-{departement}"
    return get_departement_defaut(code_postal)

# ==================== COLORATION ====================
def colorer_ligne(ws_dest, row_number):
    try:
        range_a_colorer = f"A{row_number}:P{row_number}"
        format_vert = {
            "backgroundColor": {
                "red": 0.8,
                "green": 1.0,
                "blue": 0.8
            }
        }
        ws_dest.format(range_a_colorer, format_vert)
        print(f"   🎨 Ligne {row_number} colorée", flush=True)
        return True
    except Exception as e:
        print(f"   ⚠️ Erreur coloration ligne {row_number}: {e}", flush=True)
        return False

# ==================== TROUVER LES COLONNES ====================
def trouver_colonnes(entetes):
    colonnes = {}
    for i, col in enumerate(entetes):
        col_lower = col.lower().strip()
        if 'email' in col_lower or 'e-mail' in col_lower:
            colonnes['email'] = i
        elif 'created_time' in col_lower:
            colonnes['created_time'] = i
        elif 'adset_name' in col_lower:
            colonnes['adset_name'] = i
        elif 'platform' in col_lower:
            colonnes['platform'] = i
        elif 'statut' in col_lower:
            colonnes['statut'] = i
        elif 'nom_complet' in col_lower or 'full_name' in col_lower:
            colonnes['nom_complet'] = i
        elif 'numéro_de_téléphone' in col_lower or 'telephone' in col_lower:
            colonnes['telephone'] = i
        elif 'code_postal' in col_lower or 'postal' in col_lower:
            colonnes['code_postal'] = i
    return colonnes

# ==================== LIRE TOUS LES ONGLETS ====================
def lire_onglets_avec_progression(sheet, progression_existante):
    toutes_les_lignes = []
    nouvelle_progression = {}
    premier = True
    
    for ws in sheet.worksheets():
        nom_onglet = ws.title
        print(f"   📄 Lecture de l'onglet: {nom_onglet}", flush=True)
        
        try:
            data = ws.get_all_values()
            if len(data) <= 1:
                nouvelle_progression[nom_onglet] = {"total": 0, "ligne_traitee": 0, "nouvelles": 0}
                continue
            
            nb_lignes = len(data) - 1
            ligne_depart = progression_existante.get(nom_onglet, {}).get("ligne_traitee", 0)
            
            print(f"      Total: {nb_lignes} lignes, Dernière traitée: {ligne_depart}", flush=True)
            
            if premier:
                toutes_les_lignes.append(data[0])
                if ligne_depart < nb_lignes:
                    nouvelles_donnees = data[ligne_depart + 1:]
                    toutes_les_lignes.extend(nouvelles_donnees)
                    ligne_nouvelle = nb_lignes
                else:
                    nouvelles_donnees = []
                    ligne_nouvelle = ligne_depart
                premier = False
            else:
                if ligne_depart < nb_lignes:
                    nouvelles_donnees = data[ligne_depart + 1:]
                    toutes_les_lignes.extend(nouvelles_donnees)
                    ligne_nouvelle = nb_lignes
                else:
                    nouvelles_donnees = []
                    ligne_nouvelle = ligne_depart
            
            nouvelle_progression[nom_onglet] = {
                "total": nb_lignes,
                "ligne_traitee": ligne_nouvelle,
                "nouvelles": len(nouvelles_donnees)
            }
            
            print(f"      ✅ Nouvelles lignes: {len(nouvelles_donnees)}", flush=True)
            
        except Exception as e:
            print(f"      ⚠️ Erreur lecture {nom_onglet}: {e}", flush=True)
            nouvelle_progression[nom_onglet] = {"total": 0, "ligne_traitee": 0, "nouvelles": 0}
    
    return toutes_les_lignes, nouvelle_progression

# ==================== VIDER L'EXTRACTION ====================
def vider_extraction(ws_dest):
    try:
        toutes_les_lignes = ws_dest.get_all_values()
        nb_lignes = len(toutes_les_lignes)
        
        if nb_lignes > 1:
            ws_dest.delete_rows(2, nb_lignes - 1)
            print(f"   🗑️ Onglet Extraction vidé ({nb_lignes - 1} lignes supprimées)", flush=True)
        else:
            print(f"   ✅ Onglet Extraction déjà vide", flush=True)
    except Exception as e:
        print(f"   ⚠️ Erreur vidage extraction: {e}", flush=True)

# ==================== AJOUTER ====================
def ajouter_prospect_avec_retry(ws_dest, ligne, tentative=0):
    try:
        ws_dest.append_row(ligne, value_input_option='USER_ENTERED')
        return True
    except Exception as e:
        if "429" in str(e) and tentative < MAX_TENTATIVES:
            time.sleep(5)
            return ajouter_prospect_avec_retry(ws_dest, ligne, tentative + 1)
        return False

# ==================== TRAITEMENT PRINCIPAL ====================
def traiter():
    print(f"\n{'='*50}", flush=True)
    print(f"{datetime.now().strftime('%H:%M:%S')} - Vérification...", flush=True)
    
    try:
        client = connecter()
        
        progression_existante = charger_progression()
        if progression_existante:
            print("📊 Progression chargée:", flush=True)
            for onglet, info in progression_existante.items():
                print(f"   📄 {onglet}: {info.get('ligne_traitee', 0)}/{info.get('total', 0)}")
        
        print("\n📂 Lecture des onglets du fichier source...", flush=True)
        sheet_source = client.open_by_key(SHEET_SOURCE_ID)
        toutes_les_lignes, nouvelle_progression = lire_onglets_avec_progression(sheet_source, progression_existante)
        
        if len(toutes_les_lignes) <= 1:
            print("Aucune donnée dans la source", flush=True)
            return
        
        afficher_progression(nouvelle_progression)
        
        total_nouvelles = sum(info.get("nouvelles", 0) for info in nouvelle_progression.values())
        
        if total_nouvelles == 0:
            print("Aucun nouveau prospect à ajouter", flush=True)
            return
        
        entetes = toutes_les_lignes[0]
        colonnes = trouver_colonnes(entetes)
        print("Colonnes trouvées:", list(colonnes.keys()), flush=True)
        
        if 'email' not in colonnes:
            print("❌ Colonne email non trouvée !", flush=True)
            return
        
        print("Chargement des départements...", flush=True)
        dept_dict = charger_departements(client)
        print(f"   {len(dept_dict)} départements chargés", flush=True)
        
        sheet_dest = client.open_by_key(SHEET_DEST_ID)
        ws_dest = sheet_dest.worksheet(ONGLET_DEST)
        
        print("🧹 Nettoyage de l'onglet Extraction...", flush=True)
        vider_extraction(ws_dest)
        
        nouvelles_lignes = toutes_les_lignes[1:]
        print(f"Nouveaux prospects à traiter: {len(nouvelles_lignes)}", flush=True)
        
        # 🔥 STOCKER LES PROSPECTS AVANT TRI (SANS DÉDOUBLONNAGE)
        prospects_a_ajouter = []
        
        for ligne in nouvelles_lignes:
            if len(ligne) <= max(colonnes.values()):
                continue
            
            email = ligne[colonnes['email']].strip()
            if not email or email == "":
                continue
            if "test" in email.lower() or "dummy" in email.lower():
                continue
            
            statut = ligne[colonnes['statut']].strip() if colonnes.get('statut') is not None else ""
            telephone_brut = ligne[colonnes['telephone']] if colonnes.get('telephone') is not None else ""
            nom = ligne[colonnes['nom_complet']].strip() if colonnes.get('nom_complet') is not None else ""
            postal_brut = ligne[colonnes['code_postal']] if colonnes.get('code_postal') is not None else ""
            produit_brut = ligne[colonnes['adset_name']] if colonnes.get('adset_name') is not None else ""
            date_brute = ligne[colonnes['created_time']] if colonnes.get('created_time') is not None else ""
            platform_brut = ligne[colonnes['platform']] if colonnes.get('platform') is not None else ""
            
            telephone = clean_phone(telephone_brut)
            code_postal = clean_postal(postal_brut)
            produit = clean_product(produit_brut)
            # Application du traitement sur la colonne Opportunité (CT)
            produit = traiter_opportunite_ct(produit)
            date_clean = extract_date(date_brute)
            etiquette = get_etiquette(platform_brut)
            departement = get_department(code_postal, dept_dict)
            etape = get_etape(etiquette)
            source_id = get_source_id(etiquette)
            
            nouvelle_ligne = [
                email, statut, telephone, nom, departement,
                nom, produit, "", "", date_clean,
                etiquette, "", "", "", etape, source_id
            ]
            
            # Stocker pour tri
            prospects_a_ajouter.append({
                "ligne": nouvelle_ligne,
                "email": email,
                "statut": statut,
                "telephone": telephone,
                "nom": nom,
                "departement": departement,
                "produit": produit,
                "date": date_clean,
                "etiquette": etiquette
            })
        
        # 🔥 TRIER PAR DATE (du plus ancien au plus récent)
        prospects_a_ajouter.sort(key=lambda x: x["date"])
        
        # 🔥 AJOUTER LES PROSPECTS TRIÉS
        count = 0
        export_data = []
        
        for prospect in prospects_a_ajouter:
            if ajouter_prospect_avec_retry(ws_dest, prospect["ligne"]):
                count += 1
                export_data.append([
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    prospect["email"], prospect["statut"], prospect["telephone"], prospect["nom"],
                    prospect["departement"], prospect["produit"], prospect["date"], prospect["etiquette"], "Meta"
                ])
                if count % 10 == 0:
                    print(f"   📝 {count} prospects ajoutés...", flush=True)
            
            time.sleep(DELAI_ENTRE_ECRITURES)
        
        if count > 0:
            premiere_ligne = 2
            print(f"   🎨 Coloration des lignes {premiere_ligne} à {premiere_ligne + count - 1}", flush=True)
            
            for i in range(count):
                colorer_ligne(ws_dest, premiere_ligne + i)
            
            filename = sauvegarder_progression(nouvelle_progression, total_nouvelles)
            ecrire_suivi_dans_sheet(client, nouvelle_progression, total_nouvelles)
            
            print(f"\n✅ {count} prospects ajoutés avec succès !", flush=True)
            print(f"🎨 {count} nouvelles lignes colorées", flush=True)
            
            date_str = datetime.now().strftime('%Y%m%d_%H%M%S')
            csv_filename = f"prospects_export_{date_str}.csv"
            sauvegarder_csv_local(export_data, csv_filename)
            
        else:
            print("Aucun prospect valide à ajouter", flush=True)
            
    except Exception as e:
        print(f"❌ Erreur: {e}", flush=True)
        import traceback
        traceback.print_exc()
        return False

# ==================== LANCER ====================
if __name__ == "__main__":
    print("🚀 Démarrage du script (version avec traitement Opportunité CT)", flush=True)
    traiter()
    print("✅ Script terminé", flush=True)
