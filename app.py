import streamlit as st
import sqlite3
import hashlib
import hmac
import secrets

from datetime import date, datetime, timedelta
from io import BytesIO
from xml.sax.saxutils import escape

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
)


# ------------------------------------------------------------
# Configuration
# ------------------------------------------------------------

st.set_page_config(
    page_title="Planning des vacations – SEM / DRMCO",
    page_icon="📅",
    layout="wide",
)

DB_PATH = "planning_sem.db"

DAYS = [
    "Lundi", "Mardi", "Mercredi", "Jeudi",
    "Vendredi", "Samedi", "Dimanche"
]

REGULAR_ROLES = ["Chef de service", "Contrôleur"]

# ------------------------------------------------------------
# Database
# ------------------------------------------------------------

def db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def hash_password(password: str, salt: bytes | None = None):
    salt = salt or secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        200_000
    )
    return salt.hex(), digest.hex()


def verify_password(password: str, salt_hex: str, hash_hex: str):
    salt = bytes.fromhex(salt_hex)
    _, candidate = hash_password(password, salt)
    return hmac.compare_digest(candidate, hash_hex)


def init_db():
    with db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                full_name TEXT NOT NULL,
                role TEXT NOT NULL,
                salt TEXT NOT NULL,
                password_hash TEXT NOT NULL,
                active INTEGER NOT NULL DEFAULT 1
            )
        """)

        conn.execute("""
            CREATE TABLE IF NOT EXISTS vacations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                vacation_date TEXT NOT NULL,
                note TEXT DEFAULT '',
                updated_at TEXT NOT NULL,
                UNIQUE(user_id, vacation_date),
                FOREIGN KEY(user_id) REFERENCES users(id)
            )
        """)

        conn.execute("""
            CREATE TABLE IF NOT EXISTS app_settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
        """)

        existing = conn.execute(
            "SELECT id FROM users WHERE username = ?",
            ("soukaina",)
        ).fetchone()

        if not existing:
            salt, pw_hash = hash_password("ChangeMe123!")
            conn.execute(
                """
                INSERT INTO users
                (username, full_name, role, salt, password_hash)
                VALUES (?, ?, ?, ?, ?)
                """,
                (
                    "soukaina",
                    "Boughanem Soukaina",
                    "Chef de service",
                    salt,
                    pw_hash
                ),
            )


def get_user(username):
    with db() as conn:
        return conn.execute(
            """
            SELECT * FROM users
            WHERE username=? AND active=1
            """,
            (username,)
        ).fetchone()


def list_users():
    with db() as conn:
        return conn.execute(
            """
            SELECT id, username, full_name, role, active
            FROM users
            ORDER BY id
            """
        ).fetchall()


def create_user(username, full_name, role, password):
    salt, pw_hash = hash_password(password)
    with db() as conn:
        conn.execute(
            """
            INSERT INTO users
            (username, full_name, role, salt, password_hash)
            VALUES (?, ?, ?, ?, ?)
            """,
            (username.strip(), full_name.strip(), role, salt, pw_hash),
        )

def update_user_role(user_id, new_role):
    with db() as conn:
        conn.execute(
            "UPDATE users SET role=? WHERE id=?",
            (new_role, user_id)
        )
        

def change_password(user_id, new_password):
    salt, pw_hash = hash_password(new_password)
    with db() as conn:
        conn.execute(
            """
            UPDATE users
            SET salt=?, password_hash=?
            WHERE id=?
            """,
            (salt, pw_hash, user_id),
        )


def set_user_active(user_id, active):
    """Active ou désactive un compte sans supprimer son historique."""
    with db() as conn:
        conn.execute(
            "UPDATE users SET active=? WHERE id=?",
            (1 if active else 0, user_id),
        )


def get_vacations(start_date, end_date):
    with db() as conn:
        return conn.execute(
            """
            SELECT v.id, v.user_id, v.vacation_date,
                   v.note, v.updated_at,
                   u.full_name, u.role
            FROM vacations v
            JOIN users u ON u.id=v.user_id
            WHERE v.vacation_date BETWEEN ? AND ?
            ORDER BY v.vacation_date, u.full_name
            """,
            (start_date.isoformat(), end_date.isoformat()),
        ).fetchall()


def save_vacation(user_id, vacation_date, note):
    with db() as conn:
        existing = conn.execute(
            """
            SELECT id FROM vacations
            WHERE user_id=? AND vacation_date=?
            """,
            (user_id, vacation_date.isoformat()),
        ).fetchone()

        count = conn.execute(
            """
            SELECT COUNT(*) AS n
            FROM vacations
            WHERE vacation_date=?
            """,
            (vacation_date.isoformat(),),
        ).fetchone()["n"]

        if existing:
            conn.execute(
                """
                UPDATE vacations
                SET note=?, updated_at=?
                WHERE user_id=? AND vacation_date=?
                """,
                (
                    note,
                    datetime.now().isoformat(timespec="seconds"),
                    user_id,
                    vacation_date.isoformat()
                ),
            )
            return True, "Vacation mise à jour."

        if count >= 2:
            return False, (
                "Les deux places de cette matinée sont déjà prises."
            )

        conn.execute(
            """
            INSERT INTO vacations
            (user_id, vacation_date, note, updated_at)
            VALUES (?, ?, ?, ?)
            """,
            (
                user_id,
                vacation_date.isoformat(),
                note,
                datetime.now().isoformat(timespec="seconds")
            ),
        )
        return True, "Vacation enregistrée."


def delete_vacation(vacation_id, user_id=None):
    with db() as conn:
        if user_id is None:
            conn.execute(
                "DELETE FROM vacations WHERE id=?",
                (vacation_id,)
            )
        else:
            conn.execute(
                """
                DELETE FROM vacations
                WHERE id=? AND user_id=?
                """,
                (vacation_id, user_id),
            )


def monday_of(d):
    return d - timedelta(days=d.weekday())


def five_week_period():
    start = monday_of(date.today())
    end = start + timedelta(days=34)
    return start, end


# ------------------------------------------------------------
# Export Excel : modèle vierge personnalisé
# Structure : Noms, Grade, Fonction, dates de la période,
# puis Observations. 9 lignes vierges, sans numérotation.
# ------------------------------------------------------------

def creer_modele_excel(date_debut=None, date_fin=None):

    if date_debut is None:
        date_debut = monday_of(date.today())
    if date_fin is None:
        date_fin = date_debut + timedelta(days=6)

    if date_fin < date_debut:
        raise ValueError("La date de fin doit être postérieure à la date de début.")

    dates = [
        date_debut + timedelta(days=i)
        for i in range((date_fin - date_debut).days + 1)
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "Tableau de service"

    bleu = "D9EAF7"
    gris = "E7E6E6"

    bord_fin = Side(style="thin", color="333333")
    bordures = Border(
        left=bord_fin,
        right=bord_fin,
        top=bord_fin,
        bottom=bord_fin
    )

    # 3 colonnes fixes + une colonne par date + Observations
    derniere_colonne = 4 + len(dates)  # dernière date = 3 + len(dates)
    derniere_lettre = get_column_letter(derniere_colonne)

    # Titre
    ws.merge_cells(f"A1:{derniere_lettre}1")
    ws["A1"] = "TABLEAU HEBDOMADAIRE DE SERVICE"
    ws["A1"].font = Font(
        name="Arial",
        size=15,
        bold=True,
        color="16324F"
    )
    ws["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )
    ws.row_dimensions[1].height = 30

    # Informations générales
    info_cells = [
        ("A2:B2", "PÉRIODE DU :"),
        ("C2:D2", date_debut.strftime("%d/%m/%Y")),
        ("E2:F2", "AU :"),
        ("G2:H2", date_fin.strftime("%d/%m/%Y")),
    ]

    for cell_range, value in info_cells:
        ws.merge_cells(cell_range)
        cell = ws[cell_range.split(":")[0]]
        cell.value = value
        cell.font = Font(size=9, bold=True)
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    # Informations facultatives complémentaires
    ws.merge_cells("I2:J2")
    ws["I2"] = "EXEMPLAIRE DESTINÉ À :"
    ws["I2"].font = Font(size=8, bold=True)
    ws["I2"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.merge_cells("K2:L2")
    ws["K2"] = "TRANSMIS LE :"
    ws["K2"].font = Font(size=8, bold=True)
    ws["K2"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.merge_cells("M2:N2")
    ws["M2"] = "VISA :"
    ws["M2"].font = Font(size=8, bold=True)
    ws["M2"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[2].height = 28

    # En-têtes fixes
    fixed_headers = [
        ("A", "NOMS ET PRÉNOMS", 23),
        ("B", "GRADE", 10),
        ("C", "FONCTION", 14),
    ]

    for col_letter, title, width in fixed_headers:
        ws.merge_cells(f"{col_letter}4:{col_letter}5")
        cell = ws[f"{col_letter}4"]
        cell.value = title
        cell.font = Font(bold=True, size=9)
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )
        cell.fill = PatternFill("solid", fgColor=gris)
        ws.column_dimensions[col_letter].width = width

    # Une colonne par date : aucun numéro ni sous-colonne horaire
    for index, current_date in enumerate(dates):
        col = 4 + index
        cell = ws.cell(row=4, column=col)
        day_name = DAYS[current_date.weekday()]
        cell.value = f"{day_name.upper()}\n{current_date.strftime('%d/%m/%Y')}"
        cell.font = Font(bold=True, size=8)
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )
        cell.fill = PatternFill("solid", fgColor=bleu)
        ws.merge_cells(
            start_row=4,
            start_column=col,
            end_row=5,
            end_column=col
        )
        ws.column_dimensions[get_column_letter(col)].width = 13

    # Observations immédiatement après le dernier jour
    col_obs = 4 + len(dates)
    lettre_obs = get_column_letter(col_obs)
    ws.merge_cells(
        start_row=4,
        start_column=col_obs,
        end_row=5,
        end_column=col_obs
    )
    cell_obs = ws.cell(row=4, column=col_obs)
    cell_obs.value = "OBSERVATIONS"
    cell_obs.font = Font(bold=True, size=9)
    cell_obs.alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )
    cell_obs.fill = PatternFill("solid", fgColor=gris)
    ws.column_dimensions[lettre_obs].width = 28

    # 9 lignes vierges
    premiere_ligne = 6
    derniere_ligne = 14

    for ligne in range(premiere_ligne, derniere_ligne + 1):
        for col in range(1, derniere_colonne + 1):
            cell = ws.cell(row=ligne, column=col)
            cell.border = bordures
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )
        ws.row_dimensions[ligne].height = 38

    # Bordures des en-têtes
    for ligne in range(4, 6):
        for col in range(1, derniere_colonne + 1):
            ws.cell(row=ligne, column=col).border = bordures

    # Mise en page
    ws.freeze_panes = "D6"
    ws.sheet_view.showGridLines = False

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_margins = PageMargins(
        left=0.2,
        right=0.2,
        top=0.3,
        bottom=0.3
    )

    ws.print_area = f"A1:{derniere_lettre}{derniere_ligne}"

    fichier = BytesIO()
    wb.save(fichier)
    fichier.seek(0)
    return fichier


# ------------------------------------------------------------
# Export PDF : modèle vierge personnalisé
# ------------------------------------------------------------

def creer_modele_pdf(date_debut=None, date_fin=None):

    if date_debut is None:
        date_debut = monday_of(date.today())
    if date_fin is None:
        date_fin = date_debut + timedelta(days=6)

    if date_fin < date_debut:
        raise ValueError("La date de fin doit être postérieure à la date de début.")

    dates = [
        date_debut + timedelta(days=i)
        for i in range((date_fin - date_debut).days + 1)
    ]

    elements = []

    style_titre = ParagraphStyle(
        "Titre",
        fontName="Helvetica-Bold",
        fontSize=13,
        alignment=TA_CENTER,
        spaceAfter=8
    )

    style_entete = ParagraphStyle(
        "Entete",
        fontName="Helvetica-Bold",
        fontSize=6,
        alignment=TA_CENTER,
        leading=7
    )

    style_normal = ParagraphStyle(
        "Normal",
        fontName="Helvetica-Bold",
        fontSize=7,
        alignment=TA_CENTER
    )

    elements.append(
        Paragraph("TABLEAU HEBDOMADAIRE DE SERVICE", style_titre)
    )

    infos = (
        f"PÉRIODE DU : {date_debut.strftime('%d/%m/%Y')}"
        f" &nbsp;&nbsp;&nbsp; AU : {date_fin.strftime('%d/%m/%Y')}"
        " &nbsp;&nbsp;&nbsp; EXEMPLAIRE DESTINÉ À : __________"
        " &nbsp;&nbsp;&nbsp; TRANSMIS LE : __________"
        " &nbsp;&nbsp;&nbsp; VISA : __________"
    )
    elements.append(Paragraph(infos, style_normal))
    elements.append(Spacer(1, 10))

    # Colonnes : Noms, Grade, Fonction, dates, Observations
    nombre_colonnes = 4 + len(dates)
    ligne_entetes = [""] * nombre_colonnes

    ligne_entetes[0] = Paragraph("NOMS ET<br/>PRÉNOMS", style_entete)
    ligne_entetes[1] = Paragraph("GRADE", style_entete)
    ligne_entetes[2] = Paragraph("FONCTION", style_entete)

    for index, current_date in enumerate(dates):
        day_name = DAYS[current_date.weekday()]
        ligne_entetes[3 + index] = Paragraph(
            f"{day_name.upper()}<br/>{current_date.strftime('%d/%m/%Y')}",
            style_entete
        )

    ligne_entetes[-1] = Paragraph("OBSERVATIONS", style_entete)

    donnees = [ligne_entetes]

    # Exactement 9 lignes vierges
    for _ in range(9):
        donnees.append([""] * nombre_colonnes)

    # Largeurs proportionnelles, adaptées à la largeur disponible en A3
    largeur_disponible = landscape(A3)[0] - 36
    largeur_fixes = [88, 38, 52]
    largeur_observations = 105
    largeur_jours_totale = max(
        120,
        largeur_disponible - sum(largeur_fixes) - largeur_observations
    )
    largeur_par_jour = largeur_jours_totale / len(dates)

    largeurs = (
        largeur_fixes
        + [largeur_par_jour] * len(dates)
        + [largeur_observations]
    )

    tableau = Table(
        donnees,
        colWidths=largeurs,
        rowHeights=[38] + [38] * 9,
        repeatRows=1
    )

    tableau.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (2, 0), colors.lightgrey),
        ("BACKGROUND", (3, 0), (-2, 0), colors.HexColor("#D9EAF7")),
        ("BACKGROUND", (-1, 0), (-1, 0), colors.lightgrey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 6),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("BOX", (0, 0), (-1, -1), 1, colors.black),
    ]))

    elements.append(tableau)

    fichier = BytesIO()

    document = SimpleDocTemplate(
        fichier,
        pagesize=landscape(A3),
        rightMargin=18,
        leftMargin=18,
        topMargin=18,
        bottomMargin=18
    )

    document.build(elements)
    fichier.seek(0)
    return fichier


# ------------------------------------------------------------
# Export PDF : planning final rempli depuis la base de données
# ------------------------------------------------------------

def trier_equipe(users):
    """Place les fonctions régulières en premier, puis les prévisionnistes."""
    ordre = {
        "Chef de service": 0,
        "Contrôleur": 1,
        "Prévisionniste": 2,
    }
    return sorted(
        [u for u in users if u["active"]],
        key=lambda u: (ordre.get(u["role"], 3), u["full_name"].casefold())
    )


def creer_planning_final_pdf(date_debut, date_fin, users, vacations):
    """Construit le planning réel à partir des comptes et vacations enregistrés."""

    if date_fin < date_debut:
        raise ValueError("La date de fin doit être postérieure à la date de début.")

    dates = [
        date_debut + timedelta(days=i)
        for i in range((date_fin - date_debut).days + 1)
    ]

    # Ne conserver que les utilisateurs actifs et appliquer l'ordre du tableau.
    equipe = trier_equipe(users)

    if len(equipe) > 9:
        raise ValueError(
            f"Le tableau comporte 9 lignes disponibles, mais {len(equipe)} "
            "comptes sont actifs. Désactivez les comptes inutilisés avant "
            "de générer le PDF."
        )

    # Indexer les vacations par utilisateur et par date.
    vacations_par_user = {}
    for vac in vacations:
        vacations_par_user.setdefault(vac["user_id"], {})[
            vac["vacation_date"]
        ] = vac

    style_titre = ParagraphStyle(
        "TitrePlanningFinal", fontName="Helvetica-Bold",
        fontSize=16, alignment=TA_CENTER, spaceAfter=7
    )
    style_info = ParagraphStyle(
        "InfoPlanningFinal", fontName="Helvetica",
        fontSize=10, alignment=TA_LEFT, leading=9
    )
    style_entete = ParagraphStyle(
        "EntetePlanningFinal", fontName="Helvetica-Bold",
        fontSize=8, alignment=TA_CENTER, leading=7
    )
    style_cellule = ParagraphStyle(
        "CellulePlanningFinal", fontName="Helvetica-Bold",
        fontSize=8, alignment=TA_CENTER, leading=7
    )
    style_obs = ParagraphStyle(
        "ObsPlanningFinal", fontName="Helvetica",
        fontSize=8, alignment=TA_LEFT, leading=6.5
    )

    elements = [
        Paragraph("TABLEAU DE SERVICE – PLANNING FINAL<br/>SEM – DRMCO", style_titre),
        Paragraph(
            f"<b>Période :</b> {date_debut.strftime('%d/%m/%Y')} "
            f"au {date_fin.strftime('%d/%m/%Y')} "
            f"&nbsp;&nbsp;&nbsp; <b>Édité le :</b> "
            f"{date.today().strftime('%d/%m/%Y')}",
            style_info
        ),
        Spacer(1, 8)
    ]

    entetes = [
        Paragraph("NOMS ET PRÉNOMS", style_entete),
        Paragraph("GRADE", style_entete),
        Paragraph("FONCTION", style_entete),
    ]
    for d in dates:
        entetes.append(
            Paragraph(
                f"{DAYS[d.weekday()].upper()}<br/>{d.strftime('%d/%m')}",
                style_entete
            )
        )
    entetes.append(Paragraph("OBSERVATIONS", style_entete))

    donnees = [entetes]

    for user in equipe:
        nom = escape(str(user["full_name"] or ""))
        role = escape(str(user["role"] or ""))
        # Le champ grade n'existe pas dans la base actuelle.
        ligne = [
            Paragraph(nom, style_cellule),
            Paragraph("", style_cellule),
            Paragraph(role, style_cellule),
        ]

        observations = []
        user_vacations = vacations_par_user.get(user["id"], {})

        for d in dates:
            key = d.isoformat()

            if user["role"] in REGULAR_ROLES:
                # Les fonctions régulières sont prévues du lundi au vendredi.
                contenu = "Régulier" if d.weekday() < 5 else ""
            elif key in user_vacations:
                contenu = "Matin"
                note = (user_vacations[key]["note"] or "").strip()
                if note:
                    observations.append(
                        f"{d.strftime('%d/%m')}: {note}"
                    )
            else:
                contenu = ""

            ligne.append(Paragraph(contenu, style_cellule))

        obs = "<br/>".join(escape(x) for x in observations) or " "
        ligne.append(Paragraph(obs, style_obs))
        donnees.append(ligne)

    # Le modèle final comporte exactement neuf lignes de personnel.
    while len(donnees) < 10:
        donnees.append([""] * len(entetes))

    largeur_page = landscape(A3)[0] - 36
    largeur_nom, largeur_grade, largeur_fonction, largeur_obs = 105, 35, 65, 125
    largeur_jours = (
        largeur_page - largeur_nom - largeur_grade
        - largeur_fonction - largeur_obs
    ) / len(dates)

    largeurs = (
        [largeur_nom, largeur_grade, largeur_fonction]
        + [largeur_jours] * len(dates)
        + [largeur_obs]
    )

    hauteurs = [67] + [66] * (len(donnees) - 1)
    tableau = Table(
        donnees,
        colWidths=largeurs,
        rowHeights=hauteurs,
        repeatRows=1
    )
    tableau.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (2, 0), colors.lightgrey),
        ("BACKGROUND", (3, 0), (-2, 0), colors.HexColor("#D9EAF7")),
        ("BACKGROUND", (-1, 0), (-1, 0), colors.lightgrey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    elements.append(tableau)
    # --------------------------------------------------------
    # Signature du Chef de service en bas à droite
    # --------------------------------------------------------

    elements.append(Spacer(1, 3))

    style_signature = ParagraphStyle(
        "Signature",
        fontName="Helvetica-Bold",
        fontSize=8,
        alignment=TA_CENTER,
        leading=10
    )

    signature = Table(
        [[Paragraph(
            "<br/><br/><br/>Signature : ",
            style_signature
        )]],
        colWidths=[180],
        rowHeights=[35]
    )

    signature.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))

    bloc_signature = Table(
        [["", signature]],
        colWidths=[largeur_page - 180, 180]
    )

    bloc_signature.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))

    elements.append(bloc_signature)

    fichier = BytesIO()

    document = SimpleDocTemplate(
        fichier,
        pagesize=landscape(A3),
        rightMargin=18, leftMargin=18,
        topMargin=12, bottomMargin=5
    )
    document.build(elements)
    fichier.seek(0)
    return fichier


# ------------------------------------------------------------
# Initialisation
# ------------------------------------------------------------

init_db()

# ------------------------------------------------------------
# Styling
# ------------------------------------------------------------

st.markdown("""
<style>
    .main-title {
        color: #16324F;
        font-weight: 750;
    }

    div[data-testid="stMetric"] {
        background: #F1F5F9;
        padding: 12px;
        border-radius: 10px;
    }
</style>
""", unsafe_allow_html=True)

st.title("📅 Planning des vacations")
st.caption("Service Exploitation Météorologique – DRMCO")

# ------------------------------------------------------------
# Login
# ------------------------------------------------------------

if "user" not in st.session_state:
    st.session_state.user = None

if st.session_state.user is None:

    st.subheader("Connexion")

    with st.form("login_form"):
        username = st.text_input("Identifiant")
        password = st.text_input("Mot de passe", type="password")

        submitted = st.form_submit_button(
            "Se connecter",
            use_container_width=True
        )

    if submitted:
        user = get_user(username.strip())

        if user and verify_password(
            password,
            user["salt"],
            user["password_hash"]
        ):
            st.session_state.user = {
                "id": user["id"],
                "username": user["username"],
                "full_name": user["full_name"],
                "role": user["role"],
            }
            st.rerun()
        else:
            st.error("Identifiant ou mot de passe incorrect.")

    st.stop()


current = st.session_state.user
is_admin = current["role"] == "Chef de service"


# ------------------------------------------------------------
# Sidebar
# ------------------------------------------------------------

with st.sidebar:
    st.write(f"**{current['full_name']}**")
    st.caption(current["role"])

    if st.button("Se déconnecter", use_container_width=True):
        st.session_state.user = None
        st.rerun()

    pages = [
        "Planning des 5 semaines",
        "Saisir mes vacations",
        "Mon compte"
    ]

    if is_admin:
        pages.append("Administration")

    page = st.radio("Navigation", pages)


start, end = five_week_period()
all_vacations = get_vacations(start, end)


# ------------------------------------------------------------
# Page : planning overview
# ------------------------------------------------------------

if page == "Planning des 5 semaines":

    st.subheader("Vue générale – 5 semaines")

    st.caption(
        f"Période : {start.strftime('%d/%m/%Y')} "
        f"au {end.strftime('%d/%m/%Y')}"
    )

    users = list_users()

    regulars = [
        u for u in users
        if u["role"] in REGULAR_ROLES and u["active"]
    ]

    rotation = [
        u for u in users
        if u["role"] == "Prévisionniste" and u["active"]
    ]

    m1, m2, m3 = st.columns(3)

    m1.metric("Prévisionnistes en rotation", len(rotation))
    m2.metric("Vacations saisies", len(all_vacations))
    m3.metric("Places prévues", 70)

    for week in range(5):
        week_start = start + timedelta(days=7 * week)

        st.markdown(
            f"### Semaine {week + 1} · "
            f"{week_start.strftime('%d/%m')} – "
            f"{(week_start + timedelta(days=6)).strftime('%d/%m/%Y')}"
        )

        rows = []

        for offset, day_name in enumerate(DAYS):
            current_day = week_start + timedelta(days=offset)

            day_vacs = [
                v for v in all_vacations
                if v["vacation_date"] == current_day.isoformat()
            ]

            names = [v["full_name"] for v in day_vacs]

            rows.append({
                "Jour": (
                    f"{day_name} "
                    f"{current_day.strftime('%d/%m')}"
                ),
                "Boughanem Soukaina": (
                    next(
                        (
                            u["full_name"]
                            for u in regulars
                            if u["role"] == "Chef de service"
                        ),
                        "—"
                    ) if offset < 5 else "—"
                ),
                "Contrôleur": (
                    next(
                        (
                            u["full_name"]
                            for u in regulars
                            if u["role"] == "Contrôleur"
                        ),
                        "—"
                    ) if offset < 5 else "—"
                ),
                "Vacation matin 1": (
                    names[0] if len(names) > 0 else "Libre"
                ),
                "Vacation matin 2": (
                    names[1] if len(names) > 1 else "Libre"
                ),
            })

        st.dataframe(
            rows,
            use_container_width=True,
            hide_index=True
        )

    # --------------------------------------------------------
    # Export du planning final avec les vacations enregistrées
    # --------------------------------------------------------

    st.divider()
    st.subheader("📄 Télécharger le planning final")

    st.caption(
        "Choisissez la période souhaitée. Le PDF reprendra les vacations "
        "effectivement enregistrées dans l'application."
    )

    col_date1, col_date2 = st.columns(2)

    with col_date1:
        date_export_debut = st.date_input(
            "Date de début",
            value=start,
            key="date_export_debut"
        )

    with col_date2:
        date_export_fin = st.date_input(
            "Date de fin",
            value=start + timedelta(days=6),
            min_value=date_export_debut,
            key="date_export_fin"
        )

    if date_export_fin < date_export_debut:
        st.error("La date de fin doit être égale ou postérieure à la date de début.")
    else:
        vacations_periode = get_vacations(
            date_export_debut,
            date_export_fin
        )
        equipe_active = [
            u for u in list_users()
            if u["active"]
        ]

        if len(equipe_active) > 9:
            st.error(
                f"Impossible de générer le PDF : le modèle comporte 9 lignes, "
                f"mais {len(equipe_active)} comptes sont actifs. "
                "Désactivez les comptes inutilisés dans Administration."
            )
        else:
            if not vacations_periode:
                st.info(
                    "Aucune vacation n'est enregistrée pour cette période. "
                    "Le PDF sera tout de même généré avec les fonctions "
                    "régulières et les cases de vacations laissées vides."
                )

            try:
                pdf_final = creer_planning_final_pdf(
                    date_export_debut,
                    date_export_fin,
                    equipe_active,
                    vacations_periode
                )
                st.download_button(
                    label="📥 Télécharger le planning final en PDF",
                    data=pdf_final,
                    file_name=(
                        f"Planning_Final_SEM_"
                        f"{date_export_debut.strftime('%Y%m%d')}_"
                        f"{date_export_fin.strftime('%Y%m%d')}.pdf"
                    ),
                    mime="application/pdf",
                    use_container_width=True
                )
            except ValueError as e:
                st.error(str(e))


# ------------------------------------------------------------
# Page : enter/edit own vacations
# ------------------------------------------------------------

elif page == "Saisir mes vacations":

    st.subheader("Saisir ou modifier mes vacations")

    if current["role"] in REGULAR_ROLES:
        st.info(
            "Votre fonction est enregistrée comme régulière "
            "du lundi au vendredi. Les vacations tournantes "
            "du matin sont réservées aux cinq prévisionnistes."
        )

    else:
        chosen_week = st.selectbox(
            "Semaine",
            options=list(range(5)),
            format_func=lambda x: (
                f"Semaine {x+1} – "
                f"{(start + timedelta(days=7*x)).strftime('%d/%m/%Y')}"
            ),
        )

        week_start = start + timedelta(days=7 * chosen_week)
        week_days = [
            week_start + timedelta(days=i)
            for i in range(7)
        ]

        with db() as conn:
            my_rows = conn.execute(
                """
                SELECT * FROM vacations
                WHERE user_id=?
                AND vacation_date BETWEEN ? AND ?
                """,
                (
                    current["id"],
                    week_days[0].isoformat(),
                    week_days[-1].isoformat()
                ),
            ).fetchall()

        my_dates = {
            r["vacation_date"]: r
            for r in my_rows
        }

        st.write(
            "Coche les jours où tu souhaites assurer "
            "une vacation du matin."
        )

        for d in week_days:
            existing = my_dates.get(d.isoformat())

            col1, col2, col3 = st.columns([2, 2, 3])

            col1.write(
                f"**{DAYS[d.weekday()]} {d.strftime('%d/%m')}**"
            )

            if existing:
                col2.success("Déjà inscrite")
                note_default = existing["note"] or ""

                if col3.button(
                    "Retirer",
                    key=f"del_{d.isoformat()}"
                ):
                    delete_vacation(existing["id"], current["id"])
                    st.rerun()
            else:
                note_default = ""

                if col2.button(
                    "Choisir cette vacation",
                    key=f"add_{d.isoformat()}"
                ):
                    ok, message = save_vacation(
                        current["id"],
                        d,
                        ""
                    )
                    (st.success if ok else st.warning)(message)
                    st.rerun()

            if existing and note_default:
                col3.caption(f"Note : {note_default}")

        with st.expander("Ajouter ou modifier une observation"):
            dates_with_vacation = [
                d for d in week_days
                if d.isoformat() in my_dates
            ]

            if dates_with_vacation:
                selected = st.selectbox(
                    "Vacation concernée",
                    dates_with_vacation,
                    format_func=lambda d: (
                        f"{DAYS[d.weekday()]} "
                        f"{d.strftime('%d/%m/%Y')}"
                    ),
                )

                note = st.text_area(
                    "Observation",
                    value=my_dates[selected.isoformat()]["note"] or ""
                )

                if st.button("Enregistrer l'observation"):
                    save_vacation(
                        current["id"],
                        selected,
                        note
                    )
                    st.success("Observation enregistrée.")
                    st.rerun()
            else:
                st.caption(
                    "Aucune vacation saisie pour cette semaine."
                )


# ------------------------------------------------------------
# Page : account
# ------------------------------------------------------------

elif page == "Mon compte":

    st.subheader("Mon compte")
    st.write(f"Nom : **{current['full_name']}**")
    st.write(f"Fonction : **{current['role']}**")

    with st.form("password_form"):
        new_password = st.text_input(
            "Nouveau mot de passe",
            type="password"
        )
        confirm = st.text_input(
            "Confirmer le nouveau mot de passe",
            type="password"
        )
        change = st.form_submit_button("Modifier le mot de passe")

    if change:
        if len(new_password) < 10:
            st.error(
                "Le mot de passe doit contenir "
                "au moins 10 caractères."
            )
        elif new_password != confirm:
            st.error(
                "Les deux mots de passe ne correspondent pas."
            )
        else:
            change_password(current["id"], new_password)
            st.success("Mot de passe modifié.")

elif page == "Administration" and is_admin:

    st.title("Administration des comptes")

    users = list_users()

    # =====================================================
    # MODIFICATION DU ROLE D'UN COMPTE EXISTANT
    # =====================================================

    st.subheader("Gestion des rôles")

    if users:
        users_dict = {
            f"{u['full_name']} ({u['username']}) — {u['role']}":
            u["id"]
            for u in users
            if u["active"]
        }

        if users_dict:
            selected_user = st.selectbox(
                "Sélectionner un utilisateur",
                list(users_dict.keys()),
                key="admin_selected_user"
            )

            selected_user_id = users_dict[selected_user]

            new_role = st.selectbox(
                "Attribuer un rôle",
                ["Prévisionniste", "Contrôleur"],
                key="admin_new_role"
            )

            if st.button("Enregistrer le rôle", key="save_user_role"):
                update_user_role(selected_user_id, new_role)

                st.success(
                    f"Le rôle de {selected_user.split(' (')[0]} "
                    f"a été modifié en : {new_role}."
                )

                st.rerun()
        else:
            st.info("Aucun compte actif disponible.")

    st.divider()

    # =====================================================
    # CREATION D'UN NOUVEAU COMPTE
    # =====================================================

    st.subheader("Créer un compte")

    with st.form("create_user_form"):

        username = st.text_input("Nom d'utilisateur")
        full_name = st.text_input("Nom complet")

        role = st.selectbox(
            "Rôle",
            ["Contrôleur", "Prévisionniste"]
        )

        password = st.text_input(
            "Mot de passe",
            type="password"
        )

        submitted = st.form_submit_button("Créer le compte")

        if submitted:
            if not username.strip() or not full_name.strip() or not password:
                st.error("Veuillez remplir tous les champs.")
            elif len(password) < 10:
                st.error("Le mot de passe doit contenir au moins 10 caractères.")
            else:
                try:
                    create_user(username, full_name, role, password)
                    st.success("Compte créé avec succès.")
                    st.rerun()

                except sqlite3.IntegrityError:
                    st.error("Cet identifiant existe déjà. Choisissez-en un autre.")
                except Exception as e:
                    st.error(f"Erreur lors de la création : {e}")

    st.divider()

    # =====================================================
    # LISTE DES COMPTES
    # =====================================================

    st.subheader("Comptes existants")

    users = list_users()

    for u in users:

        col1, col2, col3, col4 = st.columns([3, 2, 1, 1])

        with col1:
            st.write(f"**{u['full_name']}**")
            st.caption(u["username"])

        with col2:
            st.write(u["role"])

        with col3:
            if u["active"]:
                st.success("Actif")
            else:
                st.error("Inactif")

        with col4:
            if u["active"]:
                if u["id"] == current["id"]:
                    st.caption("Compte connecté")
                elif st.button("Désactiver", key=f"disable_{u['id']}"):
                    set_user_active(u["id"], False)
                    st.rerun()
            else:
                if st.button("Activer", key=f"enable_{u['id']}"):
                    set_user_active(u["id"], True)
                    st.rerun()

        st.divider()
        
        
# ------------------------------------------------------------
# Footer
# ------------------------------------------------------------

st.divider()
st.caption(
    "Planning des vacations SEM / DRMCO."
)




