
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
    SimpleDocTemplate, Table, TableStyle,
    Paragraph, Spacer
)


# ============================================================
# CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Planning des vacations – SEM / DRMCO",
    page_icon="📅",
    layout="wide"
)

DB_PATH = "planning_sem.db"

DAYS = [
    "Lundi", "Mardi", "Mercredi", "Jeudi",
    "Vendredi", "Samedi", "Dimanche"
]

REGULAR_ROLES = ["Chef de service", "Contrôleur"]

LEAVE_REASONS = [
    "Congé annuel",
    "Congé maladie",
    "Mission",
    "Autorisation d'absence",
    "Récupération",
    "Autre"
]


# ============================================================
# DATABASE
# ============================================================

def db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def hash_password(password, salt=None):
    salt = salt or secrets.token_bytes(16)

    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        200_000
    )

    return salt.hex(), digest.hex()


def verify_password(password, salt_hex, hash_hex):
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
            CREATE TABLE IF NOT EXISTS leaves (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                start_date TEXT NOT NULL,
                end_date TEXT NOT NULL,
                reason TEXT NOT NULL,
                note TEXT DEFAULT '',
                created_at TEXT NOT NULL,
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
            "SELECT id FROM users WHERE username=?",
            ("soukaina",)
        ).fetchone()

        if not existing:

            salt, pw_hash = hash_password("ChangeMe123!")

            conn.execute("""
                INSERT INTO users
                (username, full_name, role, salt, password_hash)
                VALUES (?, ?, ?, ?, ?)
            """, (
                "soukaina",
                "Boughanem Soukaina",
                "Chef de service",
                salt,
                pw_hash
            ))


def get_user(username):

    with db() as conn:
        return conn.execute("""
            SELECT *
            FROM users
            WHERE username=? AND active=1
        """, (username,)).fetchone()


def list_users():

    with db() as conn:
        return conn.execute("""
            SELECT id, username, full_name, role, active
            FROM users
            ORDER BY id
        """).fetchall()


def create_user(username, full_name, role, password):

    salt, pw_hash = hash_password(password)

    with db() as conn:
        conn.execute("""
            INSERT INTO users
            (username, full_name, role, salt, password_hash)
            VALUES (?, ?, ?, ?, ?)
        """, (
            username.strip(),
            full_name.strip(),
            role,
            salt,
            pw_hash
        ))


def update_user_role(user_id, new_role):

    with db() as conn:
        conn.execute("""
            UPDATE users
            SET role=?
            WHERE id=?
        """, (new_role, user_id))


def change_password(user_id, new_password):

    salt, pw_hash = hash_password(new_password)

    with db() as conn:
        conn.execute("""
            UPDATE users
            SET salt=?, password_hash=?
            WHERE id=?
        """, (salt, pw_hash, user_id))


def set_user_active(user_id, active):

    with db() as conn:
        conn.execute("""
            UPDATE users
            SET active=?
            WHERE id=?
        """, (1 if active else 0, user_id))


# ============================================================
# VACATIONS
# ============================================================

def get_vacations(start_date, end_date):

    with db() as conn:
        return conn.execute("""
            SELECT v.id, v.user_id, v.vacation_date,
                   v.note, v.updated_at,
                   u.full_name, u.role
            FROM vacations v
            JOIN users u ON u.id=v.user_id
            WHERE v.vacation_date BETWEEN ? AND ?
            ORDER BY v.vacation_date, u.full_name
        """, (
            start_date.isoformat(),
            end_date.isoformat()
        )).fetchall()


def save_vacation(user_id, vacation_date, note=""):

    with db() as conn:

        existing = conn.execute("""
            SELECT id FROM vacations
            WHERE user_id=? AND vacation_date=?
        """, (
            user_id,
            vacation_date.isoformat()
        )).fetchone()

        leave = conn.execute("""
            SELECT id FROM leaves
            WHERE user_id=?
            AND start_date<=?
            AND end_date>=?
        """, (
            user_id,
            vacation_date.isoformat(),
            vacation_date.isoformat()
        )).fetchone()

        if leave:
            return False, "Vous êtes déclaré en congé à cette date."

        count = conn.execute("""
            SELECT COUNT(*) AS n
            FROM vacations
            WHERE vacation_date=?
        """, (vacation_date.isoformat(),)).fetchone()["n"]

        if existing:

            conn.execute("""
                UPDATE vacations
                SET note=?, updated_at=?
                WHERE user_id=? AND vacation_date=?
            """, (
                note,
                datetime.now().isoformat(timespec="seconds"),
                user_id,
                vacation_date.isoformat()
            ))

            return True, "Vacation mise à jour."

        if count >= 2:
            return False, "Les deux places de cette matinée sont déjà prises."

        conn.execute("""
            INSERT INTO vacations
            (user_id, vacation_date, note, updated_at)
            VALUES (?, ?, ?, ?)
        """, (
            user_id,
            vacation_date.isoformat(),
            note,
            datetime.now().isoformat(timespec="seconds")
        ))

        return True, "Vacation enregistrée."


def delete_vacation(vacation_id, user_id=None):

    with db() as conn:

        if user_id is None:
            conn.execute(
                "DELETE FROM vacations WHERE id=?",
                (vacation_id,)
            )
        else:
            conn.execute("""
                DELETE FROM vacations
                WHERE id=? AND user_id=?
            """, (vacation_id, user_id))


# ============================================================
# LEAVES - GESTION DES CONGES
# ============================================================

def save_leave(user_id, start_date, end_date, reason, note=""):

    if end_date < start_date:
        return False, "La date de fin doit être après la date de début."

    with db() as conn:

        overlap = conn.execute("""
            SELECT id FROM leaves
            WHERE user_id=?
            AND start_date<=?
            AND end_date>=?
        """, (
            user_id,
            end_date.isoformat(),
            start_date.isoformat()
        )).fetchone()

        if overlap:
            return False, "Une déclaration de congé existe déjà sur cette période."

        conn.execute("""
            INSERT INTO leaves
            (user_id, start_date, end_date, reason, note, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            user_id,
            start_date.isoformat(),
            end_date.isoformat(),
            reason,
            note.strip(),
            datetime.now().isoformat(timespec="seconds")
        ))

    return True, "Congé enregistré avec succès."


def get_leaves(start_date, end_date):

    with db() as conn:
        return conn.execute("""
            SELECT l.*, u.full_name, u.role
            FROM leaves l
            JOIN users u ON u.id=l.user_id
            WHERE l.start_date<=?
            AND l.end_date>=?
            ORDER BY l.start_date, u.full_name
        """, (
            end_date.isoformat(),
            start_date.isoformat()
        )).fetchall()


def get_my_leaves(user_id):

    with db() as conn:
        return conn.execute("""
            SELECT *
            FROM leaves
            WHERE user_id=?
            ORDER BY start_date DESC
        """, (user_id,)).fetchall()


def update_leave(
    leave_id,
    user_id,
    start_date,
    end_date,
    reason,
    note
):

    if end_date < start_date:
        return False, "Période incorrecte."

    with db() as conn:

        overlap = conn.execute("""
            SELECT id FROM leaves
            WHERE user_id=?
            AND id<>?
            AND start_date<=?
            AND end_date>=?
        """, (
            user_id,
            leave_id,
            end_date.isoformat(),
            start_date.isoformat()
        )).fetchone()

        if overlap:
            return False, "Une autre déclaration existe sur cette période."

        result = conn.execute("""
            UPDATE leaves
            SET start_date=?, end_date=?,
                reason=?, note=?
            WHERE id=? AND user_id=?
        """, (
            start_date.isoformat(),
            end_date.isoformat(),
            reason,
            note.strip(),
            leave_id,
            user_id
        ))

        if result.rowcount == 0:
            return False, "Déclaration introuvable."

    return True, "Congé modifié avec succès."


def delete_leave(leave_id, user_id=None):

    with db() as conn:

        if user_id is None:
            conn.execute(
                "DELETE FROM leaves WHERE id=?",
                (leave_id,)
            )
        else:
            conn.execute("""
                DELETE FROM leaves
                WHERE id=? AND user_id=?
            """, (leave_id, user_id))


# ============================================================
# DATES
# ============================================================

def monday_of(d):
    return d - timedelta(days=d.weekday())


def five_week_period():

    start = monday_of(date.today())
    end = start + timedelta(days=34)

    return start, end


# ============================================================
# EXPORT EXCEL - MODELE VIERGE
# ============================================================

def creer_modele_excel(date_debut=None, date_fin=None):

    if date_debut is None:
        date_debut = monday_of(date.today())

    if date_fin is None:
        date_fin = date_debut + timedelta(days=6)

    if date_fin < date_debut:
        raise ValueError("Période incorrecte.")

    dates = [
        date_debut + timedelta(days=i)
        for i in range((date_fin-date_debut).days+1)
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "Tableau de service"

    bleu = "D9EAF7"
    gris = "E7E6E6"

    bord = Side(style="thin", color="333333")
    bordures = Border(
        left=bord, right=bord,
        top=bord, bottom=bord
    )

    derniere_colonne = 4 + len(dates)
    derniere_lettre = get_column_letter(derniere_colonne)

    ws.merge_cells(f"A1:{derniere_lettre}1")
    ws["A1"] = "TABLEAU HEBDOMADAIRE DE SERVICE"
    ws["A1"].font = Font(
        name="Arial", size=15,
        bold=True, color="16324F"
    )
    ws["A1"].alignment = Alignment(
        horizontal="center", vertical="center"
    )
    ws.row_dimensions[1].height = 30

    info_cells = [
        ("A2:B2", "PÉRIODE DU :"),
        ("C2:D2", date_debut.strftime("%d/%m/%Y")),
        ("E2:F2", "AU :"),
        ("G2:H2", date_fin.strftime("%d/%m/%Y")),
        ("I2:J2", "EXEMPLAIRE DESTINÉ À :"),
        ("K2:L2", "TRANSMIS LE :"),
        ("M2:N2", "VISA :")
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

    ws.row_dimensions[2].height = 28

    fixed_headers = [
        ("A", "NOMS ET PRÉNOMS", 23),
        ("B", "GRADE", 10),
        ("C", "FONCTION", 14)
    ]

    for letter, title, width in fixed_headers:

        ws.merge_cells(f"{letter}4:{letter}5")

        cell = ws[f"{letter}4"]
        cell.value = title
        cell.font = Font(bold=True, size=9)
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )
        cell.fill = PatternFill("solid", fgColor=gris)

        ws.column_dimensions[letter].width = width

    for index, current_date in enumerate(dates):

        col = 4 + index

        cell = ws.cell(row=4, column=col)

        cell.value = (
            f"{DAYS[current_date.weekday()].upper()}\n"
            f"{current_date.strftime('%d/%m/%Y')}"
        )

        cell.font = Font(bold=True, size=8)
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )
        cell.fill = PatternFill("solid", fgColor=bleu)

        ws.merge_cells(
            start_row=4, start_column=col,
            end_row=5, end_column=col
        )

        ws.column_dimensions[
            get_column_letter(col)
        ].width = 13

    col_obs = 4 + len(dates)
    lettre_obs = get_column_letter(col_obs)

    ws.merge_cells(
        start_row=4, start_column=col_obs,
        end_row=5, end_column=col_obs
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

    for ligne in range(6, 15):

        for col in range(1, derniere_colonne+1):

            cell = ws.cell(row=ligne, column=col)
            cell.border = bordures
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

        ws.row_dimensions[ligne].height = 38

    for ligne in range(4, 6):
        for col in range(1, derniere_colonne+1):
            ws.cell(row=ligne, column=col).border = bordures

    ws.freeze_panes = "D6"
    ws.sheet_view.showGridLines = False

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_margins = PageMargins(
        left=0.2, right=0.2,
        top=0.3, bottom=0.3
    )

    ws.print_area = f"A1:{derniere_lettre}14"

    fichier = BytesIO()
    wb.save(fichier)
    fichier.seek(0)

    return fichier


# ============================================================
# TRI DE L'EQUIPE
# ============================================================

def trier_equipe(users):

    ordre = {
        "Chef de service": 0,
        "Contrôleur": 1,
        "Prévisionniste": 2
    }

    return sorted(
        [u for u in users if u["active"]],
        key=lambda u: (
            ordre.get(u["role"], 3),
            u["full_name"].casefold()
        )
    )


# ============================================================
# EXPORT PDF - MODELE VIERGE
# ============================================================

def creer_modele_pdf(date_debut=None, date_fin=None):

    if date_debut is None:
        date_debut = monday_of(date.today())

    if date_fin is None:
        date_fin = date_debut + timedelta(days=6)

    if date_fin < date_debut:
        raise ValueError("Période incorrecte.")

    dates = [
        date_debut + timedelta(days=i)
        for i in range((date_fin-date_debut).days+1)
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

    nombre_colonnes = 4 + len(dates)
    entetes = [""] * nombre_colonnes

    entetes[0] = Paragraph("NOMS ET<br/>PRÉNOMS", style_entete)
    entetes[1] = Paragraph("GRADE", style_entete)
    entetes[2] = Paragraph("FONCTION", style_entete)

    for index, d in enumerate(dates):
        entetes[3+index] = Paragraph(
            f"{DAYS[d.weekday()].upper()}<br/>{d.strftime('%d/%m/%Y')}",
            style_entete
        )

    entetes[-1] = Paragraph("OBSERVATIONS", style_entete)

    donnees = [entetes]

    for _ in range(9):
        donnees.append([""] * nombre_colonnes)

    largeur_page = landscape(A3)[0] - 36

    largeur_fixes = [88, 38, 52]
    largeur_obs = 105

    largeur_jour = (
        largeur_page - sum(largeur_fixes) - largeur_obs
    ) / len(dates)

    largeurs = (
        largeur_fixes
        + [largeur_jour] * len(dates)
        + [largeur_obs]
    )

    tableau = Table(
        donnees,
        colWidths=largeurs,
        rowHeights=[38] + [38]*9,
        repeatRows=1
    )

    tableau.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (2, 0), colors.lightgrey),
        ("BACKGROUND", (3, 0), (-2, 0), colors.HexColor("#D9EAF7")),
        ("BACKGROUND", (-1, 0), (-1, 0), colors.lightgrey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE")
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


# ============================================================
# EXPORT PDF - PLANNING FINAL AVEC CONGES
# OBSERVATIONS : PERIODES DU ... AU ...
# ============================================================

def creer_planning_final_pdf(
    date_debut,
    date_fin,
    users,
    vacations,
    leaves=None
):

    if date_fin < date_debut:
        raise ValueError("Période incorrecte.")

    leaves = leaves or []

    dates = [
        date_debut + timedelta(days=i)
        for i in range((date_fin-date_debut).days+1)
    ]

    equipe = trier_equipe(users)

    if len(equipe) > 9:
        raise ValueError(
            "Le tableau comporte seulement 9 lignes disponibles."
        )

    vacations_par_user = {}

    for vac in vacations:
        vacations_par_user.setdefault(
            vac["user_id"], {}
        )[vac["vacation_date"]] = vac

    leaves_par_user = {}

    for leave in leaves:
        leaves_par_user.setdefault(
            leave["user_id"], []
        ).append(leave)

    style_titre = ParagraphStyle(
        "TitreFinal",
        fontName="Helvetica-Bold",
        fontSize=16,
        alignment=TA_CENTER,
        spaceAfter=7
    )

    style_info = ParagraphStyle(
        "InfoFinal",
        fontName="Helvetica",
        fontSize=10,
        alignment=TA_LEFT
    )

    style_entete = ParagraphStyle(
        "EnteteFinal",
        fontName="Helvetica-Bold",
        fontSize=8,
        alignment=TA_CENTER,
        leading=9
    )

    style_cellule = ParagraphStyle(
        "CelluleFinal",
        fontName="Helvetica-Bold",
        fontSize=7,
        alignment=TA_CENTER,
        leading=8
    )

    style_obs = ParagraphStyle(
        "ObsFinal",
        fontName="Helvetica-Bold",
        fontSize=7,
        textColor=colors.red,
        alignment=4,
        leading=9
    )

    elements = [
        Paragraph(
            "TABLEAU DE SERVICE – PLANNING FINAL<br/>SEM – DRMCO",
            style_titre
        ),
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
        Paragraph("FONCTION", style_entete)
    ]

    for d in dates:
        entetes.append(
            Paragraph(
                f"{DAYS[d.weekday()].upper()}<br/>{d.strftime('%d/%m')}",
                style_entete
            )
        )

    entetes.append(
        Paragraph("OBSERVATIONS", style_entete)
    )

    donnees = [entetes]

    for user in equipe:

        nom = escape(str(user["full_name"] or ""))
        role = str(user["role"] or "")

        ligne = [
            Paragraph(nom, style_cellule),
            Paragraph("", style_cellule),
            Paragraph(escape(role), style_cellule)
        ]

        user_vacations = vacations_par_user.get(
            user["id"], {}
        )

        user_leaves = leaves_par_user.get(
            user["id"], []
        )

        # ----------------------------------------------------
        # REMPLISSAGE DES JOURS
        # ----------------------------------------------------

        for d in dates:

            key = d.isoformat()

            leave_for_day = next(
                (
                    leave for leave in user_leaves
                    if leave["start_date"] <= key <= leave["end_date"]
                ),
                None
            )

            if leave_for_day:

                contenu = "CONGÉ"

            elif role in REGULAR_ROLES:

                contenu = "Régulier" if d.weekday() < 5 else ""

            elif key in user_vacations:

                contenu = "Matin"

            else:

                contenu = ""

            ligne.append(
                Paragraph(escape(contenu), style_cellule)
            )

        # ----------------------------------------------------
        # OBSERVATIONS : CONGES DU ... AU ...
        # ----------------------------------------------------

        lignes_observations = []

        for leave in user_leaves:

            debut = date.fromisoformat(leave["start_date"])
            fin = date.fromisoformat(leave["end_date"])

            # Limiter les dates à la période du planning exporté
            debut_affiche = max(debut, date_debut)
            fin_affiche = min(fin, date_fin)

            # Vérifier que le congé chevauche la période exportée
            if debut_affiche <= fin_affiche:

                motif = escape(
                    str(leave["reason"] or "Congé").strip()
                )

                observation = (
                    f"<b>{motif}</b> : "
                    f"du {debut_affiche.strftime('%d/%m/%Y')} "
                    f"au {fin_affiche.strftime('%d/%m/%Y')}"
                )

                note = (leave["note"] or "").strip()

                if note:
                    observation += (
                        f"<br/><i>{escape(note)}</i>"
                    )

                lignes_observations.append(observation)

        # ----------------------------------------------------
        # OBSERVATIONS DES VACATIONS
        # ----------------------------------------------------

        for vacation_date, vacation in user_vacations.items():

            note = (vacation["note"] or "").strip()

            if note:

                jour = date.fromisoformat(vacation_date)

                if date_debut <= jour <= date_fin:

                    lignes_observations.append(
                        f"<b>Vacation du "
                        f"{jour.strftime('%d/%m/%Y')}</b> : "
                        f"{escape(note)}"
                    )

        obs = (
            "<br/>".join(lignes_observations)
            if lignes_observations
            else " "
        )

        ligne.append(
            Paragraph(obs, style_obs)
        )

        donnees.append(ligne)

    while len(donnees) < 10:
        donnees.append([""] * len(entetes))

    largeur_page = landscape(A3)[0] - 36

    largeur_nom = 105
    largeur_grade = 35
    largeur_fonction = 65
    largeur_obs = 145

    largeur_jour = (
        largeur_page
        - largeur_nom
        - largeur_grade
        - largeur_fonction
        - largeur_obs
    ) / len(dates)

    largeurs = (
        [largeur_nom, largeur_grade, largeur_fonction]
        + [largeur_jour] * len(dates)
        + [largeur_obs]
    )

    tableau = Table(
        donnees,
        colWidths=largeurs,
        rowHeights=[67] + [66]*(len(donnees)-1),
        repeatRows=1
    )

    tableau.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (2, 0), colors.lightgrey),
        ("BACKGROUND", (3, 0), (-2, 0), colors.HexColor("#D9EAF7")),
        ("BACKGROUND", (-1, 0), (-1, 0), colors.lightgrey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2)
    ]))

    elements.append(tableau)

    style_signature = ParagraphStyle(
        "Signature",
        fontName="Helvetica-Bold",
        fontSize=8,
        alignment=TA_CENTER,
        leading=10
    )

    signature = Table(
        [[Paragraph("<br/><br/><br/>Signature :", style_signature)]],
        colWidths=[180],
        rowHeights=[35]
    )

    signature.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "TOP")
    ]))

    bloc_signature = Table(
        [["", signature]],
        colWidths=[largeur_page-180, 180]
    )

    bloc_signature.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP")
    ]))

    elements.append(Spacer(1, 3))
    elements.append(bloc_signature)

    fichier = BytesIO()

    document = SimpleDocTemplate(
        fichier,
        pagesize=landscape(A3),
        rightMargin=18,
        leftMargin=18,
        topMargin=12,
        bottomMargin=5
    )

    document.build(elements)
    fichier.seek(0)

    return fichier


# ============================================================
# INITIALISATION
# ============================================================

init_db()


# ============================================================
# STYLE
# ============================================================

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

.leave-label {
    color: #C0392B;
    font-weight: bold;
}
</style>
""", unsafe_allow_html=True)

st.title("📅 Planning des vacations")
st.caption("Service Exploitation Météorologique – DRMCO")


# ============================================================
# LOGIN
# ============================================================

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
                "role": user["role"]
            }

            st.rerun()

        else:
            st.error("Identifiant ou mot de passe incorrect.")

    st.stop()


current = st.session_state.user

is_admin = current["role"] == "Chef de service"


# ============================================================
# SIDEBAR
# ============================================================

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
all_leaves = get_leaves(start, end)


# ============================================================
# PAGE 1 : PLANNING GENERAL
# ============================================================

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

    c1, c2, c3, c4 = st.columns(4)

    c1.metric("Prévisionnistes", len(rotation))
    c2.metric("Vacations saisies", len(all_vacations))
    c3.metric("Places prévues", 70)
    c4.metric("Déclarations de congé", len(all_leaves))

    for week in range(5):

        week_start = start + timedelta(days=7*week)

        st.markdown(
            f"### Semaine {week+1} · "
            f"{week_start.strftime('%d/%m')} – "
            f"{(week_start+timedelta(days=6)).strftime('%d/%m/%Y')}"
        )

        rows = []

        for offset, day_name in enumerate(DAYS):

            current_day = week_start + timedelta(days=offset)

            day_key = current_day.isoformat()

            day_vacs = [
                v for v in all_vacations
                if v["vacation_date"] == day_key
            ]

            day_leaves = [
                leave for leave in all_leaves
                if leave["start_date"] <= day_key <= leave["end_date"]
            ]

            names = [v["full_name"] for v in day_vacs]

            leaves_names = [
                leave["full_name"]
                for leave in day_leaves
            ]

            chef = next(
                (
                    u["full_name"] for u in regulars
                    if u["role"] == "Chef de service"
                ),
                "—"
            )

            controleur = next(
                (
                    u["full_name"] for u in regulars
                    if u["role"] == "Contrôleur"
                ),
                "—"
            )

            if chef != "—" and chef in leaves_names:
                chef_display = f"{chef} (CONGÉ)"
            else:
                chef_display = chef if offset < 5 else "—"

            if controleur != "—" and controleur in leaves_names:
                controleur_display = f"{controleur} (CONGÉ)"
            else:
                controleur_display = controleur if offset < 5 else "—"

            rows.append({
                "Jour": (
                    f"{day_name} {current_day.strftime('%d/%m')}"
                ),
                "Chef de service": chef_display,
                "Contrôleur": controleur_display,
                "Vacation matin 1": (
                    names[0] if len(names) > 0 else "Libre"
                ),
                "Vacation matin 2": (
                    names[1] if len(names) > 1 else "Libre"
                ),
                "Agents en congé": (
                    ", ".join(leaves_names) if leaves_names else "—"
                )
            })

        st.dataframe(
            rows,
            use_container_width=True,
            hide_index=True
        )

    # --------------------------------------------------------
    # EXPORT PDF FINAL
    # --------------------------------------------------------

    st.divider()

    st.subheader("📄 Télécharger le planning final")

    st.caption(
        "Le PDF reprend les vacations et les congés enregistrés."
    )

    col1, col2 = st.columns(2)

    with col1:
        date_export_debut = st.date_input(
            "Date de début",
            value=start,
            key="date_export_debut"
        )

    with col2:
        date_export_fin = st.date_input(
            "Date de fin",
            value=start+timedelta(days=6),
            key="date_export_fin"
        )

    if date_export_fin < date_export_debut:

        st.error("La date de fin doit être après la date de début.")

    else:

        vacations_periode = get_vacations(
            date_export_debut,
            date_export_fin
        )

        leaves_periode = get_leaves(
            date_export_debut,
            date_export_fin
        )

        equipe_active = [
            u for u in list_users()
            if u["active"]
        ]

        if len(equipe_active) > 9:

            st.error(
                "Le modèle PDF comporte seulement 9 lignes. "
                "Désactivez les comptes inutilisés."
            )

        else:

            try:

                pdf_final = creer_planning_final_pdf(
                    date_export_debut,
                    date_export_fin,
                    equipe_active,
                    vacations_periode,
                    leaves_periode
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


# ============================================================
# PAGE 2 : VACATIONS ET CONGES PERSONNELS
# ============================================================

elif page == "Saisir mes vacations":

    st.subheader("Saisir ou modifier mes vacations")

    if current["role"] in REGULAR_ROLES:

        st.info(
            "Votre fonction est enregistrée comme régulière "
            "du lundi au vendredi. Les vacations tournantes "
            "du matin sont réservées aux prévisionnistes."
        )

    else:

        chosen_week = st.selectbox(
            "Semaine",
            options=list(range(5)),
            format_func=lambda x: (
                f"Semaine {x+1} – "
                f"{(start+timedelta(days=7*x)).strftime('%d/%m/%Y')}"
            )
        )

        week_start = start + timedelta(days=7*chosen_week)

        week_days = [
            week_start+timedelta(days=i)
            for i in range(7)
        ]

        with db() as conn:

            my_rows = conn.execute("""
                SELECT * FROM vacations
                WHERE user_id=?
                AND vacation_date BETWEEN ? AND ?
            """, (
                current["id"],
                week_days[0].isoformat(),
                week_days[-1].isoformat()
            )).fetchall()

        my_dates = {
            r["vacation_date"]: r
            for r in my_rows
        }

        st.write(
            "Cochez les jours où vous souhaitez assurer "
            "une vacation du matin."
        )

        for d in week_days:

            existing = my_dates.get(d.isoformat())

            col1, col2, col3 = st.columns([2, 2, 3])

            col1.write(
                f"**{DAYS[d.weekday()]} {d.strftime('%d/%m')}**"
            )

            if existing:

                col2.success("Déjà inscrit(e)")

                if col3.button(
                    "Retirer",
                    key=f"del_{d.isoformat()}"
                ):

                    delete_vacation(
                        existing["id"],
                        current["id"]
                    )

                    st.rerun()

            else:

                if col2.button(
                    "Choisir cette vacation",
                    key=f"add_{d.isoformat()}"
                ):

                    ok, message = save_vacation(
                        current["id"],
                        d,
                        ""
                    )

                    if ok:
                        st.success(message)
                    else:
                        st.warning(message)

                    st.rerun()

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
                        f"{DAYS[d.weekday()]} {d.strftime('%d/%m/%Y')}"
                    )
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

    # ========================================================
    # DECLARATION DES CONGES
    # ========================================================

    st.divider()

    st.subheader("🏖️ Déclarer mes congés")

    st.info(
        "Chaque agent peut déclarer ses congés directement "
        "depuis son compte."
    )

    with st.form("leave_form"):

        leave_start = st.date_input(
            "Date de début",
            value=date.today(),
            key="leave_start"
        )

        leave_end = st.date_input(
            "Date de fin",
            value=leave_start,
            key="leave_end"
        )

        leave_reason = st.selectbox(
            "Motif du congé",
            LEAVE_REASONS
        )

        leave_note = st.text_area(
            "Observation complémentaire",
            placeholder="Précisions éventuelles..."
        )

        submitted_leave = st.form_submit_button(
            "Enregistrer mes congés",
            use_container_width=True
        )

        if submitted_leave:

            if leave_end < leave_start:

                st.error("La période sélectionnée est incorrecte.")

            else:

                ok, message = save_leave(
                    current["id"],
                    leave_start,
                    leave_end,
                    leave_reason,
                    leave_note
                )

                if ok:
                    st.success(message)
                    st.rerun()
                else:
                    st.error(message)

    # ========================================================
    # LISTE ET MODIFICATION DES CONGES PERSONNELS
    # ========================================================

    st.divider()

    st.subheader("Mes congés enregistrés")

    my_leaves = get_my_leaves(current["id"])

    if my_leaves:

        for leave in my_leaves:

            with st.expander(
                f"{leave['reason']} : "
                f"{leave['start_date']} au {leave['end_date']}"
            ):

                st.write(
                    f"**Observation :** {leave['note'] or 'Aucune'}"
                )

                col1, col2 = st.columns(2)

                with col1:

                    if st.button(
                        "Modifier",
                        key=f"edit_leave_{leave['id']}"
                    ):

                        st.session_state.edit_leave_id = leave["id"]
                        st.rerun()

                with col2:

                    if st.button(
                        "Supprimer",
                        key=f"delete_leave_{leave['id']}"
                    ):

                        delete_leave(
                            leave["id"],
                            current["id"]
                        )

                        st.session_state.pop("edit_leave_id", None)
                        st.success("Déclaration supprimée.")
                        st.rerun()

                if st.session_state.get("edit_leave_id") == leave["id"]:

                    with st.form(f"edit_leave_form_{leave['id']}"):

                        new_start = st.date_input(
                            "Nouvelle date de début",
                            value=date.fromisoformat(leave["start_date"])
                        )

                        new_end = st.date_input(
                            "Nouvelle date de fin",
                            value=date.fromisoformat(leave["end_date"])
                        )

                        new_reason = st.selectbox(
                            "Motif",
                            LEAVE_REASONS,
                            index=(
                                LEAVE_REASONS.index(leave["reason"])
                                if leave["reason"] in LEAVE_REASONS
                                else 0
                            )
                        )

                        new_note = st.text_area(
                            "Observation",
                            value=leave["note"] or ""
                        )

                        save_edit = st.form_submit_button(
                            "Enregistrer les modifications"
                        )

                        if save_edit:

                            ok, message = update_leave(
                                leave["id"],
                                current["id"],
                                new_start,
                                new_end,
                                new_reason,
                                new_note
                            )

                            if ok:

                                st.session_state.pop(
                                    "edit_leave_id",
                                    None
                                )

                                st.success(message)
                                st.rerun()

                            else:
                                st.error(message)

    else:

        st.info("Vous n'avez déclaré aucun congé.")


# ============================================================
# PAGE 3 : MON COMPTE
# ============================================================

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
                "Le mot de passe doit contenir au moins 10 caractères."
            )

        elif new_password != confirm:

            st.error("Les deux mots de passe ne correspondent pas.")

        else:

            change_password(
                current["id"],
                new_password
            )

            st.success("Mot de passe modifié.")


# ============================================================
# PAGE 4 : ADMINISTRATION
# ============================================================

elif page == "Administration" and is_admin:

    st.title("Administration des comptes")

    users = list_users()

    # ========================================================
    # MODIFICATION DES ROLES
    # ========================================================

    st.subheader("Gestion des rôles")

    users_dict = {
        f"{u['full_name']} ({u['username']}) — {u['role']}": u["id"]
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

        if st.button(
            "Enregistrer le rôle",
            key="save_user_role"
        ):

            update_user_role(
                selected_user_id,
                new_role
            )

            st.success("Rôle modifié.")
            st.rerun()

    else:

        st.info("Aucun compte actif disponible.")

    # ========================================================
    # CREATION DES COMPTES
    # ========================================================

    st.divider()

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

            if (
                not username.strip()
                or not full_name.strip()
                or not password
            ):

                st.error("Veuillez remplir tous les champs.")

            elif len(password) < 10:

                st.error(
                    "Le mot de passe doit contenir au moins 10 caractères."
                )

            else:

                try:

                    create_user(
                        username,
                        full_name,
                        role,
                        password
                    )

                    st.success("Compte créé avec succès.")
                    st.rerun()

                except sqlite3.IntegrityError:

                    st.error(
                        "Cet identifiant existe déjà."
                    )

                except Exception as e:

                    st.error(f"Erreur : {e}")

    # ========================================================
    # COMPTES EXISTANTS
    # ========================================================

    st.divider()

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

                elif st.button(
                    "Désactiver",
                    key=f"disable_{u['id']}"
                ):

                    set_user_active(u["id"], False)
                    st.rerun()

            else:

                if st.button(
                    "Activer",
                    key=f"enable_{u['id']}"
                ):

                    set_user_active(u["id"], True)
                    st.rerun()

        st.divider()

    # ========================================================
    # SUIVI DES CONGES DE L'EQUIPE
    # ========================================================

    st.subheader("🏖️ Suivi des congés de l'équipe")

    admin_leaves = get_leaves(
        date(2000, 1, 1),
        date(2100, 12, 31)
    )

    if admin_leaves:

        for leave in admin_leaves:

            with st.expander(
                f"{leave['full_name']} – {leave['reason']} "
                f"({leave['start_date']} au {leave['end_date']})"
            ):

                st.write(f"**Fonction :** {leave['role']}")

                st.write(
                    f"**Du :** {leave['start_date']}"
                )

                st.write(
                    f"**Au :** {leave['end_date']}"
                )

                st.write(
                    f"**Motif :** {leave['reason']}"
                )

                st.write(
                    f"**Observation :** {leave['note'] or 'Aucune'}"
                )

                if st.button(
                    "Supprimer cette déclaration",
                    key=f"admin_delete_leave_{leave['id']}"
                ):

                    delete_leave(leave["id"])
                    st.success("Déclaration supprimée.")
                    st.rerun()

    else:

        st.info("Aucun congé enregistré.")


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "Planning des vacations SEM / DRMCO."
)
