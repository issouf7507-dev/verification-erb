"""Vérification de l'ERB Sanlam contre l'ERB généré par Finaptik.

Les deux ERB ont la même structure : Date | N° Pièce | N° Chèque | Libellé |
Débit/Crédit « à passer par Sanlam » | Débit/Crédit « à passer par la banque » (l'ERB Sanlam a en plus Banque et
Compte avant la date, et un Commentaire après). La ligne d'en-tête est repérée par la cellule « Date ».

Sortie (3 onglets) : la feuille ERB de Finaptik + « a passez par sanlam » + « a passer par la banque », Sanlam et
Finaptik côte à côte, Débit puis Crédit, triés par montant ; une ligne sans correspondance reste seule (statut KO).
Rapprochés : même montant dans la même colonne ; à montant égal on préfère même N° de pièce, puis même libellé,
puis la date la plus proche.
Extourne possible : ligne GL de l'ERB Finaptik (« à passer par la banque ») retrouvée côté banque dans l'ERB Sanlam
(« à passer par Sanlam »), même colonne Débit/Crédit, même montant et même pièce / n° de chèque / libellé.
Annulation même ERB : entre les deux onglets, dans un même ERB, même montant — « à passer par Sanlam · Débit » ↔
« à passer par la banque · Crédit » et « à passer par Sanlam · Crédit » ↔ « à passer par la banque · Débit »
(même n° de chèque d'abord, puis date la plus proche).

Rangement (comme les suspens), dans le dossier où on lance la commande : verification/<MOIS ANNÉE>/<dossier de l'ERB Sanlam>/<BANQUE> VERIFICATION ERB <MOIS ANNÉE>.xlsx
  le mois vient de « RAPPROCHEMENT AU … » de l'ERB Finaptik ; un 3e argument impose un autre chemin de sortie.

Usage : .venv/bin/python verifier_erb.py "ERB SANLAM.xlsx" "ERB FINAPTIK.xlsx" ["SORTIE.xlsx"]
"""

import re
import sys
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill

BLEU = PatternFill('solid', fgColor='FF0070C0')
VIOLET = PatternFill('solid', fgColor='FF7030A0')
BLANC_GRAS = Font(bold=True, color='FFFFFFFF')
GRAS = Font(bold=True)
VERT = PatternFill('solid', fgColor='FFC6EFCE')
ROUGE = PatternFill('solid', fgColor='FFFFC7CE')
ORANGE = PatternFill('solid', fgColor='FFFFEB9C')
FMT_MONTANT = '#,##0'
FMT_DATE = 'dd/mm/yyyy'
COTES = ('sanlam_debit', 'sanlam_credit', 'banque_debit', 'banque_credit')
RACINE = Path.cwd()  # le dossier « verification » est créé là où on lance la commande
MOIS = ('JANVIER', 'FEVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN', 'JUILLET', 'AOUT', 'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE',
        'DECEMBRE')


def montant(v):
    if v in (None, ''):
        return None
    try:
        v = round(float(v), 2)
    except (TypeError, ValueError):
        return None
    return v or None


def en_date(v):
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, str) and len(v) >= 10:
        for fmt, n in (('%Y-%m-%d', 10), ('%d/%m/%Y', 10)):
            try:
                return datetime.strptime(v.strip()[:n], fmt)
            except ValueError:
                pass
    return None


def normaliser(texte):
    return ' '.join(str(texte or '').upper().split())


# ---------------------------------------------------------------- lecture d'un ERB

def lire_erb(chemin):
    """Retourne {côté: [lignes]} pour les 4 colonnes de montants de l'ERB (première feuille)."""
    wb = openpyxl.load_workbook(chemin, data_only=True)  # valeurs calculées (libellés en formule « =+D12 »)
    ws = wb.worksheets[0]
    lignes = {k: [] for k in COTES}
    base = None
    for r in ws.iter_rows(values_only=True):
        r = list(r)
        if base is None:
            if 'Date' in [str(v or '').strip() for v in r]:
                base = [str(v or '').strip() for v in r].index('Date')
            continue
        r += [None] * (base + 9 - len(r))
        d = en_date(r[base])
        if d is None:
            continue
        ligne = {'banque': r[0] if base >= 2 else None, 'compte': r[1] if base >= 2 else None, 'date': d,
                 'piece': r[base + 1], 'cheque': r[base + 2], 'libelle': r[base + 3]}
        for k, cote in enumerate(COTES):
            m = montant(r[base + 4 + k])
            if m is not None:
                lignes[cote].append({**ligne, 'montant': m})
    if base is None:
        sys.exit(f'En-tête « Date » introuvable dans {chemin}')
    return lignes


def date_arrete(chemin):
    """Date de « RAPPROCHEMENT AU 30/06/2026 » dans l'en-tête de l'ERB."""
    ws = openpyxl.load_workbook(chemin, read_only=True, data_only=True).worksheets[0]
    for r in ws.iter_rows(max_row=10, values_only=True):
        for v in r:
            m = re.search(r'RAPPROCHEMENT AU\s*(\d{2}/\d{2}/\d{4})', str(v or '').upper())
            if m:
                return datetime.strptime(m.group(1), '%d/%m/%Y')
    return None


# ---------------------------------------------------------------- appariement

def ecart_jours(a, b):
    if a['date'] is None or b['date'] is None:
        return 10 ** 6
    return abs((a['date'] - b['date']).days)


def apparier(gauche, droite):
    """Paires (g, d) triées par montant ; g ou d vaut None quand il n'y a pas de correspondance."""
    restant = {}
    for d in droite:
        restant.setdefault(d['montant'], []).append(d)
    paires = []
    for g in sorted(gauche, key=lambda x: (x['montant'], x['date'])):
        candidats = restant.get(g['montant'], [])
        if candidats:
            piece, lib = normaliser(g['piece']), normaliser(g['libelle'])
            meilleur = min(candidats, key=lambda d: (not piece or normaliser(d['piece']) != piece,
                                                     normaliser(d['libelle']) != lib, ecart_jours(g, d)))
            candidats.remove(meilleur)
            paires.append((g, meilleur))
        else:
            paires.append((g, None))
    for liste in restant.values():
        paires.extend((None, d) for d in liste)
    paires.sort(key=lambda p: ((p[0] or p[1])['montant'], p[0] is None))
    return paires


LIBELLES = {'sanlam_debit': 'à passer par Sanlam · Débit', 'sanlam_credit': 'à passer par Sanlam · Crédit',
            'banque_debit': 'à passer par la banque · Débit', 'banque_credit': 'à passer par la banque · Crédit'}
FEUILLES = (('a passez par sanlam', 'à passer par Sanlam', ('sanlam_debit', 'sanlam_credit')),
            ('a passer par la banque', 'à passer par la banque', ('banque_credit', 'banque_debit')))
BLEU_CLAIR = PatternFill('solid', fgColor='FFDDEBF7')


def numeros_cheque(ligne):
    texte = ' '.join(str(ligne.get(k) or '') for k in ('libelle', 'cheque'))
    return set(re.findall(r'(?<!\d)\d{7}(?!\d)', texte))


def apparier_seules(gauches, droites, lien):
    """Apparie deux listes d'éléments seuls de même montant ; `lien(a, b)` sert à la 1re passe (lien fort),
    la 2e passe ne regarde que la date. Remplit 'avec' sur les deux éléments."""
    for passe_forte in (True, False):
        for a in gauches:
            if a.get('avec'):
                continue
            candidats = [b for b in droites if not b.get('avec') and b['montant'] == a['montant']
                         and (not passe_forte or lien(a, b))]
            if candidats:
                b = min(candidats, key=lambda b: ecart_jours(a, b))
                a['avec'], b['avec'] = b, a


def construire(sanlam, finaptik):
    """Éléments {cote, s, f, montant, date, categorie, avec} pour les 4 côtés de l'ERB."""
    elements = []
    for cote in COTES:
        for s, f in apparier(sanlam[cote], finaptik[cote]):
            x = s or f
            elements.append({'cote': cote, 's': s, 'f': f, 'montant': x['montant'], 'date': x['date'],
                             'source': 'ok' if s and f else ('s' if s else 'f'), 'avec': None})

    # Extourne possible : ligne GL de l'ERB Finaptik (« à passer par la banque ») retrouvée côté banque dans l'ERB
    # Sanlam (« à passer par Sanlam »), même colonne Débit/Crédit, même montant, même pièce / n° de chèque / libellé.
    def meme_ligne(a, b):
        sa, fb = a['s'], b['f']
        pa = normaliser(sa['piece'])
        return (pa and pa == normaliser(fb['piece'])) or bool(numeros_cheque(sa) & numeros_cheque(fb)) \
            or normaliser(sa['libelle']) == normaliser(fb['libelle'])
    seuls_s = [e for e in elements if e['source'] == 's']
    seuls_f = [e for e in elements if e['source'] == 'f']
    for a in seuls_s:
        if not a['cote'].startswith('sanlam'):
            continue
        cote_gl = a['cote'].replace('sanlam', 'banque')
        candidats = [b for b in seuls_f if not b['avec'] and b['cote'] == cote_gl and b['montant'] == a['montant']
                     and meme_ligne(a, b)]
        if candidats:
            b = min(candidats, key=lambda b: ecart_jours(a, b))
            a['avec'], b['avec'] = b, a
    for e in elements:
        if e['avec']:
            e['categorie'] = 'extourne'

    # Annulation même ERB : entre les deux onglets, dans une même source (même ERB), même montant —
    # « à passer par Sanlam · Débit » ↔ « à passer par la banque · Crédit », et Crédit ↔ Débit
    for c0, c1 in (('sanlam_debit', 'banque_credit'), ('sanlam_credit', 'banque_debit')):
        for source, cle in (('s', 's'), ('f', 'f')):
            g = [e for e in elements if e['source'] == source and e['cote'] == c0 and not e['avec']]
            d = [e for e in elements if e['source'] == source and e['cote'] == c1 and not e['avec']]
            apparier_seules(g, d, lambda a, b: bool(numeros_cheque(a[cle]) & numeros_cheque(b[cle])))
    for e in elements:
        if e['avec'] and 'categorie' not in e:
            e['categorie'] = 'annulation'
        e.setdefault('categorie', {'ok': 'ok', 's': 'sanlam_seul', 'f': 'finaptik_seul'}[e['source']])
    return elements


# ---------------------------------------------------------------- écriture

CATEGORIES = ('ok', 'sanlam_seul', 'finaptik_seul', 'extourne', 'annulation')
TITRES = ['N°', 'Banque', 'Compte', 'Date', 'N° Pièce', 'N° Chèque', 'Libellé', 'Débit', 'Crédit',
          'Date', 'N° Pièce', 'N° Chèque', 'Libellé', 'Débit', 'Crédit', 'statut', 'Extourne possible',
          'Annulation même ERB', 'Explication']


def feuille(wb, nom, titre, cotes, elements):
    ws = wb.create_sheet(nom)
    for debut, fin, texte, couleur in (('B', 'I', f'SANLAM — {titre}', BLEU), ('J', 'O', f'FINAPTIK — {titre}', VIOLET)):
        ws.merge_cells(f'{debut}1:{fin}1')
        c = ws[f'{debut}1']
        c.value, c.fill, c.font, c.alignment = texte, couleur, BLANC_GRAS, Alignment(horizontal='center')
    for col, t in enumerate(TITRES, 1):
        c = ws.cell(2, col, t)
        c.font, c.fill = GRAS, BLEU_CLAIR
    ws.freeze_panes = 'A3'

    numero = 0
    for ligne, e in enumerate(sorted((e for e in elements if e['cote'] in cotes), key=lambda e: cotes.index(e['cote'])), 3):  # débit puis crédit, par montant
        e['feuille'], e['ligne'] = nom, ligne
        if e['s']:
            numero += 1
        ecrire_element(ws, ligne, numero if e['s'] else None, e)

    for row in ws.iter_rows(min_row=3):
        for c in row:
            if c.column in (4, 10):
                c.number_format = FMT_DATE
            elif c.column in (8, 9, 14, 15):
                c.number_format = FMT_MONTANT
    plage = f'P3:P{ws.max_row}'
    ws.conditional_formatting.add(plage, CellIsRule(operator='equal', formula=['"OK"'], fill=VERT))
    ws.conditional_formatting.add(plage, CellIsRule(operator='equal', formula=['"KO"'], fill=ROUGE))
    for col, largeur in {'A': 6, 'D': 11, 'E': 12, 'F': 12, 'G': 50, 'H': 14, 'I': 14, 'J': 11, 'K': 12, 'L': 12,
                         'M': 50, 'N': 14, 'O': 14, 'P': 8, 'Q': 45, 'R': 45, 'S': 30}.items():
        ws.column_dimensions[col].width = largeur
    return ws


def ecrire_element(ws, ligne, n, e):
    k = 1 if e['cote'].endswith('credit') else 0  # 0 = colonne Débit, 1 = colonne Crédit
    ws.cell(ligne, 1, n)
    if e['s']:
        s = e['s']
        for col, v in zip(range(2, 8), (s['banque'], s['compte'], s['date'], s['piece'], s['cheque'], s['libelle'])):
            ws.cell(ligne, col, v)
        ws.cell(ligne, 8 + k, s['montant'])
    if e['f']:
        f = e['f']
        for col, v in zip(range(10, 14), (f['date'], f['piece'], f['cheque'], f['libelle'])):
            ws.cell(ligne, col, v)
        ws.cell(ligne, 14 + k, f['montant'])
    ws.cell(ligne, 16, f'=IF({"HI"[k]}{ligne}={"NO"[k]}{ligne},"OK","KO")')


def renvois(wb, elements):
    """Colonnes Extourne possible / Annulation même ERB, une fois toutes les lignes placées."""
    for e in elements:
        autre = e['avec']
        if not autre:
            continue
        ws = wb[e['feuille']]
        ou = f'{LIBELLES[autre["cote"]]} (onglet « {autre["feuille"]} », ligne {autre["ligne"]})'
        if e['categorie'] == 'extourne':
            qui = 'Finaptik' if e['source'] == 's' else 'Sanlam'
            ws.cell(e['ligne'], 17, f'Oui ↔ {qui} : {ou}').fill = ORANGE
        else:
            ws.cell(e['ligne'], 18, f'Oui ↔ {ou}').fill = ORANGE


def afficher(elements):
    for _, titre, cotes in FEUILLES:
        es = [e for e in elements if e['cote'] in cotes]
        compte = {cat: sum(1 for e in es if e['categorie'] == cat) for cat in CATEGORIES}
        print(f'{titre:24} OK {compte["ok"]:4} | Sanlam seul {compte["sanlam_seul"]:3} | Finaptik seul '
              f'{compte["finaptik_seul"]:3} | extourne possible {compte["extourne"]:3} | '
              f'annulation même ERB {compte["annulation"]:3} lignes')


def main():
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    erb_sanlam, erb_finaptik = Path(sys.argv[1]), Path(sys.argv[2])
    periode_s, periode_f = date_arrete(erb_sanlam), date_arrete(erb_finaptik)
    if periode_s and periode_f and periode_s != periode_f:
        print(f'ATTENTION : ERB Sanlam au {periode_s:%d/%m/%Y}, ERB Finaptik au {periode_f:%d/%m/%Y}')
    if len(sys.argv) > 3:
        sortie = Path(sys.argv[3])
    else:  # même rangement que les suspens : verification/JUIN 2026/1-STANBIC/STANBIC VERIFICATION ERB JUIN 2026.xlsx
        periode = periode_f or periode_s
        mois = f'{MOIS[periode.month - 1]} {periode.year}' if periode else f'{date.today():%Y%m%d}'
        dossier = erb_sanlam.resolve().parent.name
        banque = re.sub(r'^\d+-', '', dossier)
        sortie = RACINE / 'verification' / mois / dossier / f'{banque} VERIFICATION ERB {mois}.xlsx'
    sortie.parent.mkdir(parents=True, exist_ok=True)

    elements = construire(lire_erb(erb_sanlam), lire_erb(erb_finaptik))
    wb = openpyxl.load_workbook(erb_finaptik)
    for nom in tuple(f[0] for f in FEUILLES):  # relance sur un fichier déjà vérifié
        if nom in wb.sheetnames:
            del wb[nom]
    for nom, titre, cotes in FEUILLES:
        feuille(wb, nom, titre, cotes, elements)
    renvois(wb, elements)
    wb.save(sortie)
    afficher(elements)
    print(f'→ {sortie}')


if __name__ == '__main__':
    main()
