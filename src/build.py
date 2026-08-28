# -*- coding: utf-8 -*-
"""
Construit CAFE_VICTOR_HUGO_SYSTEME.xlsx : systeme de gestion complet
pour le Cafe Victor Hugo (Mohammedia, Maroc).

Principe : on saisit UNE SEULE FOIS dans la feuille JOURNAL.
Toutes les autres feuilles (recap jour, tableau de bord, suivi mensuel,
charges, salaires, fournisseurs) se calculent automatiquement.
"""
import os, sys, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl import Workbook
from openpyxl.utils import get_column_letter as GL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule, DataBarRule, ColorScaleRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.comments import Comment
from openpyxl.chart import BarChart, LineChart, PieChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import Marker
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.shapes import GraphicalProperties

import theme as T
import data_source as D

ANNEE, MOIS_N = 2026, 8
CLE_MOIS = ANNEE * 100 + MOIS_N            # 202608
JRN_1, JRN_N = 8, 1007                     # premiere / derniere ligne de donnees du JOURNAL
NB_LIGNES = JRN_N - JRN_1 + 1

# Plages absolues du JOURNAL, reutilisees par toutes les formules
J_DATE = f"JOURNAL!$B${JRN_1}:$B${JRN_N}"
J_CLE  = f"JOURNAL!$C${JRN_1}:$C${JRN_N}"
J_TYPE = f"JOURNAL!$D${JRN_1}:$D${JRN_N}"
J_CAT  = f"JOURNAL!$E${JRN_1}:$E${JRN_N}"
J_TIERS= f"JOURNAL!$F${JRN_1}:$F${JRN_N}"
J_MNT  = f"JOURNAL!$H${JRN_1}:$H${JRN_N}"
J_MODE = f"JOURNAL!$I${JRN_1}:$I${JRN_N}"

wb = Workbook()
wb.remove(wb.active)


def feuille(nom, couleur, masquer_grille=True, zoom=100):
    ws = wb.create_sheet(nom)
    ws.sheet_properties.tabColor = couleur
    ws.sheet_view.showGridLines = not masquer_grille
    ws.sheet_view.zoomScale = zoom
    return ws


def larg(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


# =========================================================================
# 1. PARAMETRES  -- listes de reference et reglages du commerce
# =========================================================================
p = feuille('PARAMETRES', T.GRIS)
larg(p, {'A': 2, 'B': 30, 'C': 22, 'D': 3, 'E': 26, 'F': 3,
         'G': 26, 'H': 3, 'I': 24, 'J': 14, 'K': 3, 'L': 22})

T.bandeau(p, 1, 2, 12, 'PARAMÈTRES DU COMMERCE',
          'Réglages et listes de référence. Modifiez uniquement les cellules sur fond jaune.')

T.titre_section(p, 4, 2, 3, 'IDENTITÉ')
ident = [
    ('Nom du commerce', 'CAFE VICTOR HUGO'),
    ('Ville', 'Mohammedia'),
    ('Pays', 'Maroc'),
    ('Gérant', 'AHMED'),
    ('Devise', 'DH (MAD)'),
]
r = 5
for lib, val in ident:
    p.cell(r, 2, lib).font = T.police(10, True)
    c = p.cell(r, 3, val)
    c.font = T.police(10, False, T.BLEU)
    c.fill = T.fond(T.JAUNE)
    c.border = T.BORD_LEGER
    c.alignment = T.GAUCHE
    r += 1

T.titre_section(p, 11, 2, 3, 'PÉRIODE ANALYSÉE')
p.cell(12, 2, 'Annee').font = T.police(10, True)
c = p.cell(12, 3, ANNEE); c.font = T.police(10, True, T.BLEU); c.fill = T.fond(T.JAUNE)
c.border = T.BORD_LEGER; c.alignment = T.CENTRE; c.number_format = '0'
p.cell(13, 2, 'Mois (1 a 12)').font = T.police(10, True)
c = p.cell(13, 3, MOIS_N); c.font = T.police(10, True, T.BLEU); c.fill = T.fond(T.JAUNE)
c.border = T.BORD_LEGER; c.alignment = T.CENTRE; c.number_format = '0'
c.comment = Comment("Changez ce numéro pour basculer tout le système "
                    "sur un autre mois : le RÉCAP JOUR et le TABLEAU DE BORD "
                    "suivent automatiquement.", 'Système', height=110, width=250)
p.cell(14, 2, 'Clé du mois').font = T.police(10, True)
c = p.cell(14, 3, '=$C$12*100+$C$13'); c.font = T.police(10, True, T.ENCRE)
c.border = T.BORD_LEGER; c.alignment = T.CENTRE; c.number_format = '0'
p.cell(15, 2, 'Premier jour du mois').font = T.police(10, True)
c = p.cell(15, 3, '=DATE($C$12,$C$13,1)'); c.font = T.police(10, True)
c.border = T.BORD_LEGER; c.alignment = T.CENTRE; c.number_format = T.MOIS
p.cell(16, 2, 'Nombre de jours').font = T.police(10, True)
c = p.cell(16, 3, '=DAY(EOMONTH($C$15,0))'); c.font = T.police(10, True)
c.border = T.BORD_LEGER; c.alignment = T.CENTRE; c.number_format = '0'

T.titre_section(p, 18, 2, 3, 'CAISSE')
p.cell(19, 2, 'Caisse au 1er du mois').font = T.police(10, True)
c = p.cell(19, 3, 0); c.font = T.police(10, True, T.BLEU); c.fill = T.fond(T.JAUNE)
c.border = T.BORD_LEGER; c.number_format = T.DH; c.alignment = T.DROITE
c.comment = Comment("Solde d’espèces en caisse le matin du 1er jour du mois. "
                    "Repris du classeur d’origine (feuille 01-08-2026, cellule A18 = 0).",
                    'Système', height=90, width=250)
p.cell(20, 2, 'Objectif recette / jour').font = T.police(10, True)
c = p.cell(20, 3, 2200); c.font = T.police(10, True, T.BLEU); c.fill = T.fond(T.JAUNE)
c.border = T.BORD_LEGER; c.number_format = T.DH; c.alignment = T.DROITE
p.cell(21, 2, 'Seuil alerte caisse').font = T.police(10, True)
c = p.cell(21, 3, 500); c.font = T.police(10, True, T.BLEU); c.fill = T.fond(T.JAUNE)
c.border = T.BORD_LEGER; c.number_format = T.DH; c.alignment = T.DROITE

# --- Listes de reference -------------------------------------------------
T.titre_section(p, 4, 5, 5, 'CATÉGORIES RECETTE')
for i, v in enumerate(D.CATEGORIES_RECETTE):
    c = p.cell(5 + i, 5, v)
    c.font = T.police(10); c.border = T.BORD_LEGER; c.alignment = T.GAUCHE
    c.fill = T.fond(T.VERT_CLAIR if i % 2 == 0 else T.BLANC)

T.titre_section(p, 4, 7, 7, 'CATÉGORIES DEPENSE')
for i, v in enumerate(D.CATEGORIES_DEPENSE):
    c = p.cell(5 + i, 7, v)
    c.font = T.police(10); c.border = T.BORD_LEGER; c.alignment = T.GAUCHE
    c.fill = T.fond(T.CREME if i % 2 == 0 else T.BLANC)

# Fournisseurs / beneficiaires observes dans l’historique
tiers = sorted({t for _, t, _, _ in D.DEPENSES} |
               {'MARWA', 'AHMED', 'INTERNET', 'LOYER', 'BANQUE', 'MOUHSSINE', 'HAMID',
                'HAMZA', 'LATIFA 1', 'LATIFA 2', 'CLIENT'})
T.titre_section(p, 4, 9, 10, 'FOURNISSEURS / BÉNÉFICIAIRES')
for i, v in enumerate(tiers):
    c = p.cell(5 + i, 9, v)
    c.font = T.police(10); c.border = T.BORD_LEGER; c.alignment = T.GAUCHE
    c.fill = T.fond(T.CREME if i % 2 == 0 else T.BLANC)
    cc = p.cell(5 + i, 10, D.CATEGORIE_PAR_TIERS.get(v, 'Divers'))
    cc.font = T.police(9, False, T.GRIS); cc.border = T.BORD_LEGER; cc.alignment = T.GAUCHE
NB_TIERS = len(tiers)

T.titre_section(p, 4, 12, 12, 'MODES DE PAIEMENT')
for i, v in enumerate(D.MODES):
    c = p.cell(5 + i, 12, v)
    c.font = T.police(10); c.border = T.BORD_LEGER; c.alignment = T.GAUCHE
    c.fill = T.fond(T.CREME if i % 2 == 0 else T.BLANC)

T.titre_section(p, 24, 2, 3, 'LÉGENDE DES COULEURS')
legende = [
    (T.JAUNE, T.BLEU, 'Cellule à remplir par vous'),
    (T.BLANC, T.ENCRE, 'Cellule calculée automatiquement — ne pas modifier'),
    (T.VERT_CLAIR, T.VERT, 'Résultat positif / objectif atteint'),
    (T.ROUGE_CLAIR, T.ROUGE, 'Résultat négatif / alerte'),
]
for i, (bg, fg, txt) in enumerate(legende):
    c = p.cell(25 + i, 2, '    ')
    c.fill = T.fond(bg); c.border = T.BORD_BOITE
    t = p.cell(25 + i, 3, txt)
    t.font = T.police(9, False, fg); t.alignment = T.GAUCHE

# Plages nommees utilisees par les listes deroulantes
LST_REC   = f"PARAMETRES!$E$5:$E${4+len(D.CATEGORIES_RECETTE)}"
LST_DEP   = f"PARAMETRES!$G$5:$G${4+len(D.CATEGORIES_DEPENSE)}"
LST_TIERS = f"PARAMETRES!$I$5:$I${4+NB_TIERS}"
LST_MODE  = f"PARAMETRES!$L$5:$L${4+len(D.MODES)}"
LST_CAT   = f"PARAMETRES!$N$5:$N${4+len(D.CATEGORIES_RECETTE)+len(D.CATEGORIES_DEPENSE)}"

# Colonne N : union recette + depense, pour la liste deroulante du JOURNAL
toutes_cat = D.CATEGORIES_RECETTE + D.CATEGORIES_DEPENSE
p.column_dimensions['N'].width = 26
T.titre_section(p, 4, 14, 14, 'TOUTES CATÉGORIES')
for i, v in enumerate(toutes_cat):
    c = p.cell(5 + i, 14, v)
    c.font = T.police(9); c.border = T.BORD_LEGER; c.alignment = T.GAUCHE
    c.fill = T.fond(T.VERT_CLAIR if v in D.CATEGORIES_RECETTE else T.CREME)

for nom, ref in [('Cat_Recette', LST_REC), ('Cat_Depense', LST_DEP),
                 ('Cat_Toutes', LST_CAT), ('Liste_Tiers', LST_TIERS),
                 ('Liste_Modes', LST_MODE)]:
    wb.defined_names.add(DefinedName(nom, attr_text=ref.replace('!', '!')))

p.freeze_panes = 'A4'


# =========================================================================
# 2. JOURNAL  -- unique point de saisie du systeme
# =========================================================================
j = feuille('JOURNAL', T.CARAMEL)
larg(j, {'A': 7, 'B': 13, 'C': 10, 'D': 12, 'E': 22, 'F': 20,
         'G': 34, 'H': 15, 'I': 12, 'J': 13})

T.bandeau(j, 1, 1, 10, 'JOURNAL DES OPÉRATIONS',
          "Toute la vie du café se saisit ici, une ligne par opération. "
          'Le reste du classeur se met à jour tout seul.')

mode_emploi = (
    "COMMENT SAISIR  ▸  1. Tapez la DATE  ▸  2. Choisissez RECETTE ou DEPENSE  "
    "▸  3. Choisissez la CATÉGORIE et le FOURNISSEUR dans la liste  "
    "▸  4. Tapez le MONTANT en DH.   La colonne CONTRÔLE passe au vert quand la ligne est complète."
)
j.merge_cells('A3:J3')
c = j['A3']; c.value = mode_emploi
c.font = T.police(9, True, T.ESPRESSO)
c.fill = T.fond(T.CREME_2)
c.alignment = T.Alignment(horizontal='left', vertical='center', indent=1)
c.border = T.BORD_BOITE
j.row_dimensions[3].height = 30

# Compteurs de tete
res_tete = [
    ('Lignes saisies', f'=COUNT($B${JRN_1}:$B${JRN_N})', T.NBR, T.MOKA),
    ('Total recettes', f'=SUMIF($D${JRN_1}:$D${JRN_N},"RECETTE",$H${JRN_1}:$H${JRN_N})', T.DH, T.VERT),
    ('Total dépenses', f'=SUMIF($D${JRN_1}:$D${JRN_N},"DEPENSE",$H${JRN_1}:$H${JRN_N})', T.DH, T.ROUGE),
    ('Résultat', f'=SUMIF($D${JRN_1}:$D${JRN_N},"RECETTE",$H${JRN_1}:$H${JRN_N})'
                 f'-SUMIF($D${JRN_1}:$D${JRN_N},"DEPENSE",$H${JRN_1}:$H${JRN_N})', T.DH, T.ESPRESSO),
]
for i, (lib, f, fmt, col) in enumerate(res_tete):
    cc = 1 + i * 2
    a = j.cell(5, cc, lib)
    a.font = T.police(8, True, T.GRIS); a.alignment = T.CENTRE
    j.merge_cells(start_row=5, start_column=cc, end_row=5, end_column=cc + 1)
    b = j.cell(6, cc, f)
    b.font = T.police(13, True, col); b.number_format = fmt; b.alignment = T.CENTRE
    b.fill = T.fond(T.CREME)
    j.merge_cells(start_row=6, start_column=cc, end_row=6, end_column=cc + 1)
    for k in (cc, cc + 1):
        j.cell(5, k).fill = T.fond(T.CREME)
        j.cell(6, k).fill = T.fond(T.CREME)
        j.cell(6, k).border = T.Border(bottom=T._s(T.CARAMEL, 'medium'))
j.row_dimensions[5].height = 16
j.row_dimensions[6].height = 26

T.entetes(j, JRN_1 - 1, 1,
          ['N°', 'DATE', 'CLE MOIS', 'TYPE', 'CATÉGORIE', 'FOURNISSEUR /\nBÉNÉFICIAIRE',
           'DESCRIPTION', 'MONTANT (DH)', 'MODE', 'CONTROLE'])

# --- Construction du flux d’operations a transferer ----------------------
BLOC_LIB = {'A': 'Achat du jour', 'F': 'Charge fixe', 'S': 'Salaire / avance',
            'V': 'Virement'}
operations = []
for jour in sorted(D.RECETTES):
    d = datetime.date(ANNEE, MOIS_N, jour)
    operations.append((d, 'RECETTE', 'Vente Caisse', 'CLIENT',
                       'Recette caisse du jour', D.RECETTES[jour], 'Espèce'))
    for jj, tiers_, montant, bloc in D.DEPENSES:
        if jj != jour:
            continue
        cat = D.CATEGORIE_PAR_TIERS.get(tiers_, 'Divers')
        mode = 'Virement' if bloc == 'V' and tiers_ != 'AVANCE LOYER' else 'Espèce'
        operations.append((d, 'DEPENSE', cat, tiers_, BLOC_LIB[bloc], montant, mode))

for i in range(NB_LIGNES):
    r = JRN_1 + i
    fond_l = T.CREME if i % 2 else T.BLANC
    valeurs = operations[i] if i < len(operations) else None

    a = j.cell(r, 1, f'=IF($B{r}="","",ROW()-{JRN_1 - 1})')
    a.font = T.police(9, False, T.GRIS); a.alignment = T.CENTRE

    b = j.cell(r, 2, valeurs[0] if valeurs else None)
    b.number_format = T.DATE; b.alignment = T.CENTRE
    b.font = T.police(10, False, T.BLEU)

    c = j.cell(r, 3, f'=IF($B{r}="","",YEAR($B{r})*100+MONTH($B{r}))')
    c.number_format = '0'; c.alignment = T.CENTRE
    c.font = T.police(9, False, T.GRIS)

    d_ = j.cell(r, 4, valeurs[1] if valeurs else None)
    d_.alignment = T.CENTRE; d_.font = T.police(9, True, T.BLEU)

    e = j.cell(r, 5, valeurs[2] if valeurs else None)
    e.alignment = T.GAUCHE; e.font = T.police(10, False, T.BLEU)

    f = j.cell(r, 6, valeurs[3] if valeurs else None)
    f.alignment = T.GAUCHE; f.font = T.police(10, False, T.BLEU)

    g = j.cell(r, 7, valeurs[4] if valeurs else None)
    g.alignment = T.GAUCHE; g.font = T.police(9, False, T.GRIS)

    h = j.cell(r, 8, valeurs[5] if valeurs else None)
    h.number_format = T.DH; h.alignment = T.DROITE
    h.font = T.police(10, True, T.BLEU)

    i_ = j.cell(r, 9, valeurs[6] if valeurs else None)
    i_.alignment = T.CENTRE; i_.font = T.police(9, False, T.BLEU)

    k = j.cell(r, 10,
               f'=IF($B{r}="","",'
               f'IF(N($H{r})<=0,"Montant ?",'
               f'IF($D{r}="RECETTE",IF(COUNTIF({LST_REC},$E{r})=0,"Catégorie ?","OK"),'
               f'IF($D{r}="DEPENSE",IF(COUNTIF({LST_DEP},$E{r})=0,"Catégorie ?","OK"),'
               f'"Type ?"))))')
    k.alignment = T.CENTRE; k.font = T.police(9, True)

    for col in range(1, 11):
        cell = j.cell(r, col)
        cell.border = T.BORD_LEGER
        if cell.fill.fgColor.rgb in (None, '00000000'):
            cell.fill = T.fond(fond_l)
    j.row_dimensions[r].height = 17

j.cell(JRN_1 - 1, 1).comment = Comment(
    "Numérotation automatique : elle se remplit dès qu’une date est saisie.",
    'Système', height=70, width=220)

# --- Listes deroulantes ---------------------------------------------------
dv_specs = [
    (f'"RECETTE,DEPENSE"', f'D{JRN_1}:D{JRN_N}', 'Type', 'Choisissez RECETTE ou DEPENSE.'),
    (f'={LST_CAT}',   f'E{JRN_1}:E{JRN_N}', 'Categorie',
     'Choisissez une catégorie de la feuille PARAMÈTRES.'),
    (f'={LST_TIERS}', f'F{JRN_1}:F{JRN_N}', 'Fournisseur',
     'Choisissez un fournisseur existant, ou ajoutez-le dans PARAMÈTRES.'),
    (f'={LST_MODE}',  f'I{JRN_1}:I{JRN_N}', 'Mode de paiement', 'Espece, Carte, Virement ou Cheque.'),
]
for formule, plage, titre, msg in dv_specs:
    dv = DataValidation(type='list', formula1=formule, allow_blank=True,
                        showDropDown=False, showErrorMessage=False)
    dv.promptTitle = titre; dv.prompt = msg; dv.showInputMessage = True
    j.add_data_validation(dv); dv.add(plage)

dv_mnt = DataValidation(type='decimal', operator='greaterThan', formula1='0',
                        allow_blank=True, showErrorMessage=True)
dv_mnt.errorTitle = 'Montant invalide'
dv_mnt.error = 'Le montant doit être un nombre positif, en dirhams.'
j.add_data_validation(dv_mnt); dv_mnt.add(f'H{JRN_1}:H{JRN_N}')

# --- Mise en forme conditionnelle ----------------------------------------
plage_all = f'A{JRN_1}:J{JRN_N}'
j.conditional_formatting.add(
    plage_all, FormulaRule(formula=[f'$D{JRN_1}="RECETTE"'],
                           fill=T.fond(T.VERT_CLAIR), stopIfTrue=False))
j.conditional_formatting.add(
    f'J{JRN_1}:J{JRN_N}',
    CellIsRule(operator='equal', formula=['"OK"'], font=T.police(9, True, T.VERT),
               fill=T.fond(T.VERT_CLAIR)))
j.conditional_formatting.add(
    f'J{JRN_1}:J{JRN_N}',
    FormulaRule(formula=[f'AND($J{JRN_1}<>"",$J{JRN_1}<>"OK")'],
                font=T.police(9, True, T.ROUGE), fill=T.fond(T.ROUGE_CLAIR)))
j.conditional_formatting.add(
    f'H{JRN_1}:H{JRN_N}',
    DataBarRule(start_type='num', start_value=0, end_type='percentile',
                end_value=95, color=T.CARAMEL, showValue=True))

j.auto_filter.ref = f'A{JRN_1 - 1}:J{JRN_N}'
j.print_area = f'A1:J{JRN_1 + 199}'
j.freeze_panes = f'A{JRN_1}'
NB_OPS = len(operations)


# =========================================================================
# 3. RECAP JOUR  -- une ligne par jour du mois, entierement calculee
# =========================================================================
GROUPES = {
    'achats':   ['Boissons', 'Café & Torréfaction', 'Boulangerie',
                 'Fruits & Légumes', 'Épicerie & Divers'],
    'nettoyage': ['Nettoyage'],
    'salaires': ['Salaires'],
    'charges':  ['Loyer', 'Eau & Électricité', 'Internet & Télécom',
                 'Taxes & Impôts', 'Abonnements', 'Entretien & Réparation'],
    'virements': ['Virement Banque', 'Virement Associés'],
}


def somme_cats(cats, crit_plage, crit):
    """Somme du JOURNAL pour une liste de categories et un critere supplementaire."""
    return '+'.join(f'SUMIFS({J_MNT},{J_CAT},"{c}",{crit_plage},{crit})' for c in cats)


rj = feuille('RECAP JOUR', T.MOKA)
larg(rj, {'A': 12, 'B': 11, 'C': 13, 'D': 12, 'E': 12, 'F': 14, 'G': 13, 'H': 12,
          'I': 12, 'J': 13, 'K': 12, 'L': 11, 'M': 14, 'N': 14, 'O': 15, 'P': 11})

T.bandeau(rj, 1, 1, 16, 'RÉCAP JOURNALIER DU MOIS',
          'Calculé automatiquement depuis le JOURNAL. Aucune saisie ici. '
          'Changez le mois dans PARAMÈTRES pour changer la période.')

rj.merge_cells('A3:F3')
c = rj['A3']; c.value = '=PARAMETRES!$C$15'
c.font = T.police(14, True, T.ESPRESSO); c.number_format = T.MOIS
c.alignment = T.Alignment(horizontal='left', vertical='center', indent=1)
rj.row_dimensions[3].height = 26

RJ_H = 5           # ligne d’en-tetes
RJ_1 = 6           # premiere ligne de donnees
RJ_N = RJ_1 + 30   # 31 jours

T.entetes(rj, RJ_H, 1,
          ['DATE', 'JOUR', 'RECETTE\nCAISSE', 'GLOVO\nESPECE', 'GLOVO\nCARTE',
           'TOTAL\nRECETTE', 'ACHATS', 'NETTOYAGE', 'SALAIRES', 'CHARGES\nFIXES',
           'VIREMENTS', 'AUTRES', 'TOTAL\nDÉPENSES', 'RÉSULTAT\nDU JOUR',
           'CAISSE\nCUMULÉE', 'vs OBJ.'])

for i in range(31):
    r, jour = RJ_1 + i, i + 1
    A = f'$A{r}'
    rj.cell(r, 1, f'=IF({jour}>PARAMETRES!$C$16,"",DATE(PARAMETRES!$C$12,PARAMETRES!$C$13,{jour}))')
    rj.cell(r, 2, f'=IF({A}="","",CHOOSE(WEEKDAY({A},2),"Lundi","Mardi","Mercredi",'
                  f'"Jeudi","Vendredi","Samedi","Dimanche"))')
    rj.cell(r, 3, f'=IF({A}="","",SUMIFS({J_MNT},{J_CAT},"Vente Caisse",{J_DATE},{A})'
                  f'+SUMIFS({J_MNT},{J_CAT},"Terrasse",{J_DATE},{A})'
                  f'+SUMIFS({J_MNT},{J_CAT},"Autre Recette",{J_DATE},{A}))')
    rj.cell(r, 4, f'=IF({A}="","",SUMIFS({J_MNT},{J_CAT},"Glovo Espèce",{J_DATE},{A}))')
    rj.cell(r, 5, f'=IF({A}="","",SUMIFS({J_MNT},{J_CAT},"Glovo Carte",{J_DATE},{A}))')
    rj.cell(r, 6, f'=IF({A}="","",SUM($C{r}:$E{r}))')
    rj.cell(r, 7, f'=IF({A}="","",{somme_cats(GROUPES["achats"], J_DATE, A)})')
    rj.cell(r, 8, f'=IF({A}="","",{somme_cats(GROUPES["nettoyage"], J_DATE, A)})')
    rj.cell(r, 9, f'=IF({A}="","",{somme_cats(GROUPES["salaires"], J_DATE, A)})')
    rj.cell(r, 10, f'=IF({A}="","",{somme_cats(GROUPES["charges"], J_DATE, A)})')
    rj.cell(r, 11, f'=IF({A}="","",{somme_cats(GROUPES["virements"], J_DATE, A)})')
    rj.cell(r, 12, f'=IF({A}="","",$M{r}-SUM($G{r}:$K{r}))')
    rj.cell(r, 13, f'=IF({A}="","",SUMIFS({J_MNT},{J_TYPE},"DEPENSE",{J_DATE},{A}))')
    rj.cell(r, 14, f'=IF({A}="","",$F{r}-$M{r})')
    if i == 0:
        rj.cell(r, 15, f'=IF({A}="","",PARAMETRES!$C$19+$N{r})')
    else:
        rj.cell(r, 15, f'=IF({A}="","",$O{r-1}+$N{r})')
    rj.cell(r, 16, f'=IF({A}="","",IF($F{r}=0,"—",$F{r}-PARAMETRES!$C$20))')

    fond_l = T.CREME if i % 2 else T.BLANC
    for col in range(1, 17):
        cell = rj.cell(r, col)
        cell.border = T.BORD_LEGER
        cell.fill = T.fond(fond_l)
        cell.font = T.police(10)
        if col == 1:
            cell.number_format = T.DATE; cell.alignment = T.CENTRE
            cell.font = T.police(10, True, T.ESPRESSO)
        elif col == 2:
            cell.alignment = T.CENTRE; cell.font = T.police(9, False, T.GRIS)
        else:
            cell.number_format = T.DH; cell.alignment = T.DROITE
    rj.cell(r, 6).font = T.police(10, True, T.VERT)
    rj.cell(r, 13).font = T.police(10, True, T.ROUGE)
    rj.cell(r, 14).font = T.police(10, True)
    rj.cell(r, 15).font = T.police(10, True, T.ESPRESSO)
    rj.row_dimensions[r].height = 18

# Ligne de totaux
RJ_T = RJ_N + 1
rj.cell(RJ_T, 1, 'TOTAL DU MOIS')
rj.merge_cells(start_row=RJ_T, start_column=1, end_row=RJ_T, end_column=2)
for col in range(3, 15):
    L = GL(col)
    cell = rj.cell(RJ_T, col, f'=SUM({L}{RJ_1}:{L}{RJ_N})')
    cell.number_format = T.DH; cell.alignment = T.DROITE
rj.cell(RJ_T, 15, f'=$O{RJ_N}')
rj.cell(RJ_T, 15).number_format = T.DH; rj.cell(RJ_T, 15).alignment = T.DROITE
rj.cell(RJ_T, 16, f'=SUMIF($F${RJ_1}:$F${RJ_N},">0",$P${RJ_1}:$P${RJ_N})')
rj.cell(RJ_T, 16).number_format = T.DH; rj.cell(RJ_T, 16).alignment = T.DROITE
T.ligne_total(rj, RJ_T, 1, 16)
rj.cell(RJ_T, 1).alignment = T.Alignment(horizontal='left', vertical='center', indent=1)

# Ligne ratio
RJ_R = RJ_T + 1
rj.cell(RJ_R, 1, 'RATIO DÉPENSE / RECETTE')
rj.merge_cells(start_row=RJ_R, start_column=1, end_row=RJ_R, end_column=2)
for col, num in [(7, 'G'), (8, 'H'), (9, 'I'), (10, 'J'), (11, 'K'), (13, 'M')]:
    cell = rj.cell(RJ_R, col, f'=IFERROR({num}{RJ_T}/$F${RJ_T},0)')
    cell.number_format = T.PCT; cell.alignment = T.DROITE
    cell.font = T.police(9, True, T.AMBRE); cell.fill = T.fond(T.CREME_2)
for col in range(1, 17):
    cl = rj.cell(RJ_R, col)
    if cl.fill.fgColor.rgb in (None, '00000000'):
        cl.fill = T.fond(T.CREME_2)
    cl.border = T.Border(bottom=T._s(T.CARAMEL, 'medium'))
rj.cell(RJ_R, 1).font = T.police(9, True, T.ESPRESSO)
rj.cell(RJ_R, 1).alignment = T.Alignment(horizontal='left', vertical='center', indent=1)
rj.row_dimensions[RJ_R].height = 20

# Couleurs conditionnelles
rj.conditional_formatting.add(
    f'N{RJ_1}:N{RJ_N}',
    CellIsRule(operator='lessThan', formula=['0'], font=T.police(10, True, T.ROUGE),
               fill=T.fond(T.ROUGE_CLAIR)))
rj.conditional_formatting.add(
    f'N{RJ_1}:N{RJ_N}',
    CellIsRule(operator='greaterThan', formula=['0'], font=T.police(10, True, T.VERT)))
rj.conditional_formatting.add(
    f'O{RJ_1}:O{RJ_N}',
    CellIsRule(operator='lessThan', formula=['PARAMETRES!$C$21'],
               font=T.police(10, True, T.ROUGE), fill=T.fond(T.ROUGE_CLAIR)))
rj.conditional_formatting.add(
    f'P{RJ_1}:P{RJ_N}',
    CellIsRule(operator='greaterThanOrEqual', formula=['0'], font=T.police(10, True, T.VERT)))
rj.conditional_formatting.add(
    f'P{RJ_1}:P{RJ_N}',
    CellIsRule(operator='lessThan', formula=['0'], font=T.police(10, True, T.ROUGE)))
rj.conditional_formatting.add(
    f'F{RJ_1}:F{RJ_N}',
    DataBarRule(start_type='num', start_value=0, end_type='percentile', end_value=98,
                color='9CCC9C', showValue=True))
rj.conditional_formatting.add(
    f'A{RJ_1}:P{RJ_N}',
    FormulaRule(formula=[f'AND($A{RJ_1}<>"",WEEKDAY($A{RJ_1},2)>5)'],
                fill=T.fond('F2EFE6'), stopIfTrue=False))

rj.freeze_panes = f'C{RJ_1}'
T.note(rj, RJ_R + 2, 1, 16,
       "Lecture : « AUTRES » regroupe les dépenses dont la catégorie ne rentre dans aucun groupe ci-dessus ; "
       "si cette colonne n’est pas vide, vérifiez la catégorie des lignes concernées dans le JOURNAL. "
       "« CAISSE CUMULÉE » part du solde saisi dans PARAMÈTRES (caisse au 1er du mois) et ajoute chaque "
       "jour le résultat. Les lignes grisées sont les samedis et dimanches.", 44)


# =========================================================================
# 4. TABLEAU DE BORD
# =========================================================================
tb = feuille('TABLEAU DE BORD', T.ESPRESSO)
for col in 'ABCDEFGHIJKLMNOP':
    tb.column_dimensions[col].width = 11.5
tb.column_dimensions['A'].width = 13

T.bandeau(tb, 1, 1, 16, 'TABLEAU DE BORD',
          'Photo du mois en cours. Tout se met à jour dès que vous saisissez dans le JOURNAL.')

tb.merge_cells('A4:F4')
c = tb['A4']; c.value = '=PARAMETRES!$C$15'
c.font = T.police(15, True, T.ESPRESSO); c.number_format = T.MOIS
c.alignment = T.Alignment(horizontal='left', vertical='center', indent=1)
tb.merge_cells('K4:P4')
c = tb['K4']; c.value = '=PARAMETRES!$C$5&" — "&PARAMETRES!$C$6'
c.font = T.police(11, True, T.CARAMEL); c.alignment = T.DROITE
tb.row_dimensions[4].height = 26

CLE = 'PARAMETRES!$C$14'
REC  = f'SUMIFS({J_MNT},{J_TYPE},"RECETTE",{J_CLE},{CLE})'
DEP  = f'SUMIFS({J_MNT},{J_TYPE},"DEPENSE",{J_CLE},{CLE})'
NBJ  = f"COUNTIF('RECAP JOUR'!$F${RJ_1}:$F${RJ_N},\">0\")"


def carte(ligne, col, libelle, formule, fmt, couleur, sous):
    """Carte d’indicateur : libelle, grande valeur, commentaire."""
    tb.merge_cells(start_row=ligne, start_column=col, end_row=ligne, end_column=col + 2)
    a = tb.cell(ligne, col, libelle)
    a.font = T.police(8.5, True, T.GRIS)
    a.alignment = T.Alignment(horizontal='left', vertical='center', indent=1)

    tb.merge_cells(start_row=ligne + 1, start_column=col, end_row=ligne + 1, end_column=col + 2)
    b = tb.cell(ligne + 1, col, formule)
    b.font = T.police(18, True, couleur); b.number_format = fmt
    b.alignment = T.Alignment(horizontal='left', vertical='center', indent=1)

    tb.merge_cells(start_row=ligne + 2, start_column=col, end_row=ligne + 2, end_column=col + 2)
    d = tb.cell(ligne + 2, col, sous)
    d.font = T.police(8, False, T.GRIS)
    d.alignment = T.Alignment(horizontal='left', vertical='top', indent=1)

    haut = T.Border(top=T._s(couleur, 'thick'), left=T._s(T.SABLE), right=T._s(T.SABLE))
    mil = T.Border(left=T._s(T.SABLE), right=T._s(T.SABLE))
    bas = T.Border(left=T._s(T.SABLE), right=T._s(T.SABLE), bottom=T._s(T.SABLE))
    for k in range(col, col + 3):
        tb.cell(ligne, k).fill = T.fond(T.CREME); tb.cell(ligne, k).border = haut
        tb.cell(ligne + 1, k).fill = T.fond(T.CREME); tb.cell(ligne + 1, k).border = mil
        tb.cell(ligne + 2, k).fill = T.fond(T.CREME); tb.cell(ligne + 2, k).border = bas
    tb.row_dimensions[ligne].height = 16
    tb.row_dimensions[ligne + 1].height = 28
    tb.row_dimensions[ligne + 2].height = 15


cartes = [
    (6, 1,  'RECETTE DU MOIS',      f'={REC}', T.DH, T.VERT,     'Total encaissé ce mois'),
    (6, 5,  'DÉPENSES DU MOIS',     f'={DEP}', T.DH, T.ROUGE,    'Achats, salaires, charges, virements'),
    (6, 9,  'RÉSULTAT NET',         f'={REC}-{DEP}', T.DH, T.ESPRESSO, 'Recette moins dépenses'),
    (6, 13, 'MARGE NETTE',          f'=IFERROR(({REC}-{DEP})/{REC},0)', T.PCT, T.AMBRE,
     'Part de la recette qui reste'),
    (10, 1, 'RECETTE MOYENNE / JOUR', f'=IFERROR({REC}/{NBJ},0)', T.DH, T.MOKA,
     'Sur les jours réellement travaillés'),
    (10, 5, 'MEILLEURE JOURNÉE',
     f"=IFERROR(MAX('RECAP JOUR'!$F${RJ_1}:$F${RJ_N}),0)", T.DH, T.VERT, 'Recette la plus haute du mois'),
    (10, 9, 'CAISSE EN FIN DE MOIS',
     f"='RECAP JOUR'!$O${RJ_T}", T.DH, T.ESPRESSO, 'Solde théorique d\'espèces'),
    (10, 13, 'JOURS SAISIS',        f'={NBJ}', T.NBR, T.MOKA, 'Jours avec une recette enregistrée'),
]
for args in cartes:
    carte(*args)

tb.conditional_formatting.add('I7:K7', CellIsRule(
    operator='lessThan', formula=['0'], font=T.police(18, True, T.ROUGE)))
tb.conditional_formatting.add('M7:O7', CellIsRule(
    operator='lessThan', formula=['0'], font=T.police(18, True, T.ROUGE)))

# --- Tableaux d’analyse (alimentent les graphiques) ----------------------
CAT_1 = 60
T.titre_section(tb, CAT_1 - 2, 1, 4, 'DÉPENSES PAR CATÉGORIE')
T.entetes(tb, CAT_1 - 1, 1, ['CATÉGORIE', 'MONTANT', 'PART', 'PAR JOUR'],
          [22, 13, 10, 12])
for i, cat in enumerate(D.CATEGORIES_DEPENSE):
    r = CAT_1 + i
    tb.cell(r, 1, cat).font = T.police(9)
    tb.cell(r, 2, f'=SUMIFS({J_MNT},{J_CAT},"{cat}",{J_CLE},{CLE})')
    tb.cell(r, 3, f'=IFERROR($B{r}/${GL(2)}${CAT_1 + len(D.CATEGORIES_DEPENSE)},0)')
    tb.cell(r, 4, f'=IFERROR($B{r}/{NBJ},0)')
    for col in range(1, 5):
        cell = tb.cell(r, col); cell.border = T.BORD_LEGER
        cell.fill = T.fond(T.CREME if i % 2 else T.BLANC)
    tb.cell(r, 1).alignment = T.GAUCHE
    tb.cell(r, 2).number_format = T.DH; tb.cell(r, 2).alignment = T.DROITE
    tb.cell(r, 2).font = T.police(9, True)
    tb.cell(r, 3).number_format = T.PCT; tb.cell(r, 3).alignment = T.DROITE
    tb.cell(r, 3).font = T.police(9, False, T.AMBRE)
    tb.cell(r, 4).number_format = T.DH; tb.cell(r, 4).alignment = T.DROITE
    tb.cell(r, 4).font = T.police(9, False, T.GRIS)
CAT_N = CAT_1 + len(D.CATEGORIES_DEPENSE) - 1
rt = CAT_N + 1
tb.cell(rt, 1, 'TOTAL DÉPENSES')
tb.cell(rt, 2, f'=SUM($B${CAT_1}:$B${CAT_N})').number_format = T.DH
tb.cell(rt, 3, f'=IFERROR($B{rt}/$B{rt},0)').number_format = T.PCT
tb.cell(rt, 4, f'=IFERROR($B{rt}/{NBJ},0)').number_format = T.DH
T.ligne_total(tb, rt, 1, 4)
for col in (2, 3, 4):
    tb.cell(rt, col).alignment = T.DROITE
tb.cell(rt, 1).alignment = T.Alignment(horizontal='left', vertical='center', indent=1)
tb.conditional_formatting.add(
    f'B{CAT_1}:B{CAT_N}',
    DataBarRule(start_type='num', start_value=0, end_type='percentile', end_value=100,
                color=T.CARAMEL, showValue=True))

# Principaux fournisseurs, classes sur l’historique transfere
poids = {}
for _, t_, m_, bloc in D.DEPENSES:
    if bloc in ('A', 'F'):
        poids[t_] = poids.get(t_, 0) + m_
top_frs = [t for t, _ in sorted(poids.items(), key=lambda x: -x[1])][:12]

FRS_1 = CAT_1
T.titre_section(tb, FRS_1 - 2, 6, 9, 'PRINCIPAUX FOURNISSEURS')
T.entetes(tb, FRS_1 - 1, 6, ['FOURNISSEUR', 'MONTANT', 'PART', 'NB ACHATS'],
          [22, 13, 10, 12])
for i, frs in enumerate(top_frs):
    r = FRS_1 + i
    tb.cell(r, 6, frs).font = T.police(9)
    tb.cell(r, 7, f'=SUMIFS({J_MNT},{J_TIERS},"{frs}",{J_CLE},{CLE})')
    tb.cell(r, 8, f'=IFERROR($G{r}/$B${rt},0)')
    tb.cell(r, 9, f'=COUNTIFS({J_TIERS},"{frs}",{J_CLE},{CLE})')
    for col in range(6, 10):
        cell = tb.cell(r, col); cell.border = T.BORD_LEGER
        cell.fill = T.fond(T.CREME if i % 2 else T.BLANC)
    tb.cell(r, 6).alignment = T.GAUCHE
    tb.cell(r, 7).number_format = T.DH; tb.cell(r, 7).alignment = T.DROITE
    tb.cell(r, 7).font = T.police(9, True)
    tb.cell(r, 8).number_format = T.PCT; tb.cell(r, 8).alignment = T.DROITE
    tb.cell(r, 8).font = T.police(9, False, T.AMBRE)
    tb.cell(r, 9).number_format = T.NBR; tb.cell(r, 9).alignment = T.DROITE
    tb.cell(r, 9).font = T.police(9, False, T.GRIS)
FRS_N = FRS_1 + len(top_frs) - 1
tb.conditional_formatting.add(
    f'G{FRS_1}:G{FRS_N}',
    DataBarRule(start_type='num', start_value=0, end_type='percentile', end_value=100,
                color='8D6E63', showValue=True))

T.note(tb, rt + 3, 1, 16,
       "Les deux tableaux ci-dessus alimentent les graphiques et se recalculent sur le mois choisi dans "
       "PARAMÈTRES. L’ordre des fournisseurs a été figé sur le poids constaté en août 2026 (données reprises "
       "de l’ancien classeur) ; les montants, eux, suivent toujours le mois en cours.", 40)


# =========================================================================
# 5. GRAPHIQUES DU TABLEAU DE BORD
# =========================================================================
PAL = ['2E7D32', 'B3261E', 'C8A265', '5D4037', '1A56A0', 'B26A00',
       '7A9E7E', '8D6E63', 'A67C52', '4E6E81', 'C0873F', '6D4C41',
       '9E9D24', '00695C', '8E24AA', '546E7A']


def habiller(ch, titre, hauteur=8.6, largeur=17.5, legende='b'):
    ch.title = titre
    ch.height, ch.width = hauteur, largeur
    ch.style = 2
    if legende:
        ch.legend.position = legende
        ch.legend.overlay = False
    else:
        ch.legend = None
    if getattr(ch, 'y_axis', None) is not None:
        ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill='E8DCC8'))
        ch.y_axis.numFmt = '#,##0'
        ch.y_axis.delete = False
    if getattr(ch, 'x_axis', None) is not None:
        ch.x_axis.delete = False
    return ch


def couleur_serie(serie, code, ligne_seule=False, largeur_pt=22000):
    gp = GraphicalProperties(solidFill=code)
    if ligne_seule:
        gp = GraphicalProperties(ln=LineProperties(solidFill=code, w=largeur_pt))
    serie.graphicalProperties = gp


# -- 1. Recette vs depenses, jour par jour --------------------------------
g1 = BarChart(); g1.type = 'col'; g1.grouping = 'clustered'; g1.gapWidth = 40
donnees = Reference(rj, min_col=6, max_col=6, min_row=RJ_H, max_row=RJ_N)
g1.add_data(donnees, titles_from_data=True)
donnees2 = Reference(rj, min_col=13, max_col=13, min_row=RJ_H, max_row=RJ_N)
g1.add_data(donnees2, titles_from_data=True)
g1.set_categories(Reference(rj, min_col=1, max_col=1, min_row=RJ_1, max_row=RJ_N))
couleur_serie(g1.series[0], '2E7D32')
couleur_serie(g1.series[1], 'B3261E')
habiller(g1, 'Recette et dépenses, jour par jour (DH)', 8.8, 18.5)
g1.y_axis.title = 'Dirhams'
tb.add_chart(g1, 'A16')

# -- 2. Caisse cumulee ----------------------------------------------------
g2 = LineChart()
g2.add_data(Reference(rj, min_col=15, max_col=15, min_row=RJ_H, max_row=RJ_N),
            titles_from_data=True)
g2.set_categories(Reference(rj, min_col=1, max_col=1, min_row=RJ_1, max_row=RJ_N))
couleur_serie(g2.series[0], '3E2723', ligne_seule=True, largeur_pt=28000)
g2.series[0].marker = Marker(symbol='circle', size=5)
g2.series[0].smooth = False
habiller(g2, 'Caisse cumulée au fil du mois (DH)', 8.8, 15.5, legende=None)
tb.add_chart(g2, 'J16')

# -- 3. Repartition des depenses -----------------------------------------
from openpyxl.chart.series import DataPoint
g3 = BarChart(); g3.type = 'bar'; g3.gapWidth = 35
g3.add_data(Reference(tb, min_col=2, min_row=CAT_1 - 1, max_row=CAT_N), titles_from_data=True)
g3.set_categories(Reference(tb, min_col=1, min_row=CAT_1, max_row=CAT_N))
pts = []
for i in range(len(D.CATEGORIES_DEPENSE)):
    dp = DataPoint(idx=i)
    dp.graphicalProperties = GraphicalProperties(solidFill=PAL[i % len(PAL)])
    pts.append(dp)
g3.series[0].data_points = pts
habiller(g3, "Où part l’argent : dépenses par catégorie (DH)", 10.5, 17, legende=None)
tb.add_chart(g3, 'A35')

# -- 4. Principaux fournisseurs ------------------------------------------
g4 = BarChart(); g4.type = 'bar'; g4.gapWidth = 45
g4.add_data(Reference(tb, min_col=7, min_row=FRS_1 - 1, max_row=FRS_N), titles_from_data=True)
g4.set_categories(Reference(tb, min_col=6, min_row=FRS_1, max_row=FRS_N))
couleur_serie(g4.series[0], '8D6E63')
habiller(g4, 'Principaux fournisseurs du mois (DH)', 10.5, 17.5, legende=None)
tb.add_chart(g4, 'J35')

for lg, txt in [(15, 'ÉVOLUTION DU MOIS'), (34, 'ANALYSE DES DÉPENSES')]:
    T.titre_section(tb, lg, 1, 16, txt)
tb.row_dimensions[15].height = 24
tb.row_dimensions[34].height = 24


# =========================================================================
# 6. SUIVI MENSUEL  -- historique 2025 + annee en cours
# =========================================================================
sm = feuille('SUIVI MENSUEL', T.MOKA)
larg(sm, {'A': 16, 'B': 15, 'C': 14, 'D': 11, 'E': 14, 'F': 14, 'G': 13,
          'H': 15, 'I': 15, 'J': 11, 'K': 14})

T.bandeau(sm, 1, 1, 11, 'SUIVI MENSUEL',
          "Comparaison mois par mois. 2025 = historique repris de l’ancien classeur, "
          '2026 = calculé automatiquement depuis le JOURNAL.')

SM_H, SM_1 = 4, 5
T.entetes(sm, SM_H, 1,
          ['MOIS', 'RECETTE', 'ACHATS\nMARCHANDISE', '% ACHATS', 'SALAIRES',
           'CHARGES\nFIXES', 'VIREMENTS', 'TOTAL\nDÉPENSES', 'RESULTAT',
           'MARGE', 'RECETTE\nMOY./JOUR'])

lignes_sm = []
for m, rec, dep in D.HISTO_2025:
    lignes_sm.append((2025, m, rec, dep))
for m in range(1, 13):
    lignes_sm.append((2026, m, None, None))

for i, (an, m, rec, achat) in enumerate(lignes_sm):
    r = SM_1 + i
    cle = an * 100 + m
    hist = (an == 2025)
    sm.cell(r, 1, datetime.date(an, m, 1)).number_format = T.MOIS
    if hist:
        sm.cell(r, 2, rec)
        sm.cell(r, 3, achat)
        for col in (5, 6, 7, 8, 9):
            sm.cell(r, col, 'n.d.')
        sm.cell(r, 10, 'n.d.')
    else:
        sm.cell(r, 2, f'=SUMIFS({J_MNT},{J_TYPE},"RECETTE",{J_CLE},{cle})')
        sm.cell(r, 3, f'={somme_cats(GROUPES["achats"], J_CLE, cle)}'
                      f'+{somme_cats(GROUPES["nettoyage"], J_CLE, cle)}')
        sm.cell(r, 5, f'={somme_cats(GROUPES["salaires"], J_CLE, cle)}')
        sm.cell(r, 6, f'={somme_cats(GROUPES["charges"], J_CLE, cle)}')
        sm.cell(r, 7, f'={somme_cats(GROUPES["virements"], J_CLE, cle)}')
        sm.cell(r, 8, f'=SUMIFS({J_MNT},{J_TYPE},"DEPENSE",{J_CLE},{cle})')
        sm.cell(r, 9, f'=$B{r}-$H{r}')
        sm.cell(r, 10, f'=IFERROR($I{r}/$B{r},0)')
    sm.cell(r, 4, f'=IFERROR($C{r}/$B{r},0)')
    sm.cell(r, 11, f'=IFERROR($B{r}/DAY(EOMONTH($A{r},0)),0)')

    fond_l = T.CREME_2 if hist else (T.CREME if i % 2 else T.BLANC)
    for col in range(1, 12):
        cell = sm.cell(r, col)
        cell.border = T.BORD_LEGER; cell.fill = T.fond(fond_l)
        cell.font = T.police(10, False, T.GRIS if hist else T.ENCRE)
        cell.number_format = T.DH; cell.alignment = T.DROITE
    sm.cell(r, 1).number_format = T.MOIS
    sm.cell(r, 1).alignment = T.GAUCHE
    sm.cell(r, 1).font = T.police(10, True, T.ESPRESSO)
    sm.cell(r, 4).number_format = T.PCT
    sm.cell(r, 4).font = T.police(10, False, T.AMBRE)
    sm.cell(r, 10).number_format = T.PCT
    if hist:
        sm.cell(r, 2).font = T.police(10, True, T.BLEU)
        sm.cell(r, 3).font = T.police(10, True, T.BLEU)
        for col in (5, 6, 7, 8, 9, 10):
            sm.cell(r, col).alignment = T.CENTRE
            sm.cell(r, col).font = T.police(8, True, 'BDB2A5')
            sm.cell(r, col).number_format = 'General'
    else:
        sm.cell(r, 2).font = T.police(10, True, T.VERT)
        sm.cell(r, 8).font = T.police(10, True, T.ROUGE)
        sm.cell(r, 9).font = T.police(10, True)
    sm.row_dimensions[r].height = 18

SM_N = SM_1 + len(lignes_sm) - 1
SM_2025_N = SM_1 + 11
SM_2026_1 = SM_1 + 12

# Sous-totaux par annee
for lib, deb, fin, dest in [('TOTAL 2025', SM_1, SM_2025_N, SM_N + 1),
                            ('TOTAL 2026', SM_2026_1, SM_N, SM_N + 2)]:
    sm.cell(dest, 1, lib)
    for col in (2, 3, 5, 6, 7, 8, 9):
        L = GL(col)
        if lib == 'TOTAL 2025' and col in (5, 6, 7, 8, 9):
            sm.cell(dest, col, 'n.d.').alignment = T.CENTRE
            continue
        sm.cell(dest, col, f'=SUM({L}{deb}:{L}{fin})')
    sm.cell(dest, 4, f'=IFERROR($C{dest}/$B{dest},0)')
    sm.cell(dest, 10, f'=IFERROR($I{dest}/$B{dest},0)' if lib == 'TOTAL 2026' else 'n.d.')
    sm.cell(dest, 11, f'=IFERROR($B{dest}/SUM($K{deb}:$K{fin})*AVERAGE($K{deb}:$K{fin}),0)')
    T.ligne_total(sm, dest, 1, 11)
    for col in range(2, 12):
        sm.cell(dest, col).alignment = T.DROITE
        sm.cell(dest, col).number_format = T.DH
    sm.cell(dest, 4).number_format = T.PCT
    sm.cell(dest, 10).number_format = T.PCT
    sm.cell(dest, 11).value = f'=IFERROR(AVERAGE($K{deb}:$K{fin}),0)'
    sm.cell(dest, 1).alignment = T.Alignment(horizontal='left', vertical='center', indent=1)

sm.conditional_formatting.add(
    f'I{SM_2026_1}:I{SM_N}',
    CellIsRule(operator='lessThan', formula=['0'], font=T.police(10, True, T.ROUGE),
               fill=T.fond(T.ROUGE_CLAIR)))
sm.conditional_formatting.add(
    f'B{SM_1}:B{SM_N}',
    DataBarRule(start_type='num', start_value=0, end_type='percentile', end_value=100,
                color='9CCC9C', showValue=True))
sm.conditional_formatting.add(
    f'D{SM_1}:D{SM_N}',
    ColorScaleRule(start_type='num', start_value=0.2, start_color='C8E6C9',
                   mid_type='num', mid_value=0.4, mid_color='FFF3C4',
                   end_type='num', end_value=0.6, end_color='F5C6C0'))

# Graphique : recette mensuelle sur deux ans
g5 = BarChart(); g5.type = 'col'; g5.gapWidth = 35
g5.add_data(Reference(sm, min_col=2, min_row=SM_H, max_row=SM_N), titles_from_data=True)
g5.add_data(Reference(sm, min_col=3, min_row=SM_H, max_row=SM_N), titles_from_data=True)
g5.set_categories(Reference(sm, min_col=1, min_row=SM_1, max_row=SM_N))
couleur_serie(g5.series[0], '2E7D32')
couleur_serie(g5.series[1], 'C8A265')
habiller(g5, 'Recette et achats marchandise, mois par mois (DH)', 9.5, 26)
sm.add_chart(g5, 'A34')

g6 = LineChart()
g6.add_data(Reference(sm, min_col=4, min_row=SM_H, max_row=SM_N), titles_from_data=True)
g6.set_categories(Reference(sm, min_col=1, min_row=SM_1, max_row=SM_N))
couleur_serie(g6.series[0], 'B26A00', ligne_seule=True, largeur_pt=26000)
g6.series[0].marker = Marker(symbol='diamond', size=6)
habiller(g6, 'Poids des achats dans la recette (%)', 8, 26, legende=None)
g6.y_axis.numFmt = '0%'
sm.add_chart(g6, 'A54')

sm.freeze_panes = 'B5'
T.note(sm, SM_N + 4, 1, 11,
       "« n.d. » = non disponible : l’ancien classeur (feuille Feuil1) ne suivait en 2025 que la recette "
       "et les achats de marchandise, pas les salaires ni les charges fixes. La colonne « % ACHATS » reste "
       "donc la seule comparable entre 2025 et 2026. À partir de 2026, toutes les colonnes se remplissent "
       "seules dès que vous saisissez dans le JOURNAL.", 44)


# =========================================================================
# 7. CHARGES FIXES  -- budget mensuel contre realise
# =========================================================================
cf = feuille('CHARGES FIXES', T.AMBRE)
larg(cf, {'A': 4, 'B': 46, 'C': 16, 'D': 16, 'E': 14, 'F': 12, 'G': 4,
          'H': 26, 'I': 14, 'J': 4})

T.bandeau(cf, 1, 2, 6, 'CHARGES FIXES ET ABONNEMENTS',
          'Budget mensuel (à vous de le fixer) comparé au réalisé calculé depuis le JOURNAL.')

CF_H, CF_1 = 4, 5
T.entetes(cf, CF_H, 2, ['POSTE', 'BUDGET / MOIS', 'RÉALISÉ DU MOIS', 'ÉCART', 'ÉTAT'])

postes = [
    ('Loyer', 17000, ['Loyer']),
    ('Eau + Électricité', 4000, ['Eau & Électricité']),
    ('Internet & télécom', 500, ['Internet & Télécom']),
    ('Achats marchandise (café, eau, épicerie)', 20500,
     GROUPES['achats']),
    ('Nettoyage (NADAFA)', 4500, ['Nettoyage']),
    ('Salaires du personnel', 17000, ['Salaires']),
    ('Taxes et impôts', 0, ['Taxes & Impôts']),
    ('Abonnements (Bein, Remo…)', 0, ['Abonnements']),
    ('Entretien et réparation', 0, ['Entretien & Réparation']),
]
for i, (lib, budget, cats) in enumerate(postes):
    r = CF_1 + i
    cf.cell(r, 2, lib).alignment = T.GAUCHE
    b = cf.cell(r, 3, budget)
    b.font = T.police(10, True, T.BLEU); b.fill = T.fond(T.JAUNE)
    cf.cell(r, 4, f'={somme_cats(cats, J_CLE, CLE)}')
    cf.cell(r, 5, f'=$C{r}-$D{r}')
    cf.cell(r, 6, f'=IF(N($C{r})=0,IF(N($D{r})=0,"—","HORS BUDGET"),IF($D{r}<=$C{r},"OK","DEPASSE"))')
    for col in range(2, 7):
        cell = cf.cell(r, col); cell.border = T.BORD_LEGER
        if col != 3:
            cell.fill = T.fond(T.CREME if i % 2 else T.BLANC)
        if col != 3:
            cell.font = T.police(10)
        cell.number_format = T.DH; cell.alignment = T.DROITE
    cf.cell(r, 2).alignment = T.GAUCHE
    cf.cell(r, 2).number_format = 'General'
    cf.cell(r, 6).number_format = 'General'; cf.cell(r, 6).alignment = T.CENTRE
    cf.cell(r, 4).font = T.police(10, True)
    cf.row_dimensions[r].height = 19

CF_N = CF_1 + len(postes) - 1
CF_T = CF_N + 1
cf.cell(CF_T, 2, 'TOTAL CHARGES DU MOIS')
for col in (3, 4, 5):
    L = GL(col)
    cf.cell(CF_T, col, f'=SUM({L}{CF_1}:{L}{CF_N})')
T.ligne_total(cf, CF_T, 2, 6)
for col in (3, 4, 5):
    cf.cell(CF_T, col).number_format = T.DH; cf.cell(CF_T, col).alignment = T.DROITE
cf.cell(CF_T, 2).alignment = T.Alignment(horizontal='left', vertical='center', indent=1)

CF_J = CF_T + 2
syntheses = [
    ('Recette minimum / jour pour couvrir les charges (seuil)',
     f'=IFERROR($C{CF_T}/30,0)', T.ROUGE,
     "Seuil de rentabilité : total du BUDGET mensuel divisé par 30 jours. "
     "En dessous de ce chiffre de recette, la journée est déficitaire."),
    ('Dépenses réelles par jour, ce mois',
     f'=IFERROR($D{CF_T}/{NBJ},0)', T.AMBRE,
     "Total réellement dépensé ce mois, divisé par le nombre de jours saisis."),
    ('Recette moyenne par jour, ce mois',
     f'=IFERROR({REC}/{NBJ},0)', T.VERT,
     "À comparer au seuil ci-dessus : au-dessus, le café couvre ses charges."),
    ('Marge de sécurité par jour',
     f'=IFERROR({REC}/{NBJ},0)-IFERROR($C{CF_T}/30,0)', T.ESPRESSO,
     "Recette moyenne moins le seuil. Négatif = les charges ne sont pas couvertes."),
]
for _i, (_lib, _f, _col, _msg) in enumerate(syntheses):
    _r = CF_J + _i
    cf.cell(_r, 2, _lib).font = T.police(10, True, T.ESPRESSO)
    c = cf.cell(_r, 3, _f)
    c.number_format = T.DH; c.alignment = T.DROITE; c.font = T.police(11, True, _col)
    c.fill = T.fond(T.CREME_2); c.border = T.BORD_BOITE
    c.comment = Comment(_msg, 'Système', height=100, width=260)
cf.conditional_formatting.add(f'C{CF_J + 3}', CellIsRule(
    operator='lessThan', formula=['0'], font=T.police(11, True, T.ROUGE)))

cf.conditional_formatting.add(f'F{CF_1}:F{CF_N}', CellIsRule(
    operator='equal', formula=['"OK"'], font=T.police(9, True, T.VERT), fill=T.fond(T.VERT_CLAIR)))
cf.conditional_formatting.add(f'F{CF_1}:F{CF_N}', CellIsRule(
    operator='equal', formula=['"DEPASSE"'], font=T.police(9, True, T.ROUGE), fill=T.fond(T.ROUGE_CLAIR)))
cf.conditional_formatting.add(f'F{CF_1}:F{CF_N}', CellIsRule(
    operator='equal', formula=['"HORS BUDGET"'], font=T.police(8, True, T.AMBRE), fill=T.fond(T.JAUNE)))
cf.conditional_formatting.add(f'E{CF_1}:E{CF_N}', CellIsRule(
    operator='lessThan', formula=['0'], font=T.police(10, True, T.ROUGE)))

# --- Abonnements (feuille Feuil2 de l ancien classeur) -------------------
T.titre_section(cf, CF_H - 1, 8, 9, 'ABONNEMENTS ANNUELS')
T.entetes(cf, CF_H, 8, ['ABONNEMENT', 'MONTANT'])
for i, (lib, mnt) in enumerate(D.ABONNEMENTS):
    r = CF_1 + i
    cf.cell(r, 8, lib).alignment = T.GAUCHE
    c = cf.cell(r, 9, mnt)
    c.number_format = T.DH; c.alignment = T.DROITE; c.font = T.police(10, True, T.BLEU)
    for col in (8, 9):
        cf.cell(r, col).border = T.BORD_LEGER
        cf.cell(r, col).fill = T.fond(T.CREME if i % 2 else T.BLANC)
    cf.cell(r, 8).font = T.police(10)
AB_N = CF_1 + len(D.ABONNEMENTS) - 1
cf.cell(AB_N + 1, 8, 'TOTAL')
cf.cell(AB_N + 1, 9, f'=SUM($I${CF_1}:$I${AB_N})').number_format = T.DH
T.ligne_total(cf, AB_N + 1, 8, 9)
cf.cell(AB_N + 1, 9).alignment = T.DROITE
cf.cell(AB_N + 1, 8).alignment = T.Alignment(horizontal='left', vertical='center', indent=1)

T.note(cf, AB_N + 3, 8, 9,
       "Repris de la feuille Feuil2 de l’ancien classeur (total 5 735 DH).", 30)

T.note(cf, CF_J + 6, 2, 6,
       "Les budgets en jaune viennent de la feuille CHARGE de l’ancien classeur (colonne « CHARGE FIXE » "
       "de droite, total 75 000 DH par mois, associés compris). Ajustez-les librement : le réalisé, lui, "
       "est toujours calculé depuis le JOURNAL sur le mois choisi dans PARAMÈTRES.", 44)

# Graphique budget vs realise
g7 = BarChart(); g7.type = 'bar'; g7.grouping = 'clustered'; g7.gapWidth = 45
g7.add_data(Reference(cf, min_col=3, min_row=CF_H, max_row=CF_N), titles_from_data=True)
g7.add_data(Reference(cf, min_col=4, min_row=CF_H, max_row=CF_N), titles_from_data=True)
g7.set_categories(Reference(cf, min_col=2, min_row=CF_1, max_row=CF_N))
couleur_serie(g7.series[0], 'C8A265')
couleur_serie(g7.series[1], '5D4037')
habiller(g7, 'Budget contre réalisé, poste par poste (DH)', 10, 24)
cf.add_chart(g7, 'B25')


# =========================================================================
# 8. SALAIRES  -- personnel et suivi des versements
# =========================================================================
sa = feuille('SALAIRES', T.MOKA)
larg(sa, {'A': 4, 'B': 24, 'C': 18, 'D': 16, 'E': 16, 'F': 14, 'G': 12, 'H': 14, 'I': 4})

T.bandeau(sa, 1, 2, 8, 'PERSONNEL ET SALAIRES',
          'Salaire convenu par personne, contre ce qui a réellement été versé ce mois-ci.')

SA_H, SA_1 = 4, 5
T.entetes(sa, SA_H, 2,
          ['EMPLOYÉ', 'POSTE', 'SALAIRE CONVENU', 'VERSÉ CE MOIS', 'RESTE À VERSER',
           'NB VERSEMENTS', 'PART'])

for i, (nom, poste, salaire) in enumerate(D.SALAIRES):
    r = SA_1 + i
    sa.cell(r, 2, nom).alignment = T.GAUCHE
    sa.cell(r, 3, poste).alignment = T.GAUCHE
    c = sa.cell(r, 4, salaire)
    c.font = T.police(10, True, T.BLEU); c.fill = T.fond(T.JAUNE)
    if salaire is None:
        c.comment = Comment(
            "Salaire convenu inconnu : ce nom apparaît dans les paiements d’août 2026 "
            "mais ne figurait pas dans la feuille CHARGE de l’ancien classeur. "
            "Renseignez le montant pour suivre le reste à verser.",
            'Système', height=110, width=260)
    sa.cell(r, 5, f'=SUMIFS({J_MNT},{J_TIERS},$B{r},{J_CAT},"Salaires",{J_CLE},{CLE})')
    sa.cell(r, 6, f'=IF($D{r}="","—",$D{r}-$E{r})')
    sa.cell(r, 7, f'=COUNTIFS({J_TIERS},$B{r},{J_CAT},"Salaires",{J_CLE},{CLE})')
    sa.cell(r, 8, f'=IF(N($D{r})=0,"—",$E{r}/$D{r})')
    for col in range(2, 9):
        cell = sa.cell(r, col); cell.border = T.BORD_LEGER
        if col != 4:
            cell.fill = T.fond(T.CREME if i % 2 else T.BLANC)
        cell.number_format = T.DH; cell.alignment = T.DROITE
        if col not in (4,):
            cell.font = T.police(10)
    sa.cell(r, 2).alignment = T.GAUCHE; sa.cell(r, 2).number_format = 'General'
    sa.cell(r, 2).font = T.police(10, True, T.ESPRESSO)
    sa.cell(r, 3).alignment = T.GAUCHE; sa.cell(r, 3).number_format = 'General'
    sa.cell(r, 3).font = T.police(9, False, T.GRIS)
    sa.cell(r, 5).font = T.police(10, True, T.VERT)
    sa.cell(r, 7).number_format = T.NBR; sa.cell(r, 7).alignment = T.CENTRE
    sa.cell(r, 8).number_format = T.PCT; sa.cell(r, 8).font = T.police(9, False, T.AMBRE)
    sa.row_dimensions[r].height = 19

SA_N = SA_1 + len(D.SALAIRES) - 1
SA_T = SA_N + 1
sa.cell(SA_T, 2, 'TOTAL PERSONNEL')
sa.merge_cells(start_row=SA_T, start_column=2, end_row=SA_T, end_column=3)
for col in (4, 5, 7):
    L = GL(col)
    sa.cell(SA_T, col, f'=SUM({L}{SA_1}:{L}{SA_N})')
sa.cell(SA_T, 6, f'=$D{SA_T}-$E{SA_T}')
sa.cell(SA_T, 8, f'=IFERROR($E{SA_T}/$D{SA_T},0)')
T.ligne_total(sa, SA_T, 2, 8)
for col in (4, 5, 6):
    sa.cell(SA_T, col).number_format = T.DH; sa.cell(SA_T, col).alignment = T.DROITE
sa.cell(SA_T, 7).number_format = T.NBR; sa.cell(SA_T, 7).alignment = T.CENTRE
sa.cell(SA_T, 8).number_format = T.PCT; sa.cell(SA_T, 8).alignment = T.DROITE
sa.cell(SA_T, 2).alignment = T.Alignment(horizontal='left', vertical='center', indent=1)

# Total verse en salaires, y compris personnes hors liste
SA_C = SA_T + 2
sa.cell(SA_C, 2, 'Total versé en salaires ce mois (toutes personnes)').font = T.police(10, True)
sa.merge_cells(start_row=SA_C, start_column=2, end_row=SA_C, end_column=4)
c = sa.cell(SA_C, 5, f'=SUMIFS({J_MNT},{J_CAT},"Salaires",{J_CLE},{CLE})')
c.number_format = T.DH; c.alignment = T.DROITE; c.font = T.police(11, True, T.ESPRESSO)
c.fill = T.fond(T.CREME_2); c.border = T.BORD_BOITE
sa.cell(SA_C + 1, 2, 'Dont versements à des personnes hors de la liste ci-dessus').font = T.police(9, False, T.GRIS)
sa.merge_cells(start_row=SA_C + 1, start_column=2, end_row=SA_C + 1, end_column=4)
c = sa.cell(SA_C + 1, 5, f'=$E{SA_C}-$E{SA_T}')
c.number_format = T.DH; c.alignment = T.DROITE; c.font = T.police(10, True, T.AMBRE)
c.border = T.BORD_LEGER
c.comment = Comment("Si ce montant n est pas nul, c est qu un salaire a ete verse a un nom "
                    "qui ne figure pas dans le tableau ci-dessus. "
                    "Ajoutez la personne au tableau pour la suivre correctement.",
                    'Système', height=110, width=260)

sa.conditional_formatting.add(f'F{SA_1}:F{SA_N}', CellIsRule(
    operator='greaterThan', formula=['0'], font=T.police(10, True, T.ROUGE), fill=T.fond(T.ROUGE_CLAIR)))
sa.conditional_formatting.add(f'F{SA_1}:F{SA_N}', CellIsRule(
    operator='lessThanOrEqual', formula=['0'], font=T.police(10, True, T.VERT)))
sa.conditional_formatting.add(f'H{SA_1}:H{SA_N}', DataBarRule(
    start_type='num', start_value=0, end_type='num', end_value=1, color='9CCC9C', showValue=True))

T.note(sa, SA_C + 3, 2, 8,
       "Salaires convenus repris de la feuille CHARGE de l’ancien classeur. La colonne « VERSÉ CE MOIS » "
       "additionne toutes les lignes du JOURNAL de catégorie « Salaires » au nom de la personne : "
       "elle inclut donc les avances. Les noms BAR MEN, BAR MEN ALI, FATIMA, LAHCEN et MARWA ont été "
       "ajoutés au tableau car ils apparaissent dans les paiements d’août 2026, mais leur salaire convenu "
       "ne figurait nulle part dans l’ancien classeur : renseignez-le dans les cellules jaunes vides, "
       "sans quoi le « reste à verser » ne peut pas être calculé et s’affiche « — ».", 56)


# =========================================================================
# 9. MARGES PRODUITS  -- prix de vente, cout matiere, marge
# =========================================================================
mp = feuille('MARGES PRODUITS', T.CARAMEL)
larg(mp, {'A': 4, 'B': 34, 'C': 14, 'D': 15, 'E': 15, 'F': 12, 'G': 14,
          'H': 15, 'I': 15, 'J': 4})

T.bandeau(mp, 1, 2, 9, 'MARGES PAR PRODUIT',
          'Ce que rapporte chaque boisson. Remplissez le coût matière en jaune pour obtenir la marge réelle.')

MP_H, MP_1 = 4, 5
T.entetes(mp, MP_H, 2,
          ['PRODUIT', 'PRIX DE VENTE', 'COÛT MATIÈRE', 'MARGE UNITAIRE', 'MARGE %',
           'QTÉ / JOUR', 'CA / JOUR', 'MARGE / JOUR'])

for i, (nom, prix, qte_j, _qte_m) in enumerate(D.PRODUITS):
    r = MP_1 + i
    mp.cell(r, 2, nom).alignment = T.GAUCHE
    c = mp.cell(r, 3, prix); c.font = T.police(10, True, T.BLEU); c.fill = T.fond(T.JAUNE)
    c = mp.cell(r, 4, None); c.font = T.police(10, True, T.BLEU); c.fill = T.fond(T.JAUNE)
    mp.cell(r, 5, f'=IF($D{r}="","—",$C{r}-$D{r})')
    mp.cell(r, 6, f'=IF($D{r}="","—",IFERROR(($C{r}-$D{r})/$C{r},0))')
    c = mp.cell(r, 7, qte_j); c.font = T.police(10, True, T.BLEU); c.fill = T.fond(T.JAUNE)
    mp.cell(r, 8, f'=$C{r}*$G{r}')
    mp.cell(r, 9, f'=IF($D{r}="","—",($C{r}-$D{r})*$G{r})')
    for col in range(2, 10):
        cell = mp.cell(r, col); cell.border = T.BORD_LEGER
        if col not in (3, 4, 7):
            cell.fill = T.fond(T.CREME if i % 2 else T.BLANC)
            cell.font = T.police(10)
        cell.number_format = T.DH; cell.alignment = T.DROITE
    mp.cell(r, 2).alignment = T.GAUCHE; mp.cell(r, 2).number_format = 'General'
    mp.cell(r, 2).font = T.police(10, True, T.ESPRESSO)
    mp.cell(r, 6).number_format = T.PCT; mp.cell(r, 6).font = T.police(10, True, T.AMBRE)
    mp.cell(r, 7).number_format = T.NBR; mp.cell(r, 7).alignment = T.CENTRE
    mp.cell(r, 8).font = T.police(10, True, T.VERT)
    mp.row_dimensions[r].height = 20

MP_N = MP_1 + len(D.PRODUITS) - 1
MP_T = MP_N + 1
mp.cell(MP_T, 2, 'TOTAL')
for col in (7, 8):
    L = GL(col)
    mp.cell(MP_T, col, f'=SUM({L}{MP_1}:{L}{MP_N})')
mp.cell(MP_T, 9, f'=IF(COUNT($D${MP_1}:$D${MP_N})=0,"—",SUMIF($D${MP_1}:$D${MP_N},"<>",$I${MP_1}:$I${MP_N}))')
mp.cell(MP_T, 6, f'=IF(COUNT($D${MP_1}:$D${MP_N})=0,"—",IFERROR($I{MP_T}/$H{MP_T},0))')
T.ligne_total(mp, MP_T, 2, 9)
mp.cell(MP_T, 6).number_format = T.PCT; mp.cell(MP_T, 6).alignment = T.DROITE
mp.cell(MP_T, 7).number_format = T.NBR; mp.cell(MP_T, 7).alignment = T.CENTRE
for col in (8, 9):
    mp.cell(MP_T, col).number_format = T.DH; mp.cell(MP_T, col).alignment = T.DROITE
mp.cell(MP_T, 2).alignment = T.Alignment(horizontal='left', vertical='center', indent=1)

MP_C = MP_T + 2
comparaisons = [
    ('CA théorique du mois (30 jours)', f'=$H{MP_T}*30', T.MOKA),
    ('Recette réellement encaissée ce mois', f'={REC}', T.VERT),
    ('Écart théorique / réel', f'=$C{MP_C + 1}-$C{MP_C}', T.ROUGE),
]
for i, (lib, f, col) in enumerate(comparaisons):
    r = MP_C + i
    mp.cell(r, 2, lib).font = T.police(10, True)
    mp.merge_cells(start_row=r, start_column=2, end_row=r, end_column=2)
    c = mp.cell(r, 3, f)
    c.number_format = T.DH; c.alignment = T.DROITE
    c.font = T.police(11, True, col); c.fill = T.fond(T.CREME_2); c.border = T.BORD_BOITE

T.note(mp, MP_C + 4, 2, 9,
       "Prix de vente et quantités journalières repris de la feuille CHARGE de l’ancien classeur "
       "(café noir 13 DH × 100, crème 13 DH × 20, petit-déjeuner 20 DH × 8, thé 12 DH × 10). "
       "La colonne COÛT MATIÈRE est volontairement vide : renseignez-la (prix de revient d’une tasse) "
       "pour que la marge se calcule ; tant qu’elle est vide, les colonnes de marge affichent « — ».", 50)

g8 = BarChart(); g8.type = 'col'; g8.gapWidth = 50
g8.add_data(Reference(mp, min_col=8, min_row=MP_H, max_row=MP_N), titles_from_data=True)
g8.add_data(Reference(mp, min_col=9, min_row=MP_H, max_row=MP_N), titles_from_data=True)
g8.set_categories(Reference(mp, min_col=2, min_row=MP_1, max_row=MP_N))
couleur_serie(g8.series[0], 'C8A265')
couleur_serie(g8.series[1], '2E7D32')
habiller(g8, "Chiffre d’affaires et marge par produit, par jour (DH)", 9, 20)
mp.add_chart(g8, 'B20')


# =========================================================================
# 10. ASSOCIÉS  -- comptes Hamid / Mouhssine
# =========================================================================
asc = feuille('ASSOCIÉS', T.MOKA)
larg(asc, {'A': 4, 'B': 6, 'C': 58, 'D': 16, 'E': 16, 'F': 16, 'G': 4})

T.bandeau(asc, 1, 2, 6, 'COMPTES DES ASSOCIÉS',
          'Apports de chaque associé : achat du fonds de commerce et aménagement du café. '
          'Historique repris de la feuille DEPENSE_CAFE.')

# --- Synthese ------------------------------------------------------------
T.titre_section(asc, 4, 2, 6, 'SYNTHÈSE')
T.entetes(asc, 5, 2, ['', 'POSTE', 'HAMID', 'MOUHSSINE', 'TOTAL'])
asc.column_dimensions['B'].width = 6

AC_1 = 6
ACQ_1, ACQ_N = 30, 30 + len(D.ACQUISITION_DETAIL) - 1
AM_1 = ACQ_N + 5
AM_N = AM_1 + len(D.AMENAGEMENT_DETAIL) - 1

synthese = [
    ('Achat du fonds de commerce (Saadaoui)', f'=$D${ACQ_N + 1}', f'=$E${ACQ_N + 1}'),
    ('Aménagement du Café Victor Hugo', f'=$D${AM_N + 1}', f'=$E${AM_N + 1}'),
]
for i, (lib, fh, fm) in enumerate(synthese):
    r = AC_1 + i
    asc.cell(r, 3, lib).alignment = T.GAUCHE
    asc.cell(r, 4, fh); asc.cell(r, 5, fm)
    asc.cell(r, 6, f'=$D{r}+$E{r}')

r = AC_1 + 2
asc.cell(r, 3, 'TOTAL APPORTÉ PAR ASSOCIÉ').alignment = T.GAUCHE
asc.cell(r, 4, f'=SUM($D${AC_1}:$D${AC_1 + 1})')
asc.cell(r, 5, f'=SUM($E${AC_1}:$E${AC_1 + 1})')
asc.cell(r, 6, f'=$D{r}+$E{r}')

r = AC_1 + 3
asc.cell(r, 3, 'Part théorique due sur le fonds (795 000 DH / 2)').alignment = T.GAUCHE
asc.cell(r, 4, D.ACQUISITION['part_par_associe'])
asc.cell(r, 5, D.ACQUISITION['part_par_associe'])
asc.cell(r, 6, f'=$D{r}+$E{r}')

r = AC_1 + 4
asc.cell(r, 3, "Reste à payer sur le fonds  (+ = doit,  - = a trop versé)").alignment = T.GAUCHE
asc.cell(r, 4, f'=$D{AC_1 + 3}-$D{AC_1}')
asc.cell(r, 5, f'=$E{AC_1 + 3}-$E{AC_1}')
asc.cell(r, 6, f'=$D{r}+$E{r}')

r = AC_1 + 5
asc.cell(r, 3, 'Écart sur le fonds seul (Hamid - Mouhssine)').alignment = T.GAUCHE
asc.cell(r, 6, f'=$D{AC_1}-$E{AC_1}')

r = AC_1 + 6
asc.cell(r, 3, "Écart sur l’aménagement seul (Hamid - Mouhssine)").alignment = T.GAUCHE
asc.cell(r, 6, f'=$D{AC_1 + 1}-$E{AC_1 + 1}')

r = AC_1 + 7
asc.cell(r, 3, 'ÉCART TOTAL ENTRE LES DEUX ASSOCIÉS').alignment = T.GAUCHE
asc.cell(r, 6, f'=$D{AC_1 + 2}-$E{AC_1 + 2}')

for i in range(8):
    rr = AC_1 + i
    gras = i in (2, 4, 7)
    for col in range(2, 7):
        cell = asc.cell(rr, col); cell.border = T.BORD_LEGER
        cell.fill = T.fond(T.CREME_2 if gras else (T.CREME if i % 2 else T.BLANC))
        cell.number_format = T.DH; cell.alignment = T.DROITE
        cell.font = T.police(10, gras, T.ESPRESSO if gras else T.ENCRE)
    asc.cell(rr, 3).alignment = T.GAUCHE
    asc.cell(rr, 3).number_format = 'General'
    asc.row_dimensions[rr].height = 20

asc.conditional_formatting.add(f'D{AC_1 + 4}:F{AC_1 + 7}', CellIsRule(
    operator='lessThan', formula=['0'], font=T.police(10, True, T.VERT)))
asc.conditional_formatting.add(f'D{AC_1 + 4}:F{AC_1 + 7}', CellIsRule(
    operator='greaterThan', formula=['0'], font=T.police(10, True, T.ROUGE)))

T.note(asc, AC_1 + 9, 2, 6,
       "Lecture : un « reste à payer » négatif signifie que l’associé a versé plus que sa part. "
       "Au dernier pointage de l’ancien classeur, Hamid avait versé 238 DH de trop et Mouhssine 7 200 DH "
       "de trop sur l’achat du fonds, soit 6 962 DH d’écart entre eux. Sur l’aménagement du café, Mouhssine "
       "a apporté 150 779 DH de plus que Hamid, ce qui porte l’écart cumulé à 157 741 DH.", 44)

# --- Detail acquisition --------------------------------------------------
T.titre_section(asc, ACQ_1 - 2, 2, 6, 'DÉTAIL — ACHAT DU FONDS DE COMMERCE (SAADAOUI, 795 000 DH)')
T.entetes(asc, ACQ_1 - 1, 2, ['N°', 'OPÉRATION', 'HAMID', 'MOUHSSINE', 'TOTAL'])


def table_detail(ws, lignes, r1, couleur_bande):
    for i, (lib, h, m) in enumerate(lignes):
        r = r1 + i
        ws.cell(r, 2, i + 1).alignment = T.CENTRE
        ws.cell(r, 2).font = T.police(8, False, T.GRIS)
        ws.cell(r, 3, lib).alignment = T.GAUCHE
        ws.cell(r, 3).font = T.police(9)
        ws.cell(r, 4, h if h else None).number_format = T.DH
        ws.cell(r, 5, m if m else None).number_format = T.DH
        ws.cell(r, 6, f'=N($D{r})+N($E{r})').number_format = T.DH
        for col in range(2, 7):
            cell = ws.cell(r, col); cell.border = T.BORD_LEGER
            cell.fill = T.fond(T.CREME if i % 2 else T.BLANC)
            if col >= 4:
                cell.alignment = T.DROITE
                cell.font = T.police(9, False, T.BLEU if col < 6 else T.ENCRE)
        ws.row_dimensions[r].height = 16
    rt = r1 + len(lignes)
    ws.cell(rt, 3, 'TOTAL')
    for col in (4, 5, 6):
        L = GL(col)
        ws.cell(rt, col, f'=SUM({L}{r1}:{L}{r1 + len(lignes) - 1})').number_format = T.DH
    T.ligne_total(ws, rt, 2, 6)
    for col in (4, 5, 6):
        ws.cell(rt, col).alignment = T.DROITE
    ws.cell(rt, 3).alignment = T.Alignment(horizontal='left', vertical='center', indent=1)
    return rt


table_detail(asc, D.ACQUISITION_DETAIL, ACQ_1, T.CREME)

T.titre_section(asc, AM_1 - 2, 2, 6, 'DÉTAIL — AMÉNAGEMENT ET ÉQUIPEMENT DU CAFÉ VICTOR HUGO')
T.entetes(asc, AM_1 - 1, 2, ['N°', 'OPÉRATION', 'HAMID', 'MOUHSSINE', 'TOTAL'])
table_detail(asc, D.AMENAGEMENT_DETAIL, AM_1, T.CREME)

T.note(asc, AM_N + 3, 2, 6,
       "Ces deux tableaux sont l’historique complet repris de la feuille DEPENSE_CAFE de l’ancien classeur, "
       "ligne par ligne, sans modification. Ce sont des données figées : elles ne bougent plus. "
       "Les montants en bleu sont des valeurs saisies, pas des formules.", 40)

g9 = BarChart(); g9.type = 'col'; g9.grouping = 'stacked'; g9.overlap = 100; g9.gapWidth = 60
g9.add_data(Reference(asc, min_col=4, min_row=5, max_row=AC_1 + 1), titles_from_data=True)
g9.add_data(Reference(asc, min_col=5, min_row=5, max_row=AC_1 + 1), titles_from_data=True)
g9.set_categories(Reference(asc, min_col=3, min_row=AC_1, max_row=AC_1 + 1))
couleur_serie(g9.series[0], '5D4037')
couleur_serie(g9.series[1], 'C8A265')
habiller(g9, 'Apports par associé (DH)', 8, 17)
asc.add_chart(g9, 'B17')


# =========================================================================
# 11. ARCHIVE  -- donnees conservees mais hors exploitation du cafe
# =========================================================================
ar = feuille('ARCHIVE', T.GRIS)
larg(ar, {'A': 4, 'B': 40, 'C': 16, 'D': 4, 'E': 40, 'F': 16, 'G': 4})
T.bandeau(ar, 1, 2, 6, 'ARCHIVE',
          'Données de l’ancien classeur conservées pour mémoire, sans lien avec l’exploitation du café.')

T.titre_section(ar, 4, 2, 3, 'COMMANDE TEST CHINE — environ 25 000 DH')
T.entetes(ar, 5, 2, ['DÉSIGNATION', 'QUANTITÉ'])
chine = [(' raccord droit 16 mm', 10000), ('te 16 mm', 5000), ('coude 16 mm', 5000),
         ('bouchon 16 mm', 5000), (' depart 16 mm', 2000), ('vannette 16 mm', 2000)]
turquie = [(' filtre disque 2"', 100), (' filtre tamis 2"', 100), (' manometre', 200),
           (' regulateur', 200), ('vanne', 200), ('raccords PE 32/40/50 mm', 500)]
for i, (lib, q) in enumerate(chine):
    r = 6 + i
    ar.cell(r, 2, lib).alignment = T.GAUCHE
    ar.cell(r, 3, q).number_format = T.NBR
    for col in (2, 3):
        ar.cell(r, col).border = T.BORD_LEGER
        ar.cell(r, col).fill = T.fond(T.CREME if i % 2 else T.BLANC)
        ar.cell(r, col).font = T.police(9)
    ar.cell(r, 3).alignment = T.DROITE

T.titre_section(ar, 4, 5, 6, 'COMMANDE TEST TURQUIE — environ 30 000 DH')
T.entetes(ar, 5, 5, ['DÉSIGNATION', 'QUANTITÉ'])
for i, (lib, q) in enumerate(turquie):
    r = 6 + i
    ar.cell(r, 5, lib).alignment = T.GAUCHE
    ar.cell(r, 6, q).number_format = T.NBR
    for col in (5, 6):
        ar.cell(r, col).border = T.BORD_LEGER
        ar.cell(r, col).fill = T.fond(T.CREME if i % 2 else T.BLANC)
        ar.cell(r, col).font = T.police(9)
    ar.cell(r, 6).alignment = T.DROITE

T.note(ar, 14, 2, 6,
       "Repris tel quel de la feuille « PVC » de l’ancien classeur. Il s’agit d’une commande de "
       "matériel d’irrigation, sans rapport avec le café : elle est conservée ici pour ne rien perdre.", 34)


# =========================================================================
# 12. ACCUEIL  -- page de garde et mode d emploi
# =========================================================================
ac = feuille('ACCUEIL', T.ESPRESSO)
larg(ac, {'A': 3, 'B': 28, 'C': 17, 'D': 20, 'E': 20, 'F': 20, 'G': 20, 'H': 20, 'I': 3})

for r in range(1, 6):
    for col in range(1, 10):
        ac.cell(r, col).fill = T.fond(T.ESPRESSO)
ac.merge_cells('B2:H2')
c = ac['B2']; c.value = '=PARAMETRES!$C$5'
c.font = T.police(30, True, T.BLANC); c.alignment = T.GAUCHE
ac.row_dimensions[2].height = 46
ac.merge_cells('B3:H3')
c = ac['B3']; c.value = '=PARAMETRES!$C$6&", "&PARAMETRES!$C$7'
c.font = T.police(13, False, T.CARAMEL); c.alignment = T.GAUCHE
ac.row_dimensions[3].height = 22
ac.merge_cells('B4:H4')
c = ac['B4']; c.value = 'SYSTÈME DE GESTION — CAISSE, DÉPENSES, CHARGES ET RÉSULTATS'
c.font = T.police(10, True, T.SABLE); c.alignment = T.GAUCHE
ac.row_dimensions[4].height = 20
ac.row_dimensions[1].height = 14
ac.row_dimensions[5].height = 10

# --- Chiffres du mois ----------------------------------------------------
ac.merge_cells('B7:C7')
c = ac['B7']; c.value = 'LE MOIS EN COURS'
c.font = T.police(11, True, T.ESPRESSO)
c.alignment = T.Alignment(horizontal='left', vertical='center', indent=1)
c.fill = T.fond(T.CARAMEL)
ac.merge_cells('D7:H7')
c = ac['D7']; c.value = '=PARAMETRES!$C$15'
c.font = T.police(11, True, T.ESPRESSO); c.number_format = T.MOIS
c.alignment = T.GAUCHE
c.fill = T.fond(T.CARAMEL)
ac.row_dimensions[7].height = 24
for col in range(2, 9):
    ac.cell(7, col).fill = T.fond(T.CARAMEL)

chiffres = [
    ('Recette', f'={REC}', T.VERT),
    ('Dépenses', f'={DEP}', T.ROUGE),
    ('Résultat', f'={REC}-{DEP}', T.ESPRESSO),
    ('Marge', f'=IFERROR(({REC}-{DEP})/{REC},0)', T.AMBRE),
    ('Caisse', f"='RECAP JOUR'!$O${RJ_T}", T.MOKA),
    ('Jours saisis', f'={NBJ}', T.GRIS),
]
for i, (lib, f, col) in enumerate(chiffres):
    cc = 2 + i
    a = ac.cell(8, cc, lib)
    a.font = T.police(8.5, True, T.GRIS); a.alignment = T.CENTRE
    a.fill = T.fond(T.CREME); a.border = T.Border(top=T._s(col, 'thick'),
                                                  left=T._s(T.SABLE), right=T._s(T.SABLE))
    b = ac.cell(9, cc, f)
    b.font = T.police(15, True, col); b.alignment = T.CENTRE
    b.fill = T.fond(T.CREME)
    b.border = T.Border(left=T._s(T.SABLE), right=T._s(T.SABLE), bottom=T._s(T.SABLE))
    b.number_format = T.PCT if lib == 'Marge' else (T.NBR if lib == 'Jours saisis' else T.DH)
ac.row_dimensions[8].height = 17
ac.row_dimensions[9].height = 32

# --- Mode d emploi -------------------------------------------------------
T.titre_section(ac, 11, 2, 8, 'COMMENT UTILISER CE CLASSEUR — 3 ÉTAPES')
etapes = [
    ('1', 'CHAQUE SOIR, APRÈS LA FERMETURE',
     "Ouvrez la feuille JOURNAL. Sur la première ligne vide : la date du jour, le type RECETTE, "
     "la catégorie « Vente Caisse », et le montant compté dans la caisse. Puis une ligne DEPENSE "
     "par achat ou paiement de la journée (fournisseur, montant)."),
    ('2', 'VÉRIFIEZ QUE LA LIGNE EST VERTE',
     "La dernière colonne du JOURNAL affiche OK en vert quand la ligne est complète. "
     "Si elle affiche « Catégorie ? » ou « Montant ? », corrigez avant de passer à la suite."),
    ('3', 'REGARDEZ LE TABLEAU DE BORD',
     "Tout est déjà calculé : recette du mois, dépenses, résultat, marge, caisse, graphiques. "
     "Vous n’avez jamais à recopier un chiffre d’une feuille à l’autre."),
]
r = 12
for num, titre, txt in etapes:
    ac.cell(r, 2, num)
    ac.merge_cells(start_row=r, start_column=2, end_row=r + 1, end_column=2)
    cn = ac.cell(r, 2)
    cn.font = T.police(30, True, T.CARAMEL); cn.alignment = T.CENTRE
    cn.fill = T.fond(T.CREME)
    ac.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
    ct = ac.cell(r, 3, titre)
    ct.font = T.police(11, True, T.ESPRESSO); ct.alignment = T.GAUCHE
    ct.fill = T.fond(T.CREME)
    ac.merge_cells(start_row=r + 1, start_column=3, end_row=r + 1, end_column=8)
    cd = ac.cell(r + 1, 3, txt)
    cd.font = T.police(9.5, False, T.ENCRE); cd.alignment = T.HAUT_G
    cd.fill = T.fond(T.CREME)
    for cc in range(2, 9):
        ac.cell(r, cc).fill = T.fond(T.CREME)
        ac.cell(r + 1, cc).fill = T.fond(T.CREME)
        ac.cell(r, cc).border = T.Border(top=T._s(T.SABLE))
        ac.cell(r + 1, cc).border = T.Border(bottom=T._s(T.SABLE))
    ac.row_dimensions[r].height = 22
    ac.row_dimensions[r + 1].height = 34
    r += 3

# --- Sommaire ------------------------------------------------------------
T.titre_section(ac, r, 2, 8, 'LES FEUILLES DU CLASSEUR')
r += 1
T.entetes(ac, r, 2, ['FEUILLE', 'À QUOI ELLE SERT', '', '', '', '', 'SAISIE ?'])
ac.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
r += 1
sommaire = [
    ('ACCUEIL', "Cette page : mode d’emploi et chiffres du mois.", 'Non'),
    ('PARAMETRES', "Nom du commerce, mois analysé, caisse de départ, listes de fournisseurs et de catégories.", 'Oui'),
    ('JOURNAL', "Le cœur du système : une ligne par opération, recette ou dépense. C’est la seule feuille à remplir tous les jours.", 'Oui'),
    ('RECAP JOUR', "Le mois jour par jour : recette, dépenses par type, résultat, caisse cumulée.", 'Non'),
    ('TABLEAU DE BORD', "La photo du mois : indicateurs et graphiques.", 'Non'),
    ('SUIVI MENSUEL', "Mois par mois, 2025 et 2026, pour voir les tendances.", 'Non'),
    ('CHARGES FIXES', "Budget mensuel par poste comparé au réalisé, et coût de fonctionnement par jour.", 'Budget'),
    ('SALAIRES', "Le personnel, le salaire convenu et ce qui a été versé dans le mois.", 'Salaires'),
    ('MARGES PRODUITS', "Ce que rapporte chaque boisson une fois le coût matière déduit.", 'Couts'),
    ('ASSOCIÉS', "Apports de Hamid et Mouhssine : achat du fonds et aménagement du café.", 'Non'),
    ('ARCHIVE', "Données de l’ancien classeur sans rapport avec le café, conservées pour mémoire.", 'Non'),
]
for i, (nom, desc, saisie) in enumerate(sommaire):
    rr = r + i
    ac.cell(rr, 2, nom).font = T.police(9.5, True, T.ESPRESSO)
    ac.cell(rr, 2).alignment = T.GAUCHE
    ac.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=7)
    ac.cell(rr, 3, desc).font = T.police(9)
    ac.cell(rr, 3).alignment = T.GAUCHE_W
    ac.cell(rr, 8, saisie).font = T.police(8.5, True, T.VERT if saisie != 'Non' else T.GRIS)
    ac.cell(rr, 8).alignment = T.CENTRE
    for cc in range(2, 9):
        ac.cell(rr, cc).border = T.BORD_LEGER
        ac.cell(rr, cc).fill = T.fond(T.CREME if i % 2 else T.BLANC)
    ac.row_dimensions[rr].height = 20
r += len(sommaire) + 1

# --- Note de reprise -----------------------------------------------------
T.titre_section(ac, r, 2, 8, "CE QUI A ÉTÉ REPRIS DE L’ANCIEN CLASSEUR")
r += 1
reprises = [
    ("27 journées d’août 2026", f"{NB_OPS} opérations transférées dans le JOURNAL",
     "54 937 DH de recettes et 50 714 DH de dépenses — identiques à la feuille SUIVI d’origine."),
    ("Historique mensuel 2025", "10 mois de recettes et d’achats",
     "Repris dans SUIVI MENSUEL depuis la feuille Feuil1."),
    ("Charges fixes et salaires", "9 postes de charges, 11 personnes",
     "Repris dans CHARGES FIXES et SALAIRES depuis la feuille CHARGE."),
    ("Comptes des associés", "105 lignes d’apports",
     "Reprises dans ASSOCIÉS depuis la feuille DEPENSE_CAFE, sans modification."),
    ("Prix de vente", "4 produits",
     "Repris dans MARGES PRODUITS depuis la feuille CHARGE."),
]
for i, (quoi, combien, detail) in enumerate(reprises):
    rr = r + i
    ac.cell(rr, 2, quoi).font = T.police(9.5, True, T.ESPRESSO)
    ac.cell(rr, 2).alignment = T.GAUCHE
    ac.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=4)
    ac.cell(rr, 3, combien).font = T.police(9, True, T.AMBRE)
    ac.cell(rr, 3).alignment = T.GAUCHE
    ac.merge_cells(start_row=rr, start_column=5, end_row=rr, end_column=8)
    ac.cell(rr, 5, detail).font = T.police(9, False, T.GRIS)
    ac.cell(rr, 5).alignment = T.GAUCHE_W
    for cc in range(2, 9):
        ac.cell(rr, cc).border = T.BORD_LEGER
        ac.cell(rr, cc).fill = T.fond(T.CREME if i % 2 else T.BLANC)
    ac.row_dimensions[rr].height = 20
r += len(reprises) + 1

T.note(ac, r, 2, 8,
       "Différence avec l’ancien classeur : il fallait auparavant créer une feuille par jour et recopier "
       "les totaux à la main dans la feuille SUIVI. Ici, une seule feuille se remplit (JOURNAL) et les "
       "31 jours du mois, le tableau de bord, le suivi mensuel et les charges se calculent seuls. "
       "Pour changer de mois, il suffit de modifier le numéro du mois dans PARAMÈTRES — rien d’autre.", 46)

# Ordre d affichage des onglets
ordre = ['ACCUEIL', 'TABLEAU DE BORD', 'JOURNAL', 'RECAP JOUR', 'SUIVI MENSUEL',
         'CHARGES FIXES', 'SALAIRES', 'MARGES PRODUITS', 'ASSOCIÉS',
         'PARAMETRES', 'ARCHIVE']
wb._sheets = [wb[n] for n in ordre]
wb.active = 0

for ws in wb.worksheets:
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True

SORTIE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                      'CAFE_VICTOR_HUGO_SYSTEME.xlsx')
wb.save(SORTIE)
print('Classeur ecrit :', SORTIE)
print('Operations transferees dans le JOURNAL :', NB_OPS)
