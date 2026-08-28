# -*- coding: utf-8 -*-
"""
Donnees historiques extraites du classeur d’origine SUIVICAISSE_VH2026_08.xlsx
(Cafe Victor Hugo - Mohammedia, Maroc).

Toutes les valeurs de ce module proviennent directement du fichier fourni par
le proprietaire. Aucune valeur n’est inventee. Les feuilles d’origine sont :
  - 01-08-2026 ... 27-08-2026 : caisse journaliere      -> JOURNAL
  - SUIVI                     : recap mensuel           -> RECAP JOUR
  - Feuil1                    : historique mensuel 2025 -> SUIVI MENSUEL
  - CHARGE                    : charges fixes + prix    -> CHARGES / MARGES
  - DEPENSE_CAFE              : comptes associes        -> ASSOCIES
  - Feuil2                    : abonnements             -> CHARGES FIXES
  - PVC                       : commande hors cafe      -> ARCHIVE DIVERS
"""

# --- Recettes journalieres (A20 de chaque feuille jour) ---------------------
RECETTES = {
    1: 2335, 2: 2456, 3: 2456, 4: 2145, 5: 2233, 6: 2074, 7: 1892,
    8: 2172, 9: 2218, 10: 1548, 11: 1973, 12: 2045, 13: 2014, 14: 1902,
    15: 2053, 16: 2100, 17: 1930, 18: 2081, 19: 2002, 20: 2253, 21: 1607,
    22: 1853, 23: 2141, 24: 2154, 25: 1900, 26: 1713, 27: 1687,
}

# --- Depenses journalieres : (jour, beneficiaire, montant, bloc d’origine) --
# bloc : A = Depense Achat, F = Charge Fixe, S = Salaire, V = Virement
DEPENSES = [
    (1, 'AIN SAISS', 306, 'A'), (1, 'BIM', 300, 'A'), (1, 'ORANGE', 165, 'A'), (1, 'NADAFA', 150, 'A'),
    (2, 'AIN SAISS', 306, 'A'), (2, 'BIM', 300, 'A'), (2, 'OULMES', 156, 'A'), (2, 'NADAFA', 150, 'A'),
    (2, 'BAR MEN', 600, 'S'), (2, 'BAR MEN ALI', 600, 'S'), (2, 'FATIMA', 400, 'S'),
    (2, 'LAHCEN', 500, 'S'), (2, 'ABDELLAH', 600, 'S'),
    (3, 'NADAFA', 150, 'A'), (3, 'BIM', 300, 'A'), (3, 'ORANGE', 165, 'A'),
    (4, 'NADAFA', 150, 'A'), (4, 'BIM', 200, 'A'), (4, 'AIN SAISS', 306, 'A'), (4, 'MECAFE', 910, 'A'),
    (5, 'NADAFA', 150, 'A'), (5, 'AIN SAISS', 306, 'A'), (5, 'ORANGE', 165, 'A'), (5, 'BIM', 100, 'A'),
    (6, 'AIN SAISS', 306, 'A'), (6, 'BOULANGERIE', 285, 'A'), (6, 'NADAFA', 150, 'A'),
    (7, 'BIM', 300, 'A'), (7, 'NADAFA', 150, 'A'), (7, 'CITRIN', 26, 'A'),
    (8, 'BIM', 300, 'A'), (8, 'AIN SAISS', 306, 'A'), (8, 'NADAFA', 150, 'A'),
    (9, 'ORANGE', 165, 'A'), (9, 'NADAFA', 150, 'A'), (9, 'BIM', 300, 'A'), (9, 'COCA', 558, 'A'),
    (9, 'BAR MEN', 600, 'S'), (9, 'BAR MEN ALI', 600, 'S'), (9, 'FATIMA', 400, 'S'),
    (9, 'LAHCEN', 500, 'S'), (9, 'ABDELLAH', 600, 'S'),
    (10, 'AIN SAISS', 306, 'A'), (10, 'ORANGE', 165, 'A'), (10, 'NADAFA', 150, 'A'),
    (11, 'BOULANGERIE', 437, 'A'), (11, 'NADAFA', 150, 'A'), (11, 'BIM', 300, 'A'),
    (11, 'TAXE NADAFA 2026', 5000, 'F'), (11, 'BAR MEN ALI', 2000, 'S'),
    (12, 'AIN SAISS', 366, 'A'), (12, 'NADAFA', 150, 'A'), (12, 'ORANGE', 165, 'A'), (12, 'CAFE', 1660, 'A'),
    (13, 'BIM', 300, 'A'), (13, 'NADAFA', 150, 'A'), (13, 'ORANGE', 165, 'A'),
    (14, 'AIN SAISS', 306, 'A'), (14, 'NADAFA', 150, 'A'), (14, 'BIM', 300, 'A'),
    (15, 'BIM', 300, 'A'), (15, 'NADAFA', 150, 'A'), (15, 'OULMES', 156, 'A'),
    (16, 'AIN SAISS', 306, 'A'), (16, 'NADAFA', 150, 'A'), (16, 'ORANGE', 165, 'A'), (16, 'KHLII', 65, 'A'),
    (16, 'EAU + ELECTRICITE', 3785, 'F'),
    (16, 'BAR MEN', 600, 'S'), (16, 'BAR MEN ALI', 400, 'S'), (16, 'FATIMA', 400, 'S'),
    (16, 'LAHCEN', 500, 'S'), (16, 'ABDELLAH', 600, 'S'),
    (17, 'AIN SAISS', 306, 'A'), (17, 'NADAFA', 150, 'A'), (17, 'BIM', 300, 'A'),
    (18, 'ORANGE', 165, 'A'), (18, 'NADAFA', 150, 'A'), (18, 'BIM', 300, 'A'),
    (19, 'AIN SAISS', 306, 'A'), (19, 'NADAFA', 150, 'A'), (19, 'BIM', 200, 'A'), (19, 'CAFE', 1560, 'A'),
    (20, 'ORANGE', 165, 'A'), (20, 'NADAFA', 150, 'A'), (20, 'BIM', 300, 'A'), (20, 'FRUITS', 55, 'A'),
    (21, 'AIN SAISS', 306, 'A'), (21, 'NADAFA', 150, 'A'), (21, 'BIM', 100, 'A'),
    (22, 'BOULANGERIE', 340, 'A'), (22, 'NADAFA', 150, 'A'), (22, 'BIM', 150, 'A'), (22, 'COCA', 558, 'A'),
    (23, 'AIN SAISS', 306, 'A'), (23, 'NADAFA', 150, 'A'), (23, 'BIM', 150, 'A'), (23, 'ORANGE', 165, 'A'),
    (23, 'BAR MEN', 600, 'S'), (23, 'BAR MEN ALI', 400, 'S'), (23, 'FATIMA', 400, 'S'),
    (23, 'LAHCEN', 500, 'S'), (23, 'ABDELLAH', 600, 'S'),
    (24, 'OULMES', 156, 'A'), (24, 'NADAFA', 150, 'A'), (24, 'BIM', 300, 'A'),
    (24, 'AVANCE LOYER', 5000, 'V'),
    (25, 'AIN SAISS', 306, 'A'), (25, 'NADAFA', 150, 'A'), (25, 'ORANGE', 165, 'A'), (25, 'FRUITS', 100, 'A'),
    (26, 'AIN SAISS', 306, 'A'), (26, 'NADAFA', 150, 'A'), (26, 'BIM', 200, 'A'), (26, 'CAFE', 780, 'A'),
    (27, 'BOULANGERIE', 241, 'A'), (27, 'NADAFA', 150, 'A'), (27, 'BIM', 200, 'A'),
]

# --- Classement fournisseur -> categorie ------------------------------------
CATEGORIE_PAR_TIERS = {
    'AIN SAISS': 'Boissons', 'OULMES': 'Boissons', 'COCA': 'Boissons', 'CITRIN': 'Boissons',
    'ORANGE': 'Fruits & Légumes', 'FRUITS': 'Fruits & Légumes',
    'BIM': 'Épicerie & Divers', 'KHLII': 'Épicerie & Divers',
    'MECAFE': 'Café & Torréfaction', 'CAFE': 'Café & Torréfaction',
    'BOULANGERIE': 'Boulangerie',
    'NADAFA': 'Nettoyage',
    'TAXE NADAFA 2026': 'Taxes & Impôts',
    'EAU + ELECTRICITE': 'Eau & Électricité',
    'INTERNET': 'Internet & Télécom',
    'LOYER': 'Loyer', 'AVANCE LOYER': 'Loyer',
    'BAR MEN': 'Salaires', 'BAR MEN ALI': 'Salaires', 'FATIMA': 'Salaires',
    'MARWA': 'Salaires', 'LAHCEN': 'Salaires', 'ABDELLAH': 'Salaires', 'AHMED': 'Salaires',
    'BANQUE': 'Virement Banque', 'MOUHSSINE': 'Virement Associés', 'HAMID': 'Virement Associés',
}

CATEGORIES_RECETTE = [
    'Vente Caisse', 'Glovo Espèce', 'Glovo Carte', 'Terrasse', 'Autre Recette',
]
CATEGORIES_DEPENSE = [
    'Boissons', 'Café & Torréfaction', 'Boulangerie', 'Fruits & Légumes',
    'Épicerie & Divers', 'Nettoyage', 'Salaires', 'Loyer', 'Eau & Électricité',
    'Internet & Télécom', 'Taxes & Impôts', 'Abonnements', 'Entretien & Réparation',
    'Virement Banque', 'Virement Associés', 'Divers',
]
MODES = ['Espèce', 'Carte', 'Virement', 'Chèque']

# --- Historique mensuel 2025 (feuille Feuil1) -------------------------------
HISTO_2025 = [
    (1, 58097, 17550), (2, 47774, 15723), (3, 33629, 19497), (4, 63101, 24871),
    (5, 67697, 27947), (6, 57443, 23534), (7, 71948, 28425), (8, 78294, 31683),
    (9, 53925, 23474), (10, 52669, 22122), (11, None, None), (12, None, None),
]

# --- Charges fixes mensuelles (feuille CHARGE, colonnes E/F = budget retenu) -
CHARGES_FIXES = [
    ('Achats marchandise (café + eau + Salah)', 15000, 'Achats'),
    ('BIM (épicerie)', 5500, 'Achats'),
    ('Internet', 500, 'Internet & Télécom'),
    ('Eau + Électricité', 4000, 'Eau & Électricité'),
    ('Loyer', 17000, 'Loyer'),
]

# --- Personnel (feuille CHARGE, colonnes E/F) -------------------------------
# (nom, poste, salaire mensuel convenu ou None si inconnu)
# Les six premiers viennent de la feuille CHARGE ; les suivants sont les noms
# reellement payes dans les feuilles jour d’aout 2026, dont le salaire convenu
# ne figurait nulle part dans l’ancien classeur.
SALAIRES = [
    ('ALI', 'Barman', 2400),
    ('LATIFA 1', 'Service', 1800),
    ('LATIFA 2', 'Service', 1800),
    ('ABDELLAH', 'Serveur', 2000),
    ('HAMZA', 'Serveur', 2000),
    ('AHMED', 'Gérant', 7000),
    ('BAR MEN', 'Barman', None),
    ('BAR MEN ALI', 'Barman', None),
    ('FATIMA', 'Service', None),
    ('LAHCEN', 'Service', None),
    ('MARWA', 'Service', None),
]
ASSOCIES_REMUNERATION = ('Hamid + Mouhcine (associés)', 16000)

# --- Prix de vente et volumes (feuille CHARGE, colonnes H/I/J) --------------
# Bloc 1 = volumes journaliers, bloc 2 = volumes mensuels constates
PRODUITS = [
    ('Café Noir', 13, 100, 357),
    ('Café Crème', 13, 20, 128),
    ('Petit Déjeuner', 20, 8, 0),
    ('Thé', 12, 10, 4),
]

# --- Abonnements (feuille Feuil2) -------------------------------------------
ABONNEMENTS = [
    ('BEIN SPORT', 600), ('REMO', 3800), ('NEGOZAR', 700),
    ('BEIN CAN 2025', 500), ('CAFE DUBOIS', 135),
]

# --- Comptes associes (feuille DEPENSE_CAFE) --------------------------------
ACQUISITION = {
    'montant_fonds': 795000,
    'part_par_associe': 397500,
    'verse_hamid': 397738,
    'verse_mouhssine': 404700,
    'paradise': 60000,
}
INVEST_HAMID = 148042
INVEST_MOUHSSINE = 298821

# --- Detail : acquisition du fonds de commerce (DEPENSE_CAFE l.10-70) ---
ACQUISITION_DETAIL = [
    ('ESPECE(HAMID:10000/MH:90000) HOURI', 10000, 90000),
    ('ESPECE', 20000, 0),
    ('ESPECE', 30000, 0),
    ('VIREMENT 1', 0, 20000),
    ('VIREMENT 2', 0, 20000),
    ('cheque ahmed', 0, 20000),
    ('LOYER NOVEMBRE', 0, 15000),
    ('NOTAIRE', 0, 0),
    ('CHEQUE', 0, 60000),
    ('cheque ahmed', 0, 20000),
    ('ESPECE  RADIATION SAADAOUI', 0, 0),
    ('ESPECE LE 15/06/24', 25000, 0),
    ('EAU+ELEC', 7800, 0),
    ('ESPECE berkane', 0, 60000),
    ('CHEQUE HAMID(VIREMENT MOUHSSINE)', 0, 20000),
    ('43TAXE LE 22/02/2024-DEBIT BOISSON', 0, 3500),
    ('TAPISSIER CHAISE', 0, 11700),
    ('credit Paradise à12/07/2024', 150000, 0),
    ('ESPECE LE 15/07/24-ateliersaadaoui-', 5000, 0),
    ('ESPECE LE (BL PARADISE zenata)', 10500, 0),
    ('ESPECE LE 19/07/24-(BL PARADISE zenata)', 110, 0),
    ('ESPECE LE 22/07/24-ateliersaadaoui-', 30000, 0),
    ('ESPECE LE 31/07/24-(BL PARADISE zenata)', 681, 0),
    ('ESPECE LE 03/08/24-LABORATOIRE 17-', 10000, 0),
    ('ESPECE LE 08/08/24-(BL PARADISE zenata)', 531, 0),
    ('ESPECE LE 29/08/24-(BL PARADISE zenata)', 3400, 0),
    ('ESPECE LE 30/08/24-(BL PARADISE zenata)', 5926, 0),
    ('ESPECE LE 05/09/24-(BL PARADISE zenata)', 367, 0),
    ('ESPECE LE 06/09/24-(BL PARADISE zenata)', 225, 0),
    ('ESPECE LE 10/09/24-(BL PARADISE zenata)', 1140, 0),
    ('ESPECE LE 08/10/24-(BL PARADISE zenata)', 24120, 0),
    ('CHEQUE LE 08/10/24-(HAMID)', 20000, 0),
    ('ESPECE LE 12/10/24-(BL PARADISE zenata)', 1040, 0),
    ('ESPECE LE 14/10/24-(BL PARADISE zenata)', 4781, 0),
    ('ESPECE LE 28/10/24-(BL PARADISE zenata)', 285, 0),
    ('ESPECE LE 06/11/24-(BL PARADISE zenata)', 11600, 0),
    ('ESPECE LE 11/11/24-(BL PARADISE zenata)', 1650, 0),
    ('ESPECE LE 14/11/24-(BL PARADISE zenata)', 395, 0),
    ('ESPECE LE 19/11/24-(BL PARADISE )', 2453, 0),
    ('ESPECE LE 21/11/24-(BL PARADISE )', 263, 0),
    ('ESPECE 30/11/2024', 0, 10000),
    ('ESPECE LE 04/12/24-(BL PARADISE )', 767, 0),
    ('ESPECE LE 05/12/24-(BL PARADISE )', 190, 0),
    ('ESPECE LE 10/12/24-(BL PARADISE )', 70, 0),
    ('ESPECE LE 24/12/24-(BL PARADISE )', 240, 0),
    ('ESPECE LE 30/12/24-(BL PARADISE )', 571, 0),
    ('ESPECE LE 31/12/24-(BL PARADISE )', 4820, 0),
    ('ESPECE LE 03/01/25', 0, 20000),
    ('ESPECE LE 11/01/25-(BL PARADISE )', 192, 0),
    ('ESPECE LE 15/01/25-(BL PARADISE )', 3790, 0),
    ('ESPECE LE 16/01/25-(BL PARADISE )', 200, 0),
    ('', 673, 0),
    ('ESPECE LE 17/01/25-(BL PARADISE )', 1020, 0),
    ('ESPECE LE 28/01/25-(BL PARADISE )', 500, 0),
    ('ESPECE STATION AFRIQUIA 17', 5000, 0),
    ('ESPECE LE 04/02/25-(BL PARADISE )', 490, 0),
    ('ESPECE LE 19/02/25-(BL PARADISE )', 1848, 0),
    ('ESPECE LE 07/03/25-(BL PARADISE )', 100, 0),
    ('ESPECE LE 13/03/25 atelier 17', 0, 14500),
    ('ESPECE LE 07/04/25-(BL PARADISE )', 0, 0),
    ('ESPECE LE 18/04/25-(CAFE VICTOR )', 0, 20000),
]

# --- Detail : amenagement du Cafe Victor Hugo (DEPENSE_CAFE l.77-120) ---
AMENAGEMENT_DETAIL = [
    ('ACHAT STOCK', 7000, 0),
    ('SARF', 2000, 0),
    ('FRIGO ASTORIA', 9000, 0),
    ('MACHINE ASTORIA', 31000, 20000),
    ('CONGELATEUR ASWAK SALAM', 2500, 0),
    ('BEIN SPORT', 0, 3200),
    ('Travaux Eléctricien', 0, 3000),
    ('ACHAT TV', 0, 30500),
    ('HABILLAGE MURE 12000 DH', 12000, 0),
    ('HABILLAGE DERIERE COMPTOIRE 2300DH', 2300, 0),
    ('SUPPORT INOX', 1400, 0),
    ('Travaux Eléctricien SUPPLIMENT', 850, 0),
    ('LOYER decembre', 0, 12750),
    ('HABILLAGE COMPTOIRE+ENSIGNE', 5000, 34800),
    ('achat 7 pots', 2100, 0),
    ('NOTAIRE', 1000, 2000),
    ('achat diffuseur', 0, 5300),
    ('POSE ALUMINIUM', 13300, 0),
    ('achat SANITAIRE', 142, 1000),
    ('REMPLISSAGE 7 pots', 0, 6000),
    ('BACHE TERRASSE+CASQUETTE', 0, 32000),
    ('ACHAT PEINTREUR', 450, 0),
    ('MAIN D ŒUVRE PEINTRE', 0, 2500),
    ('MACHINE ASTORIA(CHEQUE)', 20000, 0),
    ('PAYE AUTORISATION CAFE', 0, 1000),
    ('AVANCE SALAIRE AHMED', 0, 5000),
    ('REPARATION TV 85', 0, 10000),
    ('MACHINE GLACE', 11000, 0),
    ('LAMPE BRICOMA', 0, 1571),
    ('CHARGEUR PARFUM', 0, 1200),
    ('telephone+tablette', 0, 3300),
    ('achat diffuseur', 0, 840),
    ('placard', 3000, 0),
    ('Jama3a taxe terrasse 11/02/2025', 0, 30000),
    ('Verre terrasse 13/05/2025', 18000, 0),
    ('bardage   mur espece 5000 dh', 0, 4960),
    ('table terasse  espece 4000 dh', 0, 3900),
    ('espece 3600 dh(1300 dh glace+2 table 1300 dh+600 najar)', 0, 3200),
    ('espece 3600 dh(1300 dh glace+2 table 1300 dh+600 najar)', 0, 0),
    ('bardage:12370 elect:898 dh cor:1545 dh trnp:500 dh MDCV:4500 dh avance chese:1000 dh', 0, 29700),
    ('reliquat les cheses', 0, 31100),
    ('TAXE nadafa 2024 haj', 0, 10000),
    ('Salaire depart ahmed', 0, 10000),
    ('RELIquat SAADAOUI', 6000, 0),
]
