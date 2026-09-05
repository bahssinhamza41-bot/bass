# -*- coding: utf-8 -*-
"""Donnees reprises des documents fournis par la societe.

Trois sources :
  - PARADISEHAMZA.xlsx   : la feuille de caisse du 01/09/2026 et le grand
                           registre des cheques et effets (colonnes H a N) ;
  - le dossier de presentation (PPTX) : identite de la societe, principaux
                           fournisseurs, chiffres cles janvier - juillet 2026 ;
  - le logo, extrait du meme dossier (paradise/assets/).
"""

# --- Identite de la societe (dossier de presentation, page 2) ---------------
SOCIETE = {
    'nom':      'PARADISE ALUMINIUM',
    'forme':    'SARL',
    'adresse':  'Essalam 2, Tranche H n° 92, Sidi Moumen, Ahl Loughlam — Casablanca',
    'tel':      '07 03 09 02 90',
    'email':    'Alumparadise@gmail.com',
    'rc':       '300049',
    'patente':  '37151101',
    'cnss':     '9908700',
    'if':       '15160969',
    'ice':      '000083422000043',
    'banque':   'SGMB — Société Générale Maroc',
    'rib':      '022 780 0003730028629315 74',
}

# --- Soldes d'ouverture (feuille 01-09-26, bloc CAISSE SYSTEME) -------------
CAISSE_PARADISE_OUVERTURE = 3270
CAISSE_ZENATA_OUVERTURE   = 18188
AUTRE_CAISSE_OUVERTURE    = -58814.02      # bloc AUTRE CAISSE, ligne « Précédent »

# --- Fournisseurs -----------------------------------------------------------
# Les dix premiers viennent du graphique « Principaux fournisseurs » du
# dossier de presentation ; les suivants du registre des reglements.
FOURNISSEURS = [
    'ALUMINIUM ELMARWA', 'FOMALU', 'LAMASTORE', 'FATALUM', 'IDRALUM',
    'INDEC', 'DIMA BOIS', 'CASAPRAL', 'LCVR', 'ALUMA',
    'HAJ OMAR', 'GLOBALE VERRE', 'SAID PROFILE', 'ALUMINIUM DU MAROC',
    'QUINCAILLERIE', 'VITRAGE', 'PANNEAUX BOIS', 'TRANSPORT',
    'CARBURANT', 'DIVERS',
]

# --- Postes de charges fixes (libelles fixes des feuilles jour) -------------
# (libelle, budget mensuel indicatif en DH)
CHARGES_FIXES = [
    ('LOYER DU LOCAL',            2500),
    ('EAU + ÉLECTRICITÉ',         1200),
    ('TÉLÉPHONE & INTERNET',       500),
    ('CNSS',                      4503),
    ('TVA & IMPÔTS',              6000),
    ('ASSURANCES',                1000),
    ('COMPTABLE & HONORAIRES',     700),
    ('CARBURANT & VÉHICULE',      2400),
    ('TRANSPORT & LIVRAISON',     1500),
    ('ENTRETIEN & RÉPARATION',     800),
    ('FOURNITURES DE BUREAU',      300),
    ('AUTRE CHARGE FIXE',            0),
]

# --- Personnel (libelles fixes du bloc salaires) ----------------------------
# (nom, salaire mensuel convenu en DH)
PERSONNEL = [
    ('HOUSSEIN',        8000),
    ('AYOUB',           5000),
    ('OUVRIER 1',          0),
    ('OUVRIER 2',          0),
    ('AVANCE / PRIME',     0),
    ('AUTRE SALAIRE',      0),
]

# --- Achats du 01/09/2026 (feuille 01-09-26, bloc fournisseurs) -------------
# (BL, fournisseur, net a payer, espece, cheque, n° cheque, effet, n° effet)
ACHATS_J1 = [
    ('', 'TRANSPORT', 60,  60,  0, '', 0, ''),
    ('', 'DIVERS',    210, 210, 0, '', 0, ''),
]

# --- Chiffres cles janvier - juillet 2026 (dossier de presentation, page 4) -
REFERENCE_2026 = {
    'mois':                 7,
    'ca_facture':           1434929,
    'achats_ttc':           2382477,
    'versements_especes':   1386700,
    'clients':              34,
    'fournisseurs':         19,
    'flux_total':           2821629,
}

# --- Registre des cheques, effets et engagements ---------------------------
# Repris tel quel des colonnes H a N de PARADISEHAMZA.xlsx.
# (type de reglement, beneficiaire / objet, n° cheque ou effet,
#  montant restant du, date, montant deja paye, observation)
ECHEANCIER = [
    ('MOUHSSINE/SIIT', None, None, 65000, '2024-01-04', None, None),
    ('HAMID', None, None, 95000, '2024-01-05', None, None),
    ('MOUHSSINE/SIIT', None, None, 40000, '2024-05-30', None, None),
    ('MOUHSSINE/SIIT', None, None, 10600, '2024-05-30', None, '30000 DH VIREMENT LE 15/05/2024: PAYE 10000 DH+15000'),
    ('MOUHSSINE/GASS', None, None, 20000, '2024-05-30', None, '30001 DH VIREMENT LE 17/05/2024: PAYE 10000 DH'),
    ('MOUHSSINE/SIIT', None, None, 45000, '2024-05-30', None, '75000 DH VIREMENT : PAYE 30000 DH'),
    ('HOUSSEIN', 'CHARGE TRAVAUX', None, 120000, '2024-05-30', None, None),
    ('HAMID-COMPTABLE', None, None, 15000, '2024-05-30', None, None),
    ('HAMID-CLIENT', None, None, 12000, '2024-07-15', None, None),
    ('HOUSSEIN', 'Fatalum', '9402710', 12900, '2024-06-08', None, '20000: AVANCE 7100 DH'),
    ('Houssein', 'FATALUM', '9955189', 15000, '2024-06-12', None, 'PAYE'),
    ('Houssein', 'Fatalum', '9955194', 3000, '2024-07-02', None, '32000 DH: PAYE 27000 DH+1000 DH+1000'),
    ('XXXXX', 'SAID PROFILE', None, 20800, '2025-03-16', None, None),
    ('Chéque garantie(350000 dh)', 'Fatalum', '6784668', 269000, None, None, None),
    ('VIREMENT LOYER', 'LOYER  LOCAL MOIS 04/25', None, 2500, '2025-05-03', None, None),
    ('VIREMENT LOYER', 'LOYER  LOCAL MOIS 05/25', None, 2500, '2025-06-05', None, None),
    ('VIREMENT LOYER', 'LOYER  LOCAL MOIS 06/25', None, 2500, '2025-07-05', None, None),
    ('Virement-Houssein', 'Salaire mois 07/25', None, 8000, '2025-08-10', None, None),
    ('VIREMENT LOYER', 'LOYER  LOCAL MOIS 07/25', None, 2500, '2025-08-10', None, None),
    ('Virement-Houssein', 'Salaire mois 08/25', None, 8000, '2025-09-10', None, None),
    ('VIREMENT LOYER', 'LOYER  LOCAL MOIS 08/25', None, 2500, '2025-09-10', None, None),
    ('Virement-Houssein', 'Salaire mois 09/25', None, 8000, '2025-10-08', None, None),
    ('VIREMENT LOYER', 'LOYER  LOCAL MOIS 09/25', None, 2500, '2025-10-08', None, None),
    ('Virement-Houssein', 'Salaire mois 10/25', None, 8000, '2025-11-10', None, None),
    ('VIREMENT LOYER', 'LOYER  LOCAL MOIS 10/25', None, 2500, '2025-11-10', None, None),
    ('Virement-Houssein', 'Salaire mois 11/25', None, 8000, '2025-12-10', None, None),
    ('VIREMENT LOYER', 'LOYER  LOCAL MOIS 11/25', None, 2500, '2025-12-10', None, None),
    ('Virement-Houssein', 'LOYER  LOCAL MOIS 12/25', None, 8000, '2026-01-10', None, None),
    ('VIREMENT LOYER', 'Salaire mois 12/25', None, 2500, '2026-01-10', None, None),
    ('Virement-Houssein', 'Salaire mois 01/26', None, 8000, '2026-02-10', None, None),
    ('VIREMENT LOYER', 'LOYER  LOCAL MOIS 01/26', None, 2500, '2026-02-10', None, None),
    ('VIREMENT LOYER', 'LOYER  LOCAL MOIS 04/26', None, 2500, '2026-05-10', None, None),
    ('VIREMENT LOYER', 'LOYER  LOCAL MOIS 05/26', None, 2500, '2026-06-10', None, None),
    ('effet', 'HAJ OMAR', '0137653', 1, '2026-07-07', 30000, None),
    ('VIREMENT LOYER', 'LOYER  LOCAL MOIS 06/26', None, 2500, '2026-07-10', None, None),
    ('Chéque', 'IDRALUM', '8387627', 20000, '2026-07-20', None, None),
    ('Chéque', 'IDRALUM', '8387628', 10000, '2026-07-27', None, None),
    ('effet', 'HAJ OMAR', '0137659', 0, '2026-08-02', 29000, 'PAYE 29000 DH'),
    ('Chéque', 'HAJ OMAR', '0286165', 0, '2026-08-03', 25400, 'PAYE 25400 DH'),
    ('Chéque', 'CITRWEN', '0286177', 0, '2026-08-04', 19460.1, 'PAYE 19460,1 DH'),
    ('Chéque', 'CITRWEN IMATRICULATION', '0286174', 0, '2026-08-04', 4610, 'PAYE 4610 DH'),
    ('Chéque', 'Fatalum', '0286173', 0, '2026-08-04', 17572, 'PAYE 17572 DH'),
    ('effet', 'HAJ OMAR', '0137660', 0, '2026-08-07', 29000, 'PAYE 29000 DH'),
    ('Virement-AYOUB', 'Salaire mois 07/26', None, 5000, '2026-08-08', None, None),
    ('Virement-Houssein', 'Salaire mois 07/26', None, 0, '2026-08-08', 8000, 'PAYE 8000 DH'),
    ('VIREMENT LOYER', 'LOYER  LOCAL MOIS 07/26', None, 2500, '2026-08-08', None, None),
    ('Prélevement', 'CNSS', None, 0, '2026-08-08', 4503.03, 'PAYE 4503,03 DH'),
    ('Chéque', 'Fatalum', '0286172', 0, '2026-08-09', 15000, 'PAYE 15000 DH'),
    ('Chéque', 'XXXX', '0286182', 0, '2026-08-10', 16372, 'PAYE 16372 DH'),
    ('effet', 'HAJ OMAR', '0137661', 0, '2026-08-12', 30400, 'PAYE 30400 DH'),
    ('Chéque', 'Fatalum', '0286185', 0, '2026-08-13', 22155, 'PAYE22155 DH'),
    ('effet', 'HAJ OMAR', '0750928', 0, '2026-08-14', 28000, 'PAYE 28000 DH'),
    ('Chéque', 'ASSURANCE HASSANIA', '286191', 0, '2026-08-14', 933, 'PAYE 933 DH'),
    ('effet', 'HAJ OMAR', '0137662', 0, '2026-08-15', 17500, 'PAYE 17500 DH'),
    ('Chéque', 'IDRALUM', '0286155', 20000, '2026-08-15', None, None),
    ('Chéque', 'Fatalum', '0286171', 0, '2026-08-16', 15000, 'PAYE 15000 DH'),
    ('Chéque', 'Fatalum', '0286192', 0, '2026-08-17', 25274, 'PAYE 25274 DH'),
    ('effet', 'HAJ OMAR', '0750929', 0, '2026-08-18', 28000, 'PAYE 28000 DH'),
    ('Chéque', 'IDRALUM', '0286156', 20000, '2026-08-20', None, None),
    ('Chéque', 'Fatalum', '0286184', 0, '2026-08-20', 9300, 'PAYE 9300 DH'),
    ('effet', 'HAJ OMAR', '0750930', 0, '2026-08-21', 28241, 'PAYE 28241 DH'),
    ('Prélevement', 'TATA VIVALIS', None, 0, '2026-08-21', 2372.62, 'PAYE2372,62 DH'),
    ('effet', 'INDEC', '0137663', 0, '2026-08-23', 14580, 'PAYE  14580DH'),
    ('Chéque', 'Fatalum', '286193', 0, '2026-08-24', 27000, 'PAYE  27000DH'),
    ('effet', 'HAJ OMAR', '0750931', 0, '2026-08-25', 28500, 'PAYE 28500DH'),
    ('effet', 'HAJ OMAR', '0750932', 0, '2026-08-27', 28500, 'PAYE 28500DH'),
    ('Chéque', 'Fatalum', '0286194', 0, '2026-08-27', 20000, 'PAYE 20000DH'),
    ('effet', 'HAJ OMAR', '0750933', 0, '2026-08-29', 29000, 'PAYE 29000DH'),
    ('Prélevement', 'TVA', None, 0, '2026-08-30', 6000, 'PAYE 6000DH'),
    ('Chéque-H2', 'GLOBALE VERRE', '4888250', 0, '2026-09-01', 12500, 'PAYE12500DH'),
    ('effet', 'HAJ OMAR', '0750934', 30277, '2026-09-02', None, None),
    ('Chéque', 'IDRALUM', '0286157', 20000, '2026-09-02', None, None),
    ('effet', 'HAJ OMAR', '0750935', 25452, '2026-09-05', None, None),
    ('Chéque', 'IDRALUM', '0286158', 20000, '2026-09-06', None, None),
    ('effet', 'HAJ OMAR', '0750938', 17000, '2026-09-08', None, None),
    ('effet', 'INDEC', '0137665', 4908, '2026-09-08', None, None),
    ('effet', 'HAJ OMAR', '0750939', 17865, '2026-09-11', None, None),
    ('Chéque', 'ASSURANCE HASSANIA', '0286188', 1000, '2026-09-12', None, None),
    ('effet', 'HAJ OMAR', '0750940', 20700, '2026-09-15', None, None),
    ('effet', 'HAJ OMAR', '0750941', 21038, '2026-09-18', None, None),
    ('Chéque', 'IDRALUM', '0286164', 20000, '2026-09-23', None, None),
    ('Chéque', 'IDRALUM', '0286163', 20000, '2026-09-28', None, None),
    ('effet', 'HAJ OMAR', '0750946', 20000, '2026-09-28', None, None),
    ('Chéque', 'IDRALUM', '0286162', 20000, '2026-09-30', None, None),
    ('effet', 'HAJ OMAR', '0750947', 15200, '2026-10-03', None, None),
    ('Chéque', 'IDRALUM', '0286181', 20000, '2026-10-05', None, None),
    ('effet', 'HAJ OMAR', '0750948', 20000, '2026-10-06', None, None),
    ('Chéque', 'IDRALUM', '0286179', 18000, '2026-10-09', None, None),
    ('effet', 'HAJ OMAR', '0750949', 18000, '2026-10-10', None, None),
    ('Chéque', 'ASSURANCE HASSANIA', '0286186', 1000, '2026-10-12', None, None),
    ('effet', 'HAJ OMAR', '0750950', 18400, '2026-10-13', None, None),
    ('Chéque', 'IDRALUM', '0286183', 17000, '2026-10-14', None, None),
    ('effet', 'HAJ OMAR', '0750952', 24000, '2026-10-17', None, None),
    ('effet', 'HAJ OMAR', '0750953', 24000, '2026-10-22', None, None),
    ('effet', 'HAJ OMAR', '0750954', 24000, '2026-10-24', None, None),
    ('COMPTABLE', 'FRAIS DE CONSTITUTION H2', None, 2000, None, None, '4000: AVANCE 2000 DH'),
    ('COMPTABLE', 'TENU DE COMTABILITE PARADISE 2024', None, 900, None, None, '8400: AVANCE 2500 DH+3000DH+2000 DH'),
    ('COMPTABLE', 'TENU DE COMTABILITE PARADISE 2025', None, 8400, None, None, None),
    ('COMPTABLE', 'TENU DE COMTABILITE H2 2025', None, 2000, None, None, None),
]
