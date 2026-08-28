# Café Victor Hugo — Système de gestion

Classeur Excel de gestion complet pour le **Café Victor Hugo** (Mohammedia, Maroc).

**Fichier livré : [`CAFE_VICTOR_HUGO_SYSTEME.xlsx`](CAFE_VICTOR_HUGO_SYSTEME.xlsx)**

## Le principe

L'ancien classeur (`SUIVICAISSE_VH2026_08.xlsx`) demandait **une feuille par jour** —
27 feuilles pour un mois — et il fallait recopier chaque total à la main dans la
feuille de suivi. Une erreur de recopie passait inaperçue.

Le nouveau système inverse la logique : **on saisit une seule fois, dans une seule
feuille**, et les onze autres se calculent seules.

```
                    ┌──────────────┐
   Saisie du soir → │   JOURNAL    │  1 ligne = 1 opération
                    └──────┬───────┘
                           │  SUMIFS
       ┌───────────┬───────┴────┬─────────────┬──────────────┐
       ▼           ▼            ▼             ▼              ▼
  RECAP JOUR  TABLEAU DE   SUIVI       CHARGES FIXES    SALAIRES
   (31 j.)      BORD       MENSUEL     (budget/réel)   (versements)
```

## Les feuilles

| Feuille | Rôle | Saisie |
|---|---|---|
| **ACCUEIL** | Mode d'emploi en 3 étapes, chiffres du mois, sommaire | — |
| **TABLEAU DE BORD** | 8 indicateurs + 4 graphiques du mois | — |
| **JOURNAL** | Une ligne par opération. **La seule feuille à remplir chaque jour** | ✅ |
| **RECAP JOUR** | Les 31 jours du mois : recette, dépenses par type, caisse cumulée | — |
| **SUIVI MENSUEL** | 2025 + 2026 mois par mois, avec graphiques de tendance | — |
| **CHARGES FIXES** | Budget par poste contre réalisé, coût de fonctionnement/jour | Budget |
| **SALAIRES** | Personnel, salaire convenu, versé, reste à verser | Salaires |
| **MARGES PRODUITS** | Prix de vente, coût matière, marge par boisson | Coûts |
| **ASSOCIÉS** | Apports Hamid / Mouhssine : fonds de commerce + aménagement | — |
| **PARAMÈTRES** | Mois analysé, caisse de départ, listes de référence | ✅ |
| **ARCHIVE** | Données sans rapport avec le café, conservées | — |

## Changer de mois

Un seul geste : dans **PARAMÈTRES**, modifiez le numéro du mois (cellule jaune `C13`).
Le RECAP JOUR, le TABLEAU DE BORD et les CHARGES FIXES suivent immédiatement.

## Données reprises de l'ancien classeur

| Source | Contenu | Destination |
|---|---|---|
| 27 feuilles jour d'août 2026 | 145 opérations | JOURNAL |
| `SUIVI` | Totaux de contrôle | (vérification) |
| `Feuil1` | 10 mois de 2025 | SUIVI MENSUEL |
| `CHARGE` | 9 postes de charges, 11 personnes, 4 prix de vente | CHARGES FIXES / SALAIRES / MARGES |
| `DEPENSE_CAFE` | 105 lignes d'apports des associés | ASSOCIÉS |
| `Feuil2` | 5 abonnements | CHARGES FIXES |
| `PVC` | Commande d'irrigation (hors café) | ARCHIVE |

### Contrôle de reprise

Les totaux du nouveau classeur ont été rapprochés de la feuille `SUIVI` d'origine :

| Poste | Ancien classeur | Nouveau | Écart |
|---|---|---|---|
| Recettes août 2026 | 54 937 DH | 54 937 DH | 0 |
| Dépenses août 2026 | 50 714 DH | 50 714 DH | 0 |
| Apports Hamid (fonds) | 397 738 DH | 397 738 DH | 0 |
| Apports Mouhssine (fonds) | 404 700 DH | 404 700 DH | 0 |
| Aménagement Hamid | 148 042 DH | 148 042 DH | 0 |
| Aménagement Mouhssine | 298 821 DH | 298 821 DH | 0 |

## Deux écarts volontaires avec l'ancien classeur

1. **`AVANCE LOYER` (5 000 DH, 24/08)** était classée en « Virement » ; elle est
   désormais dans la catégorie **Loyer**, qui correspond à sa nature réelle. Le
   total des dépenses est inchangé (50 714 DH).
2. **Cinq personnes payées en août** (BAR MEN, BAR MEN ALI, FATIMA, LAHCEN,
   MARWA) ne figuraient dans aucune liste de salaires. Elles ont été ajoutées à
   la feuille SALAIRES ; leur salaire convenu reste à renseigner (cellules jaunes).

## Contrôles automatiques intégrés

- **Colonne CONTRÔLE du JOURNAL** — passe au vert (`OK`) quand la ligne est
  complète, signale `Catégorie ?`, `Montant ?` ou `Type ?` sinon.
- **Colonne AUTRES du RÉCAP JOUR** — doit rester vide. Si un montant y apparaît,
  une dépense a été saisie avec une catégorie hors des groupes connus.
- **« Hors liste » dans SALAIRES** — montre les salaires versés à quelqu'un
  qui n'est pas dans le tableau du personnel.
- **État HORS BUDGET dans CHARGES FIXES** — un poste dépensé sans budget prévu.

## Code de couleurs

- 🟡 **Fond jaune, texte bleu** — à remplir par vous
- ⬜ **Fond blanc/crème** — calculé automatiquement, ne pas modifier
- 🟢 **Vert** — résultat positif, objectif atteint
- 🔴 **Rouge** — résultat négatif, budget dépassé

## Régénérer le classeur

```bash
pip install openpyxl
python3 src/build.py
```

- `src/data_source.py` — toutes les données reprises de l'ancien classeur
- `src/theme.py` — charte graphique (palette, polices, briques de mise en page)
- `src/build.py` — construction des 11 feuilles
