# Café Victor Hugo — Système de caisse

Classeurs Excel de gestion pour le **Café Victor Hugo** (Mohammedia, Maroc).

**Fichier principal : [`CAISSE_VICTOR_HUGO_08-2026.xlsx`](CAISSE_VICTOR_HUGO_08-2026.xlsx)**

## Le principe

Un classeur = **un mois**. La saisie se fait comme dans l'ancien classeur —
**une feuille par jour** — mais tout ce qui suit se calcule seul.

```
   Saisie du soir  ─►  feuilles  01 … 31   (une par jour)
                            │
                            ▼
                   RÉCAP DU JOURNAL   ◄── les 4 réglages du mois
                            │
          ┌─────────────────┼─────────────────┐
          ▼                 ▼                 ▼
   TABLEAU DE BORD    CHARGES FIXES    SUIVI MENSUEL
```

## Les cinq parties du système

| Feuille | Rôle | Saisie |
|---|---|---|
| **01 … 31** | Une feuille de caisse par jour : 4 blocs de dépenses + recette du soir | ✅ tous les jours |
| **RÉCAP DU JOURNAL** | Le mois entier ligne par ligne, + les 4 réglages | Réglages |
| **TABLEAU DE BORD** | 8 indicateurs et 4 graphiques du mois | — |
| **CHARGES FIXES** | Budget par poste contre réalisé, salaires, seuil de rentabilité | Budgets |
| **SUIVI MENSUEL** | L'année mois par mois + rappel 2025 | Mois clos |

Une feuille `LISTES`, **masquée**, ne contient que les listes déroulantes
(fournisseurs, personnel…). Pour l'afficher : clic droit sur un onglet ▸
*Afficher* ▸ `LISTES`.

## La feuille d'un jour

Même disposition que l'ancien classeur, en quatre blocs côte à côte :

| Bloc | Contenu | Libellés |
|---|---|---|
| **ACHATS DU JOUR** | fournisseurs du jour | libres (liste déroulante) |
| **CHARGES FIXES** | loyer, eau/élec, internet, taxes… | **fixes** |
| **SALAIRES ET AVANCES** | les 11 personnes | **fixes** |
| **VIREMENTS ET DIVERS** | banque, associés, avance loyer… | **fixes** |

En bas : la **recette comptée** (caisse + Glovo) et le **solde du soir**, qui
devient automatiquement la caisse de la veille sur la feuille du lendemain.

> Les libellés des trois derniers blocs sont fixes **volontairement** : c'est ce
> qui permet à CHARGES FIXES de totaliser chaque poste sur les 31 jours.
> N'en changez pas l'ordre.

## Changer de mois

1. Faites une copie du fichier ;
2. dans **RÉCAP DU JOURNAL**, changez `Mois` (et `Année` si besoin) ;
3. mettez `Caisse au 1er du mois` au solde de fin du mois précédent ;
4. effacez les montants des feuilles 01 à 31.

Les dates, les jours de la semaine et le nombre de jours du mois suivent seuls
(testé : en passant à septembre, la feuille 31 se vide d'elle-même).

## Données reprises de l'ancien classeur

| Source | Contenu | Destination |
|---|---|---|
| 27 feuilles jour d'août 2026 | 145 lignes | feuilles 01 à 27 |
| `SUIVI` | totaux de contrôle | (vérification) |
| `Feuil1` | 10 mois de 2025 | SUIVI MENSUEL, rappel 2025 |
| `CHARGE` | budgets et salaires convenus | CHARGES FIXES |
| `DEPENSE_CAFE` | 105 lignes d'apports | `ANNEXE_ASSOCIES.xlsx` |
| `PVC` | commande d'irrigation | `ANNEXE_ASSOCIES.xlsx` |

### Contrôle de reprise

Chaque colonne a été rapprochée de la feuille `SUIVI` d'origine :

| Poste | Ancien classeur | Nouveau | Écart |
|---|---|---|---|
| Recette | 54 937 DH | 54 937 DH | 0 |
| Achats | 24 529 DH | 24 529 DH | 0 |
| Charges fixes | 8 785 DH | 8 785 DH | 0 |
| Salaires | 12 400 DH | 12 400 DH | 0 |
| Virements | 5 000 DH | 5 000 DH | 0 |
| **Total dépenses** | **50 714 DH** | **50 714 DH** | **0** |
| Caisse en fin de mois | 4 223 DH | 4 223 DH | 0 |

La chaîne de caisse jour après jour est également identique
(1 414 → 258 → 2 099 → 2 678 …).

## Annexe

**[`ANNEXE_ASSOCIES.xlsx`](ANNEXE_ASSOCIES.xlsx)** — *ne fait pas partie du
système quotidien.* Archive des apports de Hamid et Mouhssine (achat du fonds
de commerce et aménagement du café, 105 lignes, 1 249 301 DH au total) et de la
commande de matériel d'irrigation. Conservée pour ne pas perdre ces données.

## Code de couleurs

- 🟡 **Fond jaune, texte bleu** — à remplir
- ⬜ **Fond blanc/crème** — calculé, ne pas modifier
- 🟢 **Vert** — recette, résultat positif
- 🔴 **Rouge** — dépense, résultat négatif, budget dépassé

## Régénérer les classeurs

```bash
pip install openpyxl
python3 src/build.py          # CAISSE_VICTOR_HUGO_08-2026.xlsx
python3 src/build_annexe.py   # ANNEXE_ASSOCIES.xlsx
```

- `src/data_source.py` — les données reprises de l'ancien classeur
- `src/theme.py` — charte graphique (palette, polices, briques de mise en page)
- `src/build.py` — les 36 feuilles du système
- `src/build_annexe.py` — l'annexe
