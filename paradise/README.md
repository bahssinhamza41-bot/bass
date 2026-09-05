# Paradise Aluminium — Système de caisse

Classeurs Excel de gestion pour **PARADISE ALUMINIUM SARL**
(négoce, façonnage et pose d'aluminium — Sidi Moumen, Casablanca).

**Classeurs du mois :**
- [`CAISSE_PARADISE_ALUMINIUM_09-2026.xlsx`](../CAISSE_PARADISE_ALUMINIUM_09-2026.xlsx)
  — septembre 2026, ouvert sur les soldes du 01/09
- [`CAISSE_PARADISE_ALUMINIUM_10-2026.xlsx`](../CAISSE_PARADISE_ALUMINIUM_10-2026.xlsx)
  — octobre 2026, vide, prêt à remplir

## Le principe

Un classeur = **un mois**. La saisie se fait **une feuille par jour**, comme
sur la feuille de caisse actuelle. Tout ce qui suit se calcule seul.

```
   Saisie du soir  ─►  feuilles  01 … 31   (une par jour)
                            │
                            ▼
                   RÉCAP DU JOURNAL   ◄── les réglages du mois
                            │
      ┌──────────┬──────────┼──────────┬──────────┐
      ▼          ▼          ▼          ▼          ▼
  TABLEAU    FOURNIS-   ÉCHÉANCIER  CHARGES    SUIVI
  DE BORD    SEURS                  FIXES      MENSUEL
```

## Les sept parties du système

| Feuille | Rôle | Saisie |
|---|---|---|
| **01 … 31** | La journée : achats fournisseurs, charges, salaires, dépenses, encaissements, les deux caisses | ✅ tous les jours |
| **RÉCAP DU JOURNAL** | Le mois entier ligne par ligne, + les réglages du mois | Réglages |
| **TABLEAU DE BORD** | 8 indicateurs et 4 graphiques | — |
| **FOURNISSEURS** | Achats, règlements et crédit, fournisseur par fournisseur | — |
| **ÉCHÉANCIER** | Les chèques et les effets, par date d'échéance | ✅ à chaque chèque |
| **CHARGES FIXES** | Budget par poste contre réalisé, salaires | Budgets |
| **SUIVI MENSUEL** | L'année mois par mois + rappel du dossier bancaire | Mois clos |

Une feuille `LISTES`, **masquée**, ne contient que les listes déroulantes
(fournisseurs, origines d'encaissement). Pour l'afficher : clic droit sur un
onglet ▸ *Afficher* ▸ `LISTES`.

## La feuille d'un jour

Tout tient sur un écran — et sur une page à l'impression.

| Bloc | Contenu | Libellés |
|---|---|---|
| **ACHATS FOURNISSEURS DU JOUR** | BL, fournisseur, net à payer, et le règlement : espèce / chèque (n°) / effet (n°). Ce qui reste devient le **crédit fournisseur** | libres (liste déroulante) |
| **CHARGES FIXES DU JOUR** | loyer, eau/élec, CNSS, TVA, assurances, comptable… | **fixes** |
| **SALAIRES ET AVANCES** | les personnes | **fixes** |
| **DÉPENSES TRAVAUX ET DIVERS** | chantiers, fournitures, imprévus | libres |
| **ENCAISSEMENTS DU JOUR** | clients et apports, en espèce ou par chèque/virement | libres (liste déroulante) |

En bas, côte à côte : **CAISSE PARADISE**, **CAISSE ZENATA**,
**CONTRÔLE ET AUTRE CAISSE** et le **RÉSUMÉ DU JOUR**.

- La **caisse PARADISE** se remplit presque seule : elle prend le solde de la
  veille, ajoute les encaissements en espèce, et retire les achats réglés en
  espèce, les charges, les salaires et les dépenses du jour.
- La **caisse ZENATA** se tient à la main, ligne par ligne, sur le même modèle.
- Chaque caisse demande le soir les **espèces réellement comptées** : l'**écart**
  avec le solde théorique apparaît tout seul, en rouge s'il n'est pas nul.
- Le solde du soir devient automatiquement le solde de la veille sur la feuille
  du lendemain.

> Les libellés des charges fixes et des salaires sont fixes **volontairement** :
> c'est ce qui permet à CHARGES FIXES de totaliser chaque poste sur les 31 jours.
> N'en changez pas l'ordre. Pour ajouter un poste, servez-vous de la ligne
> « AUTRE CHARGE FIXE » ou du bloc « dépenses travaux et divers ».

## L'échéancier

C'est la feuille propre à ce métier : plus de 99 % des règlements fournisseurs
passent par chèque et effet. Chaque ligne porte une date d'échéance, un montant
restant dû et un montant déjà payé ; la colonne **ÉTAT** se recalcule seule à
chaque ouverture du classeur :

| État | Sens | Couleur |
|---|---|---|
| **PAYÉ** | plus rien à débiter | vert |
| **EN RETARD** | l'échéance est passée et le montant reste dû | rouge |
| **À ÉCHOIR ≤ 30 J** | à provisionner dans le mois | ambre |
| **À VENIR** | au-delà de 30 jours | neutre |
| **SANS DATE** | engagement sans échéance (chèque de garantie…) | gris |

Quand un chèque est débité : portez le montant dans *déjà payé* et ramenez le
*restant dû* à zéro. La ligne passe au vert et les quatre compteurs du haut
suivent. Un filtre est posé sur les en-têtes pour trier par échéance ou par
bénéficiaire.

## Créer le classeur d'un nouveau mois

Le plus simple — régénérer un classeur vide :

```bash
python3 paradise/src/build.py --mois 11 --annee 2026 \
    --caisse-paradise <solde PARADISE fin octobre> \
    --caisse-zenata   <solde ZENATA fin octobre> \
    --autre-caisse    <solde autre caisse fin octobre>
```

Ou à la main, sans Python :

1. faites une copie d'un classeur existant ;
2. dans **RÉCAP DU JOURNAL**, changez `Mois` (et `Année` si besoin) ;
3. mettez les trois soldes d'ouverture aux soldes de fin du mois précédent ;
4. effacez les montants des feuilles 01 à 31 ;
5. sur **ÉCHÉANCIER**, gardez les lignes encore dues et supprimez celles qui
   sont soldées si la liste devient trop longue.

Les dates, les jours de la semaine et le nombre de jours suivent seuls. Un mois
de moins de 31 jours neutralise les feuilles en trop : la feuille affiche
« — ce jour n'existe pas dans le mois — » et sa ligne disparaît du récap.

Le **report de caisse** est le seul lien entre deux mois : le solde du soir du
dernier jour devient le solde d'ouverture du mois suivant.

## Données reprises des documents de la société

| Source | Contenu | Destination |
|---|---|---|
| `PARADISEHAMZA.xlsx`, feuille `01-09-26` | caisses ZENATA 18 188 DH et PARADISE 3 270 DH | soldes d'ouverture de septembre |
| idem, bloc fournisseurs | les 2 lignes d'achats du 01/09 (270 DH) | feuille `01` |
| idem, bloc `AUTRE CAISSE` | solde précédent −58 814,02 DH | autre caisse au 1er |
| idem, colonnes H à N | 99 lignes de chèques, effets et engagements | `ÉCHÉANCIER` |
| Dossier de présentation (PPTX) | identité, principaux fournisseurs, chiffres janvier – juillet 2026 | `LISTES`, `SUIVI MENSUEL` |
| Dossier de présentation (PPTX) | le logo | `assets/`, `TABLEAU DE BORD` |

### Contrôle de reprise

| Poste | Ancien classeur | Nouveau | Écart |
|---|---|---|---|
| Caisse réelle au 01/09 (ZENATA + PARADISE) | 21 188 DH | 21 188 DH | 0 |
| Achats du 01/09 | 270 DH | 270 DH | 0 |
| Total restant dû (chèques, effets, engagements) | 1 380 441 DH | 1 380 441 DH | 0 |

Les deux premiers se lisent sur la feuille `01`, le troisième en haut de
`ÉCHÉANCIER`.

> Les lignes de 2024 et 2025 du registre (avances des associés, travaux,
> honoraires) ressortent en **EN RETARD** : elles sont datées et toujours dues.
> C'est bien ce que dit le classeur d'origine. Le filtre de la colonne
> *type de règlement* permet de ne regarder que les chèques et les effets.

## Code de couleurs

- 🟡 **Fond jaune, texte bleu** — à remplir
- ⬜ **Fond blanc/gris clair** — calculé, ne pas modifier
- 🔵 **Bleu marine** — totaux et soldes
- 🟢 **Vert** — encaissement, échéance payée
- 🔴 **Rouge** — crédit fournisseur, retard, écart de caisse, budget dépassé

## Régénérer les classeurs

```bash
pip install openpyxl pillow

# septembre 2026, avec les données reprises du classeur d'origine
python3 paradise/src/build.py --mois 9 --annee 2026

# octobre 2026 : pas de journée reprise, mais l'échéancier suit
python3 paradise/src/build.py --mois 10 --annee 2026 \
    --caisse-paradise 3000 --caisse-zenata 18188 --autre-caisse -58814.02
```

`--vide` force un classeur entièrement vide, échéancier compris.

- `paradise/src/data_source.py` — les données reprises des documents fournis
- `paradise/src/theme.py` — charte graphique (couleurs du logo, briques de mise en page)
- `paradise/src/build.py` — les 38 feuilles du système
- `paradise/assets/` — le logo, extrait du dossier de présentation
