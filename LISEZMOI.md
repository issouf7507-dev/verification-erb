# Vérification ERB Sanlam ↔ Finaptik

Compare l'ERB (état de rapprochement bancaire) préparé par Sanlam avec l'ERB généré par Finaptik, et produit un
fichier Excel de vérification ligne à ligne.

👉 **Guide pas à pas pour les utilisateurs : [GUIDE_UTILISATEUR.md](GUIDE_UTILISATEUR.md)**

## Installation (une seule fois)

1. Installer Python 3.10 ou plus récent : https://www.python.org/downloads/
   (sous Windows, cocher « Add Python to PATH »).
2. Récupérer le dépôt puis installer la dépendance :

```bash
git clone https://github.com/issouf7507-dev/verification-erb.git
cd verification-erb
python -m venv .venv
# macOS / Linux
source .venv/bin/activate
# Windows
.venv\Scripts\activate
pip install -r requirements.txt
```

## Utilisation

Se placer dans le dossier de travail (celui qui contient les données, par ex. `resou`), puis :

```bash
python chemin/vers/verifier_erb.py "PROJET FINAPTIK 2/JUIN 2026/1-STANBIC/ERB STANBIC JUIN 26.xlsx" "ERB_STANBIC_S_A____STANBIC_2026-06-30.xlsx"
```

- 1er argument : ERB Sanlam
- 2e argument : ERB Finaptik
- 3e argument (facultatif) : chemin du fichier de sortie

Sans 3e argument, le résultat est rangé comme les suspens, dans le dossier où la commande est lancée :

```
verification/JUIN 2026/1-STANBIC/STANBIC VERIFICATION ERB JUIN 2026.xlsx
```

- le mois vient de « RAPPROCHEMENT AU … » de l'ERB Finaptik ;
- le dossier banque (`1-STANBIC`) est celui qui contient l'ERB Sanlam.

Le script affiche un avertissement si les deux ERB ne sont pas arrêtés à la même date.

## Le fichier produit

Trois onglets :

1. **ERB** : l'ERB Finaptik, tel quel.
2. **a passez par sanlam** : « à passer par Sanlam » (Débit puis Crédit), Sanlam (bleu) à gauche, Finaptik (violet) à droite.
3. **a passer par la banque** : « à passer par la banque » (Crédit puis Débit), même présentation.

Les lignes sont triées par montant. Deux lignes sont mises côte à côte quand elles ont le même montant dans la même
colonne ; à montant égal, on préfère le même N° de pièce, puis le même libellé, puis la date la plus proche.

| Colonne | Contenu |
|---|---|
| **statut** | `OK` (vert) si les montants Sanlam et Finaptik sont égaux, `KO` (rouge) sinon |
| **Extourne possible** | Ligne GL de l'ERB Finaptik (« à passer par la banque ») retrouvée côté banque dans l'ERB Sanlam (« à passer par Sanlam »), même colonne Débit/Crédit, même montant, même pièce / n° de chèque / libellé |
| **Annulation même ERB** | Deux lignes d'un même ERB qui s'annulent entre les deux onglets : « à passer par Sanlam · Débit » ↔ « à passer par la banque · Crédit », et Crédit ↔ Débit (même n° de chèque d'abord, puis date la plus proche) |
| **Explication** | Libre, pour les commentaires |

Les cellules « Oui ↔ … » indiquent l'onglet et la ligne de la ligne correspondante. Ce sont des **propositions** à
vérifier, surtout quand l'appariement ne repose que sur la date (plusieurs lignes de même montant).

## Format attendu des ERB

Première feuille du classeur, avec une ligne d'en-tête contenant `Date`, puis dans l'ordre :
`Date | N° Pièce | N° Chèque | Libellé | Débit | Crédit (à passer par Sanlam) | Débit | Crédit (à passer par la banque)`.
L'ERB Sanlam peut avoir en plus `Banque` et `Compte` avant la date.

## Données

Les fichiers ERB, grands livres et relevés sont des données bancaires du client : ils ne doivent **jamais** être
ajoutés au dépôt (le `.gitignore` exclut `*.xlsx`, `*.csv`, `*.pdf` et le dossier `verification/`).
