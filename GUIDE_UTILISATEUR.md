# Guide utilisateur : vérification ERB Sanlam ↔ Finaptik

Ce guide explique comment comparer l'ERB préparé par Sanlam avec l'ERB généré par Finaptik, et comment lire le
fichier produit. Aucune connaissance en programmation n'est nécessaire.

---

## 1. À quoi sert le script ?

Pour un compte bancaire et une date d'arrêté, on dispose de deux états de rapprochement :

- l'**ERB Sanlam**, préparé à la main par les équipes Sanlam ;
- l'**ERB Finaptik**, généré par la plateforme.

Le script met les deux ERB côte à côte, ligne par ligne. Il signale :

- les lignes identiques des deux côtés (**OK**) ;
- les lignes présentes d'un seul côté (**KO**) ;
- les lignes qui sont probablement des **extournes** ou des **annulations**.

---

## 2. Installation (une seule fois par poste)

### 2.1 Installer Python

1. Télécharger Python 3 (version 3.10 ou plus récente) sur https://www.python.org/downloads/
2. **Sous Windows**, cocher la case **« Add Python to PATH »** au premier écran de l'installation.
3. Vérifier l'installation dans un terminal (Windows : *Invite de commandes* ou *PowerShell* ; macOS : *Terminal*) :

   ```bash
   python --version      # Windows : py --version
   ```

   La commande doit afficher `Python 3.x.x`.

### 2.2 Récupérer le script

Avec Git (recommandé, pour recevoir les mises à jour) :

```bash
git clone https://github.com/issouf7507-dev/verification-erb.git
```

Sans Git : télécharger le dépôt en ZIP depuis GitHub (bouton **Code → Download ZIP**), puis le décompresser.

### 2.3 Installer la dépendance

```bash
cd verification-erb
python -m venv .venv

# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
```

> Ensuite, à chaque nouvelle session de terminal, il suffit de réactiver l'environnement
> (`.venv\Scripts\activate` ou `source .venv/bin/activate`).

### 2.4 Mettre à jour le script

```bash
cd verification-erb
git pull
```

---

## 3. Lancer une vérification

### 3.1 Préparer les fichiers

Il faut deux fichiers Excel **arrêtés à la même date** :

| Fichier | Exemple |
|---|---|
| ERB Sanlam | `PROJET FINAPTIK 2/JUIN 2026/1-STANBIC/ERB STANBIC JUIN 26.xlsx` |
| ERB Finaptik (export de la plateforme) | `ERB_STANBIC_S_A____STANBIC_2026-06-30.xlsx` |

Conseil : rangez l'ERB Sanlam dans le dossier de sa banque (`1-STANBIC`, `2-SCB SANLAM`…). Le script reprend ce nom
de dossier pour ranger le résultat.

### 3.2 Lancer la commande

1. Ouvrir un terminal et activer l'environnement (voir 2.3).
2. Se placer dans le dossier de travail, celui qui contient les données :

   ```bash
   cd "chemin/vers/resou"
   ```

3. Lancer le script avec, dans l'ordre, l'**ERB Sanlam** puis l'**ERB Finaptik** :

   ```bash
   python chemin/vers/verification-erb/verifier_erb.py "PROJET FINAPTIK 2/JUIN 2026/1-STANBIC/ERB STANBIC JUIN 26.xlsx" "ERB_STANBIC_S_A____STANBIC_2026-06-30.xlsx"
   ```

   > Mettez les chemins entre guillemets : les noms de fichiers contiennent des espaces.
   > Astuce : glissez-déposez un fichier dans le terminal pour coller son chemin.

### 3.3 Ce qui s'affiche

```
à passer par Sanlam      OK  150 | Sanlam seul   4 | Finaptik seul   2 | extourne possible   3 | annulation même ERB   4 lignes
à passer par la banque   OK   80 | Sanlam seul   1 | Finaptik seul   5 | extourne possible   3 | annulation même ERB   4 lignes
→ .../verification/JUIN 2026/1-STANBIC/STANBIC VERIFICATION ERB JUIN 2026.xlsx
```

La dernière ligne indique l'emplacement du fichier produit.

### 3.4 Où est le résultat ?

Le résultat est créé dans le dossier où la commande a été lancée, avec le même rangement que les suspens :

```
verification/
└── JUIN 2026/                                  ← mois lu dans « RAPPROCHEMENT AU 30/06/2026 »
    └── 1-STANBIC/                              ← dossier de l'ERB Sanlam
        └── STANBIC VERIFICATION ERB JUIN 2026.xlsx
```

Pour enregistrer ailleurs, ajoutez un 3e chemin à la fin de la commande :
`... "ERB FINAPTIK.xlsx" "mon dossier/mon fichier.xlsx"`.

Relancer la commande **remplace** le fichier existant. Pensez donc à recopier vos commentaires de la colonne
*Explication* avant de relancer.

---

## 4. Lire le fichier produit

### 4.1 Les onglets

| Onglet | Contenu |
|---|---|
| **ERB** | L'ERB Finaptik, tel quel |
| **a passez par sanlam** | La partie « à passer par Sanlam » des deux ERB : Débit, puis Crédit |
| **a passer par la banque** | La partie « à passer par la banque » des deux ERB : Crédit, puis Débit |

### 4.2 Les colonnes

| Colonnes | Contenu |
|---|---|
| **A** | N° de la ligne Sanlam |
| **B → I** (bandeau **bleu**) | ERB **Sanlam** : Banque, Compte, Date, N° Pièce, N° Chèque, Libellé, Débit, Crédit |
| **J → O** (bandeau **violet**) | ERB **Finaptik** : Date, N° Pièce, N° Chèque, Libellé, Débit, Crédit |
| **P – statut** | `OK` (vert) ou `KO` (rouge) |
| **Q – Extourne possible** | Voir 4.4 |
| **R – Annulation même ERB** | Voir 4.5 |
| **S – Explication** | Libre, pour vos commentaires |

Les lignes sont triées par **montant croissant**. Deux lignes sont mises sur la même ligne Excel quand elles ont
le même montant dans la même colonne. À montant égal, le script choisit d'abord la ligne avec le même N° de pièce,
puis celle qui a le même libellé, puis la date la plus proche.

### 4.3 Statut OK / KO

| Cas | Statut | Ce que ça veut dire |
|---|---|---|
| Ligne remplie des deux côtés | **OK** | La même ligne est dans les deux ERB |
| Seulement la partie bleue (Sanlam) | **KO** | Ligne présente chez Sanlam, absente de l'ERB Finaptik |
| Seulement la partie violette (Finaptik) | **KO** | Ligne présente chez Finaptik, absente de l'ERB Sanlam |

> Un **OK** veut dire « même montant », pas forcément « même opération ». Quand plusieurs lignes ont le même
> montant (par exemple des frais à 11 000), vérifiez que les N° de pièce correspondent.

### 4.4 Extourne possible (colonne Q)

**Règle :** une ligne **GL de l'ERB Finaptik** (« à passer par la banque ») est retrouvée **côté banque dans l'ERB
Sanlam** (« à passer par Sanlam »). Les deux lignes ont la même colonne Débit/Crédit, le même montant, et la même
pièce, le même n° de chèque ou le même libellé.

Exemple (fictif) : la pièce **5200000001** (250 000) est :

- chez Finaptik en « à passer par la banque · Crédit » ;
- chez Sanlam en « à passer par Sanlam · Crédit ».

La cellule indique où trouver l'autre ligne :

```
Oui ↔ Sanlam : à passer par Sanlam · Crédit (onglet « a passez par sanlam », ligne 120)
```

**Que faire ?** Vérifier quel classement est le bon, puis l'indiquer dans *Explication*. Si la ligne doit être
extournée, prévoir l'écriture ou la règle correspondante.

### 4.5 Annulation même ERB (colonne R)

**Règle :** deux lignes d'un **même ERB** (les deux Sanlam, ou les deux Finaptik) ont le même montant et
s'annulent entre les deux onglets :

- « à passer par Sanlam · Débit » ↔ « à passer par la banque · Crédit » ;
- « à passer par Sanlam · Crédit » ↔ « à passer par la banque · Débit ».

Le script cherche d'abord deux lignes avec le même n° de chèque, puis la date la plus proche.

Exemple (fictif) : dans l'ERB Sanlam, `CHQ BANQUE 1234567` (150 000, côté banque) et la pièce 5200000002
(chèque 1234567, côté GL). C'est la même opération des deux côtés : elle aurait dû être rapprochée.

**Que faire ?** Confirmer le rapprochement et le noter dans *Explication*. Si l'appariement repose seulement sur
la date (n° de chèque absent ou différent), vérifier que les deux lignes vont bien ensemble.

### 4.6 Méthode de travail conseillée

1. Filtrer la colonne **statut** sur `KO`.
2. Traiter d'abord les lignes qui ont un **« Oui ↔ … »** en colonne Q ou R : elles ont une explication probable.
3. Analyser ensuite les KO restants, ligne par ligne.
4. Noter la conclusion de chaque ligne dans **Explication** : par exemple « rapproché », « écriture d'annulation »,
   « règle à créer », « à extourner manuellement ».

---

## 5. Messages et problèmes fréquents

| Message / problème | Cause | Solution |
|---|---|---|
| `ATTENTION : ERB Sanlam au 31/05/2026, ERB Finaptik au 30/06/2026` | Les deux ERB ne sont pas du même mois | Prendre l'ERB Sanlam du même mois que l'ERB Finaptik |
| `En-tête « Date » introuvable dans …` | La 1re feuille du fichier n'a pas de ligne d'en-tête « Date » | Vérifier que c'est bien un ERB et que la feuille de l'ERB est la première du classeur |
| `PermissionError` | Le fichier de résultat est ouvert dans Excel | Fermer le fichier dans Excel et relancer |
| `ModuleNotFoundError: No module named 'openpyxl'` | L'environnement n'est pas activé, ou la dépendance n'est pas installée | Activer `.venv` puis `pip install -r requirements.txt` |
| `python` n'est pas reconnu (Windows) | Python n'est pas dans le PATH | Utiliser `py` à la place de `python`, ou réinstaller Python en cochant « Add Python to PATH » |
| `No such file or directory` | Chemin de fichier incorrect | Mettre les chemins entre guillemets, ou glisser-déposer le fichier dans le terminal |
| Tous les statuts sont vides à l'ouverture | Excel n'a pas recalculé les formules | Les formules se calculent à l'ouverture ; sinon, appuyer sur `F9` |

---

## 6. Confidentialité

Les ERB, grands livres et relevés sont des **données bancaires du client** :

- ne **jamais** les ajouter au dépôt Git : le `.gitignore` exclut déjà `*.xlsx`, `*.csv`, `*.pdf` et le
  dossier `verification/` ;
- les garder sur votre poste ou sur le partage du projet ;
- ne pas envoyer les fichiers de vérification en dehors de l'équipe projet.
