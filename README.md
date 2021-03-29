# gentabdecharges

Un script python 3 destiné à générer automatiquement les tableaux de décharge des syndicats avec leur quotité de décharge pré-remplie.



## Usage



`python3 pytab.py -q quotité.xlsx template.xlsx`



`-q` ou `–quotite` est obligatoirement suivi du chemin vers le un XLSX structuré ainsi :



|      | A         | B                |
| ---- | --------- | ---------------- |
| 1    | Syndicat  | Quotité proposée |
| 2    | Ain (01)  | 0,453            |
| 3    | Aine (02) | 0,489            |
| 4    | Total     | 26,489           |



Les en-têtes de colonne doivent être présentes, de même qu’une ligne « total » à la fin. Le programme parcourt lignes de la deuxième à l’avant-dernière.

Si on veut changer ce comportement, il faut éditer le code 

```python
for ligne in syndicat.iter_rows(min_row=2,
                                    max_row=syndicat.max_row - 1,
                                    min_col=1,
                                    max_col=2,
                                    values_only=True):
```

et modifier les valeurs *min_row* et *max_row*.



`template.xlsx` est un fichier XLSX qui sert de modèle. Ce fichier sera ouvert et modifié puis enregistré en série.



## Installation

Pour fonctionner, le programme a besoin d’une version de l’interpréteur python suffisamment récente (3.6+), et des modules :

1. `sys`, `os`, `shutil`
2. `argparse`
3. `openpyxl`

Pour installer les modules :

`pip install module` ou `pip install -r requirements.txt`

## Compilation

Sur sa plateforme, `pyinstaller -F pytab.py` permet de créer un exécutable.