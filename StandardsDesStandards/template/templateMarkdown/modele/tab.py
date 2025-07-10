import re
import sys

def numeroter_tableaux(fichier):
    with open(fichier, 'r', encoding='utf-8') as f:
        contenu = f.read()

    # Expression régulière pour détecter les variantes : [TAB], [tab], [[TAB]], [[tab]]
    pattern = re.compile(r'\[{1,2}[Tt][Aa][Bb]\]{1,2}')

    def remplacement(match):
        remplacement.compteur += 1
        return f'Table: {remplacement.compteur}'

    remplacement.compteur = 0
    contenu_modifie = pattern.sub(remplacement, contenu)

    with open(fichier, 'w', encoding='utf-8') as f:
        f.write(contenu_modifie)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage : python tableau.py chemin/vers/le/fichier.md")
        sys.exit(1)

    chemin_fichier = sys.argv[1]
    numeroter_tableaux(chemin_fichier)

