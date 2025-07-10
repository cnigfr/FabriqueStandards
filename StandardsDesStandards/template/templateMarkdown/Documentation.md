# Documentation de la conversion Markdown vers Word / PDF


## Introduction 

Cette documentation décrit le processus complet permettant de convertir un document rédigé au format Markdown en un document finalisé au format Word ou PDF. Elle explique également comment intégrer une page de garde au document final et assurer une mise en page normalisée.

L'objectif est de formaliser ce processus afin de le rendre facilement reproductible. 

Le processus consiste d'abord à préparer l'environnement avec les outils requis puis de structurer les fichiers sources avant de convertir le Markdown en Word en appliquant un modèle de mise en forme. La suite consiste à exporter le document Word en PDF, à compiler la page de garde LaTeX et enfin de fusionner le tout pour obtenir le document final. La suite de ce document détaille pas à pas ces étapes, les prérequis à installer et les commandes à utiliser.


## Sommaire 

1) [Installations préalables](#installations-préalables)

2) [Ressources dans standard](#ressources-dans-standard) 

3) [Conversion d'un document Markdown en PDF](#conversion-dun-document-markdown-en-pdf)   
3.a.  [Méthode 1 : Conversion en passant par Word](#méthode-1--conversion-en-passant-par-word)   
3.b.  [Méthode 2 : Conversion directe vers PDF](#méthode-2--conversion-directe-vers-pdf)   
3.c.  [Comparaison des deux méthodes](#comparaison-des-deux-méthodes) 

4) [Utilisation](#utilisation)  
4.a.  [Comment utiliser le modèle : "Document.md" ?](#comment-utiliser-le-modèle--documentmd-)   
4.b.  [Comment utiliser le modèle : "page_de_garde.tex" ?](#comment-utiliser-le-modèle--page_de_gardetex-)   
4.c.  [Les bases du langage Markdown](#les-bases-du-langage-markdown) 

5) [Adaptation](#adaptation)  
5.a.  [Comment numéroter automatiquement les figures et les tableaux ?](#comment-numéroter-automatiquement-les-figures-et-les-tableaux-)   
5.b.  [Comment générer automatiquement une liste de figures ou de tableaux ?](#comment-générer-automatiquement-une-liste-de-figures-ou-de-tableaux-)   
5.c.  [Comment générer une table des matières ?](#comment-générer-une-table-des-matières-)   
5.d.  [Comment adapter sa mise en page ?](#comment-adapter-sa-mise-en-page-)   
5.e.  [Comment adapter la page de garde de son document ?](#comment-adapter-la-page-de-garde-de-son-document-) 



## Installations préalables 

! Attention aux versions des outils que vous utilisez. Certaines posent des problèmes de compatibilité. En cas de doute, utilisez les versions indiquées dans la documentation. !

-> Installez l'outil Pandoc (https://github.com/jgm/pandoc/releases/tag/3.1.11.1).

Pandoc est le convertisseur universel de formats de documents.

-> Installez l'outil pandoc-crossref (https://github.com/lierdakil/pandoc-crossref/releases/tag/v0.3.15.0)

Pandoc-crossref est un filtre pour Pandoc qui ajoute aux documents des fonctions de numérotation automatique et de références croisées pour les tables et les figures. Il est préférable de le télécharger dans le même dossier que celui de votre pandoc. 

-> Installez l'outil Visual Studio Code (https://code.visualstudio.com/).

Utile pour avoir un environnement de travail en local et visualiser les changements effectués sur les documents.

-> Installez Python (ex version : 3.13.3) (https://www.python.org/downloads/). 

Utile pour préparer ton document Markdown avant la conversion, en automatisant des tâches comme par exemple la numérotation des figures.

-> Installez MiKTeX (https://miktex.org/download)  >  Cochez "Yes" pour "Install missing packages on-the-fly" (voir image ci-dessous). 

<img src=".\ressources_documentation\MikTex_Installation.PNG" alt="texte alternatif" width="280" />


Utile pour gérer la compilation LaTeX de la page de garde.


-> Installez MiKTeX (https://miktex.org/download)  >  Cochez "Yes" pour "Install missing packages on-the-fly" (voir image ci-dessous). 




## Ressources dans standard 

`Document.md`

-> c’est le document principal rédigé en Markdown qu'on convertira avec Pandoc.

-> il contient le texte structuré, les titres, les images, etc.

`page_de_garde.tex`

 -> c’est le fichier LaTeX de la page de garde

-> il définit la présentation de la couverture (titres, logos, date, etc.)

-> il est compilé avec pdflatex pour produire la page de garde en PDF.

`Dossier modele :` 
- Modele-styles.docx

-> c’est un modèle Word qui contient la mise en forme standardisée (styles de titres, pieds de page, en-têtes, etc.)

-> Pandoc l’utilise comme référence pour appliquer la bonne mise en page quand il convertit ton Markdown en Word.

- fig.py

-> c’est un script Python

-> il remplace dans le Markdown les balises [FIG] par une numérotation automatique des figures (par exemple “Figure: 1”, “Figure: 2”…).

-> cela permet d'éviter de numéroter les figures à la main.

- tab.py

-> c’est un script Python

-> il remplace dans le Markdown les balises [TAB] par une numérotation automatique des figures (par exemple “Table: 1”, “Table: 2”…).

-> cela permet d'éviter de numéroter les tableaux à la main.

`Dossier ressources :`

- le dossier contient : les logos, illustrations, schémas ...
    

`Dossier documentation :`

- Documentation.md

-> c’est la documentation d’explication

-> elle sert de guide technique, pour expliquer comment utiliser tous ces fichiers et reproduire le processus de conversion.

- Dossier ressources_documentation



## Conversion d'un document Markdown en PDF 

### Méthode 1 : Conversion en passant par Word 

Nous allons maintenant aborder les différentes étapes pour convertir un document Markdown en un document PDf. Ces étapes doivent être réalisé dans un terminal de commande (de préférence GitBash).

**Etape 1 (optionnelle si on utilise à l'étape 2 : --filter pandoc-crossref): Génération automatique des numéros des figures et des tables**
<span id="etape1"></span>

````
python ./modele/fig.py Document.md
````
````
python ./modele/tab.py Document.md
````
- "python ./modele/fig.py Document.md" permet de lancer un script python. Son but est de générer dans le Markdown à l'endroit et à la place de la balise [FIG] une numérotation automatique de la figure du tye "Figure: X". De même, il remplace la balise [TAB] par "Table: X".


**Etape 2 : Conversion du Markdown en un document .docx**

````
pandoc -s -f markdown -t docx --toc --toc-depth=3 --filter pandoc-crossref -o Document.docx --reference-doc=./modele/Modele-styles.docx Document.md
````
- "pandoc" est un outil très puissant pour convertir des documents d’un format à un autre.

- "s" (standalone) indique de produire un document complet (pas un fragment). Par exemple, en DOCX, cela génère un fichier Word complet.

- "f markdown" est le format d’entrée : Markdown Standard, donc le fichier source Document.md est interprété avec la syntaxe Markdown.

- "t docx" est le format de sortie : DOCX (fichier Word).

- "--toc" ajoute une table des matières (table of contents) au document.

- "--toc-depth=3" indique de prendre en compte les titres de niveaux 1, 2 et 3 pour la table des matières.

- "--filter pandoc-crossref" permet la numérotation automatique et les références croisées de tables et figures. 

- "-o Document.docx" indique le fichier de sortie : ici Document.docx.

- "--reference-doc=./modele/Modele-styles.docx" indique un modèle Word (Modele-styles.docx) que Pandoc va utiliser pour reprendre les styles, polices, etc. Cela permet d’obtenir un rendu conforme à une charte graphique.

- "Document.md" correspond au fichier source Markdown à convertir.


**Etape 3 : Exporter le .docx en PDF**

Fichier Document.docx-> Exporter -> Créer PDF -> Options -> Cocher créer des signets à l'aide de "Titres" -> Publier

<img src=".\ressources_documentation\word_pdf.PNG" alt="texte alternatif" width="500" />



- Cette étape permet d'exporter le document .docx en document .pdf tout en permettant la conservation des tables de matières.


**Etape 4 : Conversion de la page de garde Latex en PDF**
<span id="etape4"></span>
````
pdflatex -interaction=nonstopmode -halt-on-error page_de_garde.tex
````
- "pdflatex" est le compilateur LaTeX qui transforme le fichier source page_de_garde.tex en un fichier PDF.

- "-interaction=nonstopmode" est une option qui indique à "pdflatex" de continuer la compilation même s'il rencontre des erreurs.

- "-halt-on-error" est une option qui indique d'arrêter la compilation si une erreur fatale survient afin d'éviter de produire un PDF corrompu.


**Etape 5 : Suppression des fichiers auxiliaires**
<span id="etape5"></span>
````
rm -f page_de_garde.{aux,log,out}
````
- la commande "rm" supprime les fichiers auxiliaires générés par LaTeX comme : page_de_garde.aux , page_de_garde.log , page_de_garde.out. Ces fichiers contiennnet des informations de compilation qui encombrent inutilement le répertoire.


**Etape 6 : Fusion de la page de garde et du document principal**
<span id="etape6"></span>
````
pdfunite page_de_garde.pdf Document.pdf document_final.pdf
````
- "pdfunite page_de_garde.pdf rapport.pdf document_final.pdf" permet de fusionner le fichier latex de la page de garde avec le document principal.


### Méthode 2 : Conversion directe vers PDF 

Nous pouvons également convertir directement le fichier Markdown en PDF. Cette méthode présente l’avantage d’être plus concise et de comporter moins d’étapes intermédiaires que la précédente. 

Cependant, la mise en page ne sera pas personnalisée puisque cette méthode ne dispose pas d'un document Word pour choisir le style. Ainsi, cette méthode est pour l’instant à utiliser avec précaution et mérite d’être développée davantage par la suite.

**Etape 1 (optionnelle si on utilise à l'étape 2 : --filter pandoc-crossref): Génération automatique des numéros des figures et des tables** [(voir ci-dessus)](#etape1).

**Etape 2 : Conversion du Markdown en un document .pdf**

````
pandoc Document.md -o Document.pdf --pdf-engine=xelatex
````
- "pandoc" est un outil très puissant pour convertir des documents d’un format à un autre.
- "Document.md" correspond au fichier source Markdown à convertir.
- "-o Document.pdf" indique le fichier de sortie : ici Document.pdf.
- "--pdf-engine=xelatex" est une option qui précise quel moteur LaTeX utiliser pour produire le PDF. Par défaut, Pandoc ne crée pas directement de PDF. Il transforme d'abord le MAarkdown en Latex, puis utilise un moteur Latex pour le compiler en pdf.


**Etape 3 : Conversion de la page de garde Latex en PDF**
[(voir ci-dessus)](#etape4).

**Etape 4 : Suppression des fichiers auxiliaires**
[(voir ci-dessus)](#etape5).

**Etape 5 : Fusion de la page de garde et du document principal**
[(voir ci-dessus)](#etape6).


### Comparaison des deux méthodes 

**Méthode 1 :**

<img src=".\ressources_documentation\methode1.PNG" alt="texte alternatif" width="400" />

**Méthode 2 :**

<img src=".\ressources_documentation\methode2.PNG" alt="texte alternatif" width="400" />


## Utilisation  
### Comment utiliser le modèle : "Document.md" ? 

-> Ce modèle de standard est une ossature sur laquelle vous pourrez vous appuyer pour écrire votre standard conformément aux normes d'écritures en vigueur et aux bonnes pratiques du CNIG

-> Si une partie ou section est optionnelle, cela sera indiqué. Autrement, elle devra apparaître dans votre standard.

-> Les aides et explications sont `surlignés` et entourés des symboles <>. Ls exemples sont seulement `surlignés` (si vous choisissez de reprendre le texte de l'exemple, retirez le surlignage).

### Comment utiliser le modèle : "page_de_garde.tex" ? 

-> Si vous avez plusieurs sponsors, ils pourront être indiqués en partie 1.2,

-> Préférez un logo officiel en format PNG (souvent disponible sur le site du collaborateur via le service communication) plutôt qu'une image récupérée sur un moteur de recherche,

-> Le logo du sponsor doit être de la même hauteur que le logo du CNIG.

### Les bases du langage Markdown 

<img src=".\ressources_documentation\BasesMarkdown.PNG" alt="texte alternatif" width="400" />

-> Pour aller à la ligne (sans sauter de ligne), appuyez deux fois sur la touche espace puis appuyez sur la touche Entrée.  
Utilisez < br > pour aller à la ligne suivante (sans sauter de ligne) si vous êtes dans un tableau.

-> Pour souligner un mot : `<u>mot</u>`

-> Pour insérer une image : `![Texte alternatif](chemin/vers/image.jpg)`



## Adaptation 

### Comment numéroter automatiquement les figures et les tableaux ? 

### Comment générer automatiquement une liste de figures ou de tableaux ? 
#### Méthode 1 : Utilisations de balises et de scripts python
##### Numérotation automatique des figures 

La balise [FIG] et la ligne de commande : "python fig.py Document.md" permet de générer à l'endroit et à la place de la balise une numérotation automatique de la figure du type "Figure: X". Pour ce faire, il suffit simplement de placer la balise [FIG] une ou plusieurs fois à ou aux endroits, où vous souhaiteriez indiquer la numérotation d'une figure dans le **Document.md**. Ainsi, une fois la ligne de commande "python fig.py Document.md" lancée, elle remplacera automatiquement la balise par la bonne numérotation.

**Exemple :**

<img src=".\ressources_documentation\Exemple_Balise.PNG" alt="texte alternatif" width="500" />    

Après exécution du script :  

<img src=".\ressources_documentation\Exemple2_Balise.PNG" alt="texte alternatif" width="500" />


##### Numérotation automatique des tableaux 

La balise [TAB] et la ligne de commande : "python tab.py Document.md" permet de générer à l'endroit et à la place de la balise une numérotation automatique du tableau du type "Table: X". Pour ce faire, il suffit simplement de placer la balise [TAB] une ou plusieurs fois à ou aux endroits, où vous souhaiteriez indiquer la numérotation d'un tableau dans le **Document.md**. Ainsi, une fois la ligne de commande "python tab.py Document.md" lancée, elle remplacera automatiquement la balise par la bonne numérotation.

#### Méthode 2 : Utilisations du filtre pandoc-crossref
##### Numérotation automatique des figures et des tableaux 

Pandoc-crossref est un filtre pour Pandoc qui ajoute aux documents des fonctions de numérotation automatique et de références croisées pour les tables et les figures.

Concrètement, lorsque vous écrivez dans le document Document.md vos tableaux et figures, après conversion pandoc, cela affiche automatiquement la numérotation comme ceci :  

<img src=".\ressources_documentation\pandoc-crossref.PNG" alt="texte alternatif" width="500" />

! Attention : pour que pandoc-crossref fonctionne normalement, il faut impérativement respecter la syntaxe d'écriture Markdown de vos tableaux et figures comme indiquée dans l'image ci-dessus !

##### Génération automatique de listes de figures et de tableaux 

Pandoc-crossref permet également de générer automatiquement une liste de figures et une liste de tableaux avec leur numérotation et leur légende dans le document Markdown.

Il suffit d'insérer des balises spéciales dans le document Markdown à l'endroit où vous voulez que la liste apparaisse :  

- Pour la liste des figures :

`\listoffigures`

- Pour la liste des tableaux : 

`\listoftables`


### Comment générer une table des matières ? 

L'option de la ligne de commande Pandoc "--toc --toc-depth=3" permet de générer automatiquement au début du document un sommaire qui reprend les titres et sous-titres du même document. "depth=3" indiqe le niveau maximal de titres à inclure, dans cet exemple il est préréglé à 3. On ne peut pas positionner la table des matières là où on le souhaite, elle s'affiche directement au début du document.


### Comment adapter sa mise en page ? 

Le "Modele-styles.docx" est un fichier Word, utile pour la mise en page du Document.md lors de sa conversion en fichier Word puis en fichier pdf.

#### En-têtes 
Le modèle contient des styles prédéfinis pour les en-têtes et pieds de page. Lors de la conversion, Pandoc applique automatiquement ces styles, ce qui garantit une uniformité sur toutes les pages. Pour personnaliser les en-têtes, modifiez-les directement. Vous pouvez y insérer des numéros de pages ou toute autre information répétée.

#### Style titres et texte 
Le modèle définit des styles pour les différents niveaux de titres (Titre 1, Titre 2, Titre 3, etc.) ainsi que pour le corps du texte (Normal). Pour que la conversion applique correctement les styles, veillez à utiliser ces styles dans votre Markdown via la hiérarchie des titres (#,##,###), que Pandoc associera aux styles correspondants dans Word. Il est aussi possible de personnaliser la police, la taille, l'interligne et les couleurs en modifiant le modèle.

#### Mise en page des tableaux 
Les tableaux dans le document converti adoptent le style défini dans le modèle Word, notamment en termes de police, bordures, taille, espacements, alignements, styles des titres de colonne ou encore les couleurs. Pour ajuster la présentation des tableaux : sélectionnez un tableau (ou insertion>Tableau) dans Word -> cliquez sur l'onglet Conception de la table -> repérez la section Styles de tableau -> cliquez sur la petite flèche en bas pour ouvrir le panneau des styles -> Allez sur Modifier le style de tableau ... -> Modifiez le style du tableau à votre convenance -> Cliquez sur ok.


<img src=".\ressources_documentation\Style_Tableau.PNG" alt="texte alternatif" width="500" />


D'autres fonctionnalités comme les répétitions des titres des colonnes lors des changements de page sont disponibles en cliquant le bouton Format.


#### Mise en page des légendes 

Vous pouvez choisir le style de vos légendes en modifiant dans le Word de référence les styles : - d'Image Caption (pour les légendes des figures/images) ou - de Table Caption (pour les légendes des tables).

Pandoc adopte ce style uniquement si vous écrivez dans le Document.md :

- pour la légende d'une figure/image : il faut que la légende soit située en-dessous de la figure. Ex :

`![texte alternatif](chemin/image.png)`  
`Figure: La légende`    

ou en utilisant la balise :

`![texte alternatif](chemin/image.png)`  
`[FIG] La légende` 

- pour la légende d'une table : il faut que la légende soit située au-dessus du tableau. Ex :

`Table: La légende`

`| a | b |`  
`|---|---|`  
`| 1 | 2 |`


### Comment adapter la page de garde de son document ? 

Le fichier page_de_garde.tex est un fichier LaTeX qui inclut la configuration et le contenu de la page de garde.

Ce fichier contient les packages et paramètres LaTeX qui définissent la mise en forme générale de la page de garde :
-  `\usepackage{geometry} avec margin=2.5cm` définit les marges

Il contient également le contenu avec les différents titres, sous-titres, logos etc.
Pour personnaliser la page de garde :
- Il suffit de modifier les textes et d'adapter la taille de leur police (Huge, Large ...). 
- Remplacez les chemins des images par les vôtres. 
- Vous pouvez aussi adpater les espacements en cm. 
- Une page blanche est insérée avec `\newpage \thispagestyle{empty} \mbox{} \newpage` pour séparer la page de garde du reste du document pour l'impression.
