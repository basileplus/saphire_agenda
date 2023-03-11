# saphire_agenda

Le programme permet, à partir du document excel accessible en ligne contenant l'emploi du temps des Saphires, de générer un fichier .ics (fichier contenant un calendrier). 

Les fichiers .ics est directement importable sur Google Agenda, Microsoft Outlook ainsi que la plupart des applications de calendrier.

<img src="images/oldToNewAgenda.png" >

L'application est simple d'utilisation. Il suffit d'entrer sa filière, ses groupes (envoyés par Pierre Mella par mail) et le nombre de semaines de l'emploi du temps à reporter sur le fichier .ics.

![Utilisation du logiciel](images/Animation.gif)

## Fonctionnalités

- Prise en compte des groupes de l'élève
- Prise en compte du lieu des cours
- Bonne compréhension des créneaux notés "CM" et "TD" à la fois
- Possibilités d'écrire "23X" pour désigner {"230","231",...}
- Mémorisation des données saisies : le programme stocke les groupes/préférences entrées par l'utilisateur dans un fichier ``data.json``
- Ajout d'une description pour chaque évènement, contenant toutes les informations lues par le programme ; par exemple le nom de l'intervenant.
- Possibilité de ne pas ajouter les "évènements particuliers", évènements pour lesquels le programme n'a pas su identifier le cours ("231",...) et le type de cours ("CM",...)

## Limites
  * Le programme détecte mal les évènements particuliers, comme les "partiels". Un évènement sera tout de même ajouté au calendrier mais avec le titre "Evènement particulier", le détail sera présent dans sa description.
  * Les créneau de moins de deux heures peuvent poser problème
  * Le programme ne fonctionnera pas si l'emploi du temps change de format, par exemple si une ligne blanche est ajoutée au début de l'excel contenant l'emploi du temps.
