# Templating
Description du Processus
Le projet implémente une solution de templating pour automatiser la création de documents Word relatifs à la scission de sociétés. Le processus se déroule en deux étapes principales :

Préparation du Template Word :

Un document Word anonymisé (Template_Traitement_Scission.docx) sert de modèle de base. Ce modèle contient des marqueurs spécifiques (placeholders) qui définissent les emplacements où les données personnalisées doivent être insérées.
Extraction et Insertion des Données :

Les données de chaque société sont extraites d'un fichier Excel anonymisé (Donnees_Societes_Scission.xlsx), qui contient toutes les informations nécessaires à la personnalisation des documents.
Un script ou une application lit ces données et remplace les marqueurs dans le modèle Word par les informations correspondantes pour chaque société.

Avantages de la Technique
- Efficacité : Réduit considérablement le temps nécessaire à la création de documents personnalisés.
- Uniformité : Assure la cohérence de la mise en forme et du contenu à travers tous les documents générés.
- Précision : Minimise les erreurs humaines en automatisant le remplissage des données.
