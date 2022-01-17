import csv
import unidecode
import openpyxl
import numpy as np
from pip._vendor.pyparsing import col
from tabulate import tabulate
import pandas as pd
# TRAITE LE FICHIER EXCEL EXTRAIT DE WEB OF SCIENCE :
# ON OUVRE LE FICHIER, ON Y AJOUTE  DES COLONNES DEMANDEES PAR L'HCERES 
# ET ON LE SAUVE SOUS UN AUTRE NOM

# PREREQUIS :  Sortir  un fichier excel de Web OF Sciences 
# Acceder a WOS depuis l'ENT / Ressource selectroniques / Revues de A-Z   / WOS / Nouvelle Version
# Aller dans Advanced Search
# Requete 
# AD=((Inst Sci Mol Marseille OR ISM2 OR 7313 OR UMR7313 OR 6263 OR UMR6263  OR "ism 2" ) same (marseille or aix))
# AND PY=2016-2021
# ON La croise avec une requete  de  HAL-AMU (acceder par ent) / CONSULTER Par Laboratoire d'aMU / institut des sciences moleculaires de marseille
# / Selectioner les annees / puis  Outils / Export avance / Faire glisser les champs autre que Identifiant DOI vers la gauche
# Preparer l'export / Telecharger l'export
#  Traiter la liste des DOI sous WORD : remplacer les marques de fin de paragraphe (^p) par OR et entourer par : DO=( ... ) 
# faire une requete avec ces DOI dans WOS, revenir dans advanced search combiner les deux requetes avec OR
# cliquez sur le nombre de resultats de la requete resultante 
# de meme, possiblite de fusionner avec les ORCID des checheurs ( sous WOS, requete AI= ( XXX OR YYYY OR ...) AND PY=2016-2021 ) 
# cliquer sur EXPORT / Excel /  Record From= Nombre exact de resultats Record Content = Full Record 
# Ouvrir le fichier excel et enregister au format XLSX : articles_source.xlsx

# fichier / colonne d'origine
fichier_wos_source="articles_original.xlsx"
fichier_personnel_with_neworcids="personnel orcids.xlsx"

colname_publication_type="Publication Type"
colname_authors="Authors"
colname_authors_fullname="Author Full Names"
colname_article_title="Article Title"
colname_source_title="Source Title"
colname_book_series_title="Book Series Title"
colname_book_series_subtitle="Book Series Subtitle"
colname_language="Language"
colname_document_type="Document Type"
colname_conference_title="Conference Title"
colname_conference_date="Conference Date"
colname_conference_location="Conference Location"
colname_addresses="Addresses"
colname_reprint_addresses="Reprint Addresses"
colname_isbn="ISBN"
colname_publication_annee="Publication Year"
colname_orcids="ORCIDs"
colname_volume="volume"
colname_issue="issue"
colname_start_page="Start Page"
colname_end_page="End Page"
colname_doi="DOI"
colname_ut="UT (Unique WOS ID)"
colname_openaccess= "Open Access Designations"


# fichier / colonnes cible
fichier_cible="articles_hceres.xlsx"
colname_formatted_authors_fullname="Author Full Names pour HCERES"
colname_volume_issue="volume / issue"
colname_pages="Pages"
colname_interequipes="Interequipes"
colname_doctorant_coauteur="doctorant coauteur"
colname_corresponding_authors="premier, dernier ou correspoding auteur"
colname_formatted_openaccess = "Open Access"

liste_type_personnel=["permanent", "doctorant"]
liste_equipes=["BiosCiences", "Chirosciences",  "CTOM", "STeRéO"]

colnames_newcols=[colname_formatted_authors_fullname,colname_volume_issue,colname_pages,colname_interequipes,colname_doctorant_coauteur,colname_corresponding_authors,colname_formatted_openaccess]
colnames_articles=[colname_publication_type, colname_authors_fullname,  colname_article_title, colname_source_title, colname_language, colname_document_type, colname_addresses,colname_reprint_addresses,colname_doi,colname_ut] 
colnames_ouv=colnames_articles + [colname_book_series_title, colname_book_series_subtitle, colname_isbn] 
colnames_conferences=colnames_articles + [colname_conference_title, colname_conference_date, colname_conference_location]
colnames_autres=colnames_articles+ [colname_book_series_title, colname_book_series_subtitle, colname_isbn] + [colname_conference_title, colname_conference_date, colname_conference_location]
dict_colnames_by_sheetname={"articles":colnames_articles, "ouvrages":colnames_ouv, "conferences": colnames_conferences, "autre":colnames_autres}
dict_sheetname_by_doctype={ "Article":"articles", 
                  "Article; Early Access" : "articles",
                  "Review" : "articles",
                  "Article; Book Chapter": "ouvrages",
                  "Editorial Material; Book Chapter": "ouvrages",
                  "Review; Book Chapter" : "ouvrages",
                  "Proceedings Paper" : "ouvrages",
                  "Article; Proceedings Paper" : "conferences"}

try:
    workbook_source = openpyxl.load_workbook(fichier_wos_source)
except:
    print("impossible d'ouvrir le fichier :", fichier_wos_source)
    exit()
# feuilles du fichier excel
worksheet_source=workbook_source.active
num_rows = worksheet_source.max_row 
num_cols = worksheet_source.max_column 

workbook_cible=openpyxl.Workbook()
workbook_cible.remove_sheet(workbook_cible.active)
workbook_cible.create_sheet("articles")
workbook_cible.create_sheet("ouvrages")
workbook_cible.create_sheet("conferences")
workbook_cible.create_sheet("autre")

# recupere un numero de colonne par rapport a son nom en ligne 1
def getnumcol_source(col_title):
    col=1
    found=False
    while (col <=num_cols) and not found :
        cell= worksheet_source.cell(column=col, row=1)
        if (cell.value != None) and (cell.value.upper()==col_title.upper()) :
            found=True
        else :
            col=col+1
    if found :
        #print ("colonne {} est la numero {}".format(col_title, col))
        return col
    else :
        print ("colonne {} non trouvee, arret du script".format(col_title))
        exit() 
 
# affiche les erreurs sous forme de tableau
def displayErrors(lignes_en_erreur) : 
    if len (lignes_en_erreur) > 0 : 
        df = pd.DataFrame(data=lignes_en_erreur)
        print(tabulate(df, headers=("ligne", "erreur"), tablefmt='psql', showindex=False))
   
# lit le fichier permanent provoenant de l 'intranet, qui a ete recuperer avec le script export_personnel.php
# lance dans le site ism2 drupal dans le module devel - execute php code
# et met le contenu du fichier dnas un dictionnaire
print ("LISTE DU PERSONNEL")
with open('personnel.csv', newline='\n') as csvfile:
    reader = csv.reader(csvfile, delimiter=';', quotechar='"')
    headers_personnel=next(reader)
    personnel=[dict(zip(headers_personnel,i)) for i in reader]

# affiche le fichier personnel 
df = pd.DataFrame(data=personnel)
print(tabulate(df, headers='keys', tablefmt='psql', showindex=False))

ligne = 1
lignes_en_erreur=[]
print ("lignes en erreur dans le fichier du personnel ")
for row in personnel:
    if not row["type"] : 
        lignes_en_erreur.append((ligne, " type_personnel non renseigne"))
    elif not row["type"] in liste_type_personnel :
        lignes_en_erreur.append((ligne, "type_personnel {} inconnu".format(row["type"])))
    
    if not row["nom"] : 
        lignes_en_erreur.append((ligne, "le nom n'est pas renseigne"))
    
    if not row["prenom"] : 
        lignes_en_erreur.append((ligne, "le prenom n'est pas renseigne"))
           
    if not row["equipe"] in liste_equipes : 
        lignes_en_erreur.append((ligne, "equipe {} inconnue".format(row["equipe"]) ))
    ligne = ligne + 1

displayErrors(lignes_en_erreur)
    
# INITIALISATION DES NUMERO DE COLONNES 
# verifie que les colonens existent et initialise leurs numero 
print ("VERIFICATION DES COLONNES EXISTANTES : ")
col_publication_type=getnumcol_source(colname_publication_type)
col_authors=getnumcol_source(colname_authors)
col_authors_fullname=getnumcol_source(colname_authors_fullname)
col_article_title=getnumcol_source(colname_article_title)
col_source_title=getnumcol_source(colname_source_title)
col_book_series_title=getnumcol_source(colname_book_series_title)
col_book_series_subtitle=getnumcol_source(colname_book_series_subtitle)
col_language=getnumcol_source(colname_language)
col_document_type=getnumcol_source(colname_document_type)
col_conference_title=getnumcol_source(colname_conference_title)
col_conference_date=getnumcol_source(colname_conference_date)
col_conference_location=getnumcol_source(colname_conference_location)
col_addresses=getnumcol_source(colname_addresses)
col_reprint_addresses=getnumcol_source(colname_reprint_addresses)
col_isbn=getnumcol_source(colname_isbn)
col_orcids=getnumcol_source(colname_orcids)
col_publication_annee=getnumcol_source(colname_publication_annee)
col_volume=getnumcol_source(colname_volume)
col_issue=getnumcol_source(colname_issue)
col_start_page=getnumcol_source(colname_start_page)
col_end_page=getnumcol_source(colname_end_page)
col_doi=getnumcol_source(colname_doi)
col_ut=getnumcol_source(colname_ut)
col_openaccess=getnumcol_source(colname_openaccess)

# ajoutecles titre des nouvelles colonnes dans chaque feuille
for sheetname in dict_colnames_by_sheetname.keys() : 
    colnames=dict_colnames_by_sheetname[sheetname] + colnames_newcols
    for i in range(len(colnames)) : 
        workbook_cible[sheetname].cell(column=i+1, row=1).value = colnames[i] 

# parcours les lignes fichier excel : une ligne = une publication 
print ("TRAITEMENT DU FICHIER EXCEL WOS", fichier_wos_source)
print( 'nombre de lignes =', num_rows, 'nombre de colonnes =', num_cols )
lignes_en_erreur=[]
dict_labo_orcids_decouverts={}
for num_row in range(2, num_rows + 1):
    #print("traitement de la ligne", num_row)
    # sur chacune des lignes on recupere une liste d'auteurs nom prenom dans la cellule et on les parcourt
    # on met en forme la liste des auteurs  nom1, prenom1; nom2 prenom2,... en NOM1 prenom1, NOM2, prenom2...
    # on repere les collaborations interequuipe et les doctorants : pour cela verifie si un des auteur est dans la liste des personnel - equipe
    # s'il est permanent on met  sont equipe dans une  liste d'equipes dedoublonee, s'il est doctorant on indique qu'un doctorant est participant 
    list_authors_shortnames=worksheet_source.cell(column=col_authors, row=num_row).value.split("; ")
    list_authors_fullnames=worksheet_source.cell(column=col_authors_fullname, row=num_row).value.split("; ")
    list_corresponding_authors=worksheet_source.cell(column=col_reprint_addresses, row=num_row).value.split(" (corresponding author)")[0].split("; ")
    # decoupe la cellule orcid qui est de la forme, pour en extraire les champs : nom, prenom/orcid ...
    # et les mettre dans un dictionnarie indexé ainsi : (NOM, PRENOM) => orcid
    dict_authors_orcids={}
    if worksheet_source.cell(column=col_orcids, row=num_row).value :
        for token in worksheet_source.cell(column=col_orcids, row=num_row).value.split("; "):
            try :
                nomprenom, orcid = token.split("/") 
                nom,prenom= nomprenom.upper().split(", ")
                dict_authors_orcids[(nom,prenom)]=orcid
            except ValueError :
                lignes_en_erreur.append((num_row, "Champ orcids malformé, impossible d'extraire  les informations du le token {},".format(token)))
                             
    interequipe=[];
    comma=""
    formatted_authors_fullname=""
    doctorant_coauteur="N"
    corresponding_author="N"
    for indice_author_fullname in range(len(list_authors_fullnames)) : 
        author_fullname=list_authors_fullnames[indice_author_fullname]
        author_lastname="" 
        author_firstname=""
        if ", " in author_fullname :
            author_lastname, author_firstname= author_fullname.split(", ")
        else :
            author_lastname = author_fullname
        formatted_authors_fullname= formatted_authors_fullname + comma + author_lastname.upper() + " " + author_firstname
        author_shortname = list_authors_shortnames[indice_author_fullname]
        # pour chacun des auteurs de la publi on parcourt le fichier du personnel : si c'est un personnel de lism2, on recupere son equipe dans une liste dedoublonner et 
        # on verifie si l'auteur est doctorant coauteur / corresponding auteur ou dernier ou premier auteur
        # on extrait son orcid de la colonnes orcids pour le conserver sur un intranet perso et construire une requete plus affinée 
        comma=", "
        for row in personnel:
            personnel_nom_upper=unidecode.unidecode(row["nom"].upper());
            personnel_prenom_upper=unidecode.unidecode(row["prenom"].upper())
            if (personnel_nom_upper == author_lastname.upper()) and (personnel_prenom_upper == author_firstname.upper()) :
                # assertion, l'auteur fait partie de l'ism2
                if row["type"]=="permanent" :
                    if not row["equipe"] in interequipe :
                        interequipe.append(row["equipe"] )
                elif row["type"]=="doctorant"  :
                    doctorant_coauteur="O"
                if (indice_author_fullname==0) or (indice_author_fullname==(len(list_authors_fullnames) -1)) or (author_shortname in list_corresponding_authors) :
                    corresponding_author="O"
                # on tente de  determiner le orcid grace au wos
                # si on devcouvre un orcid different de celui du fichier du personnel on le sauvegarde
                orcid=dict_authors_orcids.get((personnel_nom_upper, personnel_prenom_upper))
                if orcid and not dict_labo_orcids_decouverts.get((personnel_nom_upper, personnel_prenom_upper)) :
                    dict_labo_orcids_decouverts[(personnel_nom_upper, personnel_prenom_upper)]=orcid

    #s'il ya plus d'ine equipe dans la liste interequipe dedoublonee, on enregistre une colaboration interequipe
    formatted_interequipes=""
    comma=""
    for equipe in interequipe :
        try:
            formatted_interequipes = formatted_interequipes + comma + str(liste_equipes.index(equipe) + 1)
            comma=", "   
        except ValueError :
            lignes_en_erreur.append((num_row, "L'équipe {} existe pas".format(equipe)))

    # on prepare la colonne issue, pages, open access
    volume=worksheet_source.cell(column=col_volume, row=num_row).value
    issue=worksheet_source.cell(column=col_issue, row=num_row).value
    volume_issue="{}({})".format(volume, issue)
    start_page=worksheet_source.cell(column=col_start_page, row=num_row).value
    end_page=worksheet_source.cell(column=col_end_page, row=num_row).value
    pages=" - ".join(map(str,[start_page, end_page]))

    if ( "GREEN" in worksheet_source.cell(column=col_openaccess, row=num_row).value.upper()) :
        openaccess="O"
    else :
        openaccess="N"
        
    # on determine la feuille excel en fonction du type de doncument en vue d'y copier la publication 
    doctype= worksheet_source.cell(column=col_document_type, row=num_row).value
    sheetname=dict_sheetname_by_doctype.get(doctype)
    if not sheetname :
        sheetname="autre"
    
    #on recopie les colonnes a conserver de la publication  dans la feuille correspondant a son type de document 
    colnames=dict_colnames_by_sheetname[sheetname]
    nbcol=len(colnames);
    nb_row=workbook_cible[sheetname].max_row
    for i in range(nbcol) : 
        value_source=  worksheet_source.cell(column=getnumcol_source(colnames[i]), row=num_row).value
        workbook_cible[sheetname].cell(column=i+1, row=nb_row+1).value = value_source
    # on ajoute les colonnes demandes par l'HCERES
    # colonne auteurs formtee pour hceres
    workbook_cible[sheetname].cell(column=nbcol+1, row=nb_row+1).value=formatted_authors_fullname
    # on concatene volume et issue
    workbook_cible[sheetname].cell(column=nbcol+2, row=nb_row+1).value=volume_issue
    # on remplit la colonne pages
    workbook_cible[sheetname].cell(column=nbcol+3, row=nb_row+1).value = pages
    # colonne intereiquoê   
    workbook_cible[sheetname].cell(column=nbcol+4, row=nb_row+1).value = formatted_interequipes
    # on remplit la colonne doctorant coauteur
    workbook_cible[sheetname].cell(column=nbcol+5, row=nb_row+1).value = doctorant_coauteur 
    # on remplit la colonne doctorant coauteur
    workbook_cible[sheetname].cell(column=nbcol+6, row=nb_row+1).value = corresponding_author
    # on remplit la colonne openAccess à O, s'il y a le mot green dans la colonne open access designation de wos
    workbook_cible[sheetname].cell(column=nbcol+7, row=nb_row+1).value =openaccess
    
# affiche les erreurs
displayErrors(lignes_en_erreur)

# orcid decouverts : si on a decouvert des orcids
if len (dict_labo_orcids_decouverts) > 0 : 
    # afficchage des orcid a l'ecran on transforme le dict orcids en list pour compatibilite avec DataFrame
    list_labo_orcids_decouverts =[ (nomprenom[0] + " " + nomprenom[1] , orcid) for nomprenom, orcid in dict_labo_orcids_decouverts.items() ]
    print ("LISTE DES ORCIDs DECOUVERTS :")
    df = pd.DataFrame(data=list_labo_orcids_decouverts)
    print(tabulate(df, headers=("nom prénom", "orcid"), tablefmt='psql', showindex=False))
    # On complemete le fichier excel du personnel avec leur orcid decouverts 
    # Pour cela on cree une liste a partir du dictionnaire personnel, a laquelle on rajoute les orcid pour chacun des agents
    # et ensuite on cree un nouveau fichier excel a partir de cette liste 
    wbOrcids = openpyxl.Workbook()
    ws = wbOrcids.active
    headers_personnel_neworcids = headers_personnel + ["orcid découverts"]
    listToSave = [headers_personnel_neworcids]
    req_orcid_wos="AI=("
    separator=""
    for row in personnel:
        orcid=dict_labo_orcids_decouverts.get((unidecode.unidecode(row["nom"].upper()), unidecode.unidecode(row["prenom"].upper())))
        if orcid :
            req_orcid_wos=req_orcid_wos+ separator + orcid
            xlsLine=list(row.values())
            xlsLine.append(orcid)  
            separator=" OR "
        else :
            xlsLine=list(row.values())
        listToSave.append(xlsLine) 
    req_orcid_wos=req_orcid_wos+")"
    print ("requete sur ORCID pour WOS")
    print (req_orcid_wos)
    for i in range(len(listToSave)) :
        for j in range(len(listToSave[i])) :
            ws.cell(column=j+1, row=i+1).value = listToSave[i][j]
    ws.cell(column=1, row=ws.max_row+2).value = "requete à effectuer sur les ORCID pour WOS :"
    ws.cell(column=1, row=ws.max_row+1).value =req_orcid_wos
    wbOrcids.save(fichier_personnel_with_neworcids)    
    print ("veuillez consulter le des orcids decouverts", fichier_personnel_with_neworcids)
    
# sauve le fichier
workbook_cible.save(fichier_cible)   

print ("FIN : veuillez consulter le fichier final", fichier_cible)
