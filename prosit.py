import sys, os
from PyQt5.QtWidgets import QApplication, QFileDialog, QWidget
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PyQt5.QtCore import QDir
from docx import Document
import re
import docx 
import openai

app = QApplication(sys.argv)
dialog = QFileDialog()  # Définissez la boîte de dialogue comme une variable globale

def get_docx_file():
    dialog.setFileMode(QFileDialog.ExistingFiles)
    dialog.setFilter(QDir.Files)

    if dialog.exec_():
        file_name = dialog.selectedFiles()
        if file_name[0].endswith('.docx'):
            print(file_name[0])
            return True, str(file_name[0])
        else:
            print("Erreur : Le fichier sélectionné n'est pas un docx")
            return False, ""
    else:
        print("Erreur : Pas de fichier sélectionné")
        return False, ""

def generate_text(prompt):
    
    openai.api_key = "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"

    completions = openai.Completion.create(
        engine="text-davinci-002",
        prompt=prompt,
        max_tokens=1024,
        n=1,
        stop=None,
        temperature=0.5,
    )

    message = completions.choices[0].text
    return message.strip()

def est_titre_de_categorie(ligne, tab_categorie):
    # Vérifie si la ligne est un titre de catégorie suivi de ":"
    for catégorie in tab_categorie:
        if ligne.strip().lower() == catégorie.lower() + " :":
            return True
        if ligne.strip().lower() + " :" == catégorie.lower():
            return True
        if ligne.strip().lower() == catégorie.lower() :
            return True
    return False


def afficher_categories(texte, tab_categorie):
    # Convertir le texte en lignes
    lignes = texte.split('\n')
    
    # Initialiser une variable pour suivre le type actuel (1 pour catégorie, 2 pour mots clés, 3 pour le texte)
    type_actuel = 1  # Texte par défaut
    

	# Parcourir chaque ligne du texte
    titre = True

    for ligne in lignes:

        if titre == True and ligne != "":
            titre = doc_final.add_heading(ligne, level=1)  
            titre.alignment = WD_ALIGN_PARAGRAPH.CENTER

            titre = False

        elif est_titre_de_categorie(ligne, tab_categorie):    # Vérifier si la ligne est un titre de catégorie suivi de ":"

            type_actuel = 1  
            doc_final.add_heading(ligne, level=1)  

            if 'Mots clés' in ligne or 'Mots clés ' in ligne or 'Mots-clés' in ligne or'Mot clés' in ligne or 'Mot-clés'in ligne or 'Définition mots-clefs'in ligne or 'Définition mots-clef'in ligne or 'Mots clefs : ' in ligne:
                type_actuel = 2  # Passer au type "Mots clés"

            elif 'Problématique' in ligne or 'Problématiques' in ligne or 'problématique' in ligne or 'problématiques' in ligne or 'Problématique ? ' in ligne:
                type_actuel = 4 
            else:
            	type_actuel = 1  


        elif type_actuel == 2 and ligne.strip() : # mots cles a traiter
            #doc_final.add_paragraph("mots cles" + ligne)  
            doc_final.add_paragraph("        - " + ligne)  

        elif type_actuel == 4  and ligne.strip(): # Problématique a traiter

            paragraph = doc_final.add_paragraph()
            run = paragraph.add_run("               - " + ligne)
            bold_style = run.bold = True

        else:
            type_actuel = 3  # Texte par défaut
            doc_final.add_paragraph(ligne)

# Récupérer le nom du fichier Word
importOk, doc_name = get_docx_file()

# Si le fichier a été sélectionné avec succès, copier son contenu dans un fichier temp.txt et appliquer les traitements
if importOk:
    try:

        doc = Document(doc_name)

        chemin_fichier  = os.path.dirname(doc_name)
        print("path :" , chemin_fichier )
        contenu = ""

        # Parcourir le contenu du fichier Word et le copier dans contenu
        for paragraphe in doc.paragraphs:
            contenu += paragraphe.text + '\n'

        # Copier le contenu dans un fichier temp.txt
        with open('temp.txt', 'w', encoding='utf-8') as fichier_temp:
            fichier_temp.write(contenu)

        # Liste des catégories
        tab_categorie = ['Analyse du contexte','Définition mots-clefs ', 'Contexte', 'Mots clés', 'Mots-clés', 'Mot clés', 'Mot-clés', 'Problématique', 'Contrainte', 'Livrable','Livrables', 'Généralisation', 'Pistes de solutions', 'Plan d’action', 'Réalisation du plan d’action']

        # Lire le contenu du fichier temp.txt et appliquer les traitements
        with open('temp.txt', 'r', encoding='utf-8') as fichier_temp:
            contenu_temp = fichier_temp.read()


        doc_final = docx.Document()
        doc_final.add_picture("logo_cesi.jpg")

        # Afficher le contenu en respectant les catégories et les couleurs
        afficher_categories(contenu_temp, tab_categorie)

        match = re.search(r'\d+', doc_name)
        if match:
            var_num = match.group()
        print("Prosit numéro : ", var_num)

        print("fichier editer avec succes")

        doc_fini_name = "CER_prosit_" + var_num + "_hugo_laplace.docx"
        doc_final.save(chemin_fichier + "\\"+ doc_fini_name)
        print("fichier enregistre sous : ", chemin_fichier + "\\"+ doc_fini_name)

        os.remove('temp.txt')
        #os.system(doc_fini_name)


    except Exception as e:
        print("Erreur lors de l'extraction du texte :", str(e))
