import spacy
import openpyxl
from openpyxl import Workbook
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.popup import Popup
import uuid
import re

# Charger le modèle de langue anglais de spaCy
nlp = spacy.load("en_core_web_sm")

class MyWidget(BoxLayout):

    def __init__(self, **kwargs):
        super(MyWidget, self).__init__(**kwargs)
        self.orientation = 'vertical'

        self.label = Label(text='Entrez votre texte ci-dessous :')
        self.add_widget(self.label)

        self.text_input = TextInput(multiline=True, size_hint=(1, 0.8))
        self.add_widget(self.text_input)

        self.button = Button(text='Créer Fichier Excel', size_hint=(1, 0.2))
        self.button.bind(on_press=self.creer_fichier_excel)
        self.add_widget(self.button)

    def analyser_texte(self, texte):
        # Patterns regex pour extraire les informations spécifiques
        name_pattern = re.compile(r'\b[A-Z][a-z]+\b')
        age_pattern = re.compile(r'\b\d{1,2}\s*ans?\b')
        job_pattern = re.compile(
            r'\b(développeur|ingénieur|médecin|enseignant|avocat|architecte|comptable|designer|électricien|plombier|maçon|menuisier|infirmier|journaliste|pharmacien|psychologue|vétérinaire|pilote|écrivain|acteur|musicien|chanteur|danseur|peintre|sculpteur|photographe|chef|serveur|barman|cuisinier|pâtissier|agriculteur|horticulteur|pêcheur|chauffeur|mécanicien|technicien|scientifique|recherchiste|bibliothécaire|libraire|vérificateur|assistant|secrétaire|gestionnaire|administrateur|directeur|manager|consultant|analyste|spécialiste|stratège|entraîneur|coach|sportif|athlète|etc)\b',
            re.IGNORECASE
        )
        hobby_pattern = re.compile(
    r'\baime\s+[a-z]+\b|\bjouer\s+à\s+[a-z]+\b|\bregarder\s+[a-z]+\b|'
    r'\bfaire\s+de\s+[a-z]+\b|\bpratiquer\s+[a-z]+\b|\bcuisiner\s+[a-z]+\b|'
    r'\bécouter\s+[a-z]+\b|\blire\s+[a-z]+\b|\bpeindre\s+[a-z]+\b|\bdessiner\s+[a-z]+\b|'
    r'\bnager\s+[a-z]+\b|\bcourir\s+[a-z]+\b|\bsport\s+[a-z]+\b|'
    r'\bjardinage\s+[a-z]+\b|\bphotographie\s+[a-z]+\b|\bvoyager\s+[a-z]+\b|'
    r'\bchanter\s+[a-z]+\b|\bdanser\s+[a-z]+\b|\bcollectionner\s+[a-z]+\b|'
    r'\bfaire\s+du\s+[a-z]+\b|\bfaire\s+de\s+la\s+[a-z]+\b|\bfaire\s+des\s+[a-z]+\b|'
    r'\btravailler\s+[a-z]+\b|\bétudier\s+[a-z]+\b|\bexplorer\s+[a-z]+\b|'
    r'\bcréer\s+[a-z]+\b|\binventer\s+[a-z]+\b|\bjouer\s+avec\s+[a-z]+\b|'
    r'\bparticiper\s+à\s+[a-z]+\b|\bvisiter\s+[a-z]+\b|\badmirer\s+[a-z]+\b',
    re.IGNORECASE
)
        names = name_pattern.findall(texte)
        ages = age_pattern.findall(texte)
        jobs = job_pattern.findall(texte)
        hobbies = hobby_pattern.findall(texte)

        # Organiser les données
        donnees = [["NAME", "AGE", "JOB", "HOBBY"]]

        max_length = max(len(names), len(ages), len(jobs), len(hobbies))

        for i in range(max_length):
            row = []
            row.append(names[i] if i < len(names) else "")
            row.append(ages[i] if i < len(ages) else "")
            row.append(jobs[i] if i < len(jobs) else "")
            row.append(hobbies[i] if i < len(hobbies) else "")
            donnees.append(row)

        return donnees

    def creer_fichier_excel(self, instance):
        texte_utilisateur = self.text_input.text
        donnees = self.analyser_texte(texte_utilisateur)

        wb = Workbook()
        ws = wb.active

        for ligne in donnees:
            ws.append(ligne)

        # Générer un nom de fichier aléatoire
        nom_fichier = 'resultat_' + str(uuid.uuid4()) + ".xlsx"
        wb.save(nom_fichier)

        popup = Popup(title='Succès',
                      content=Label(text=f"Fichier {nom_fichier} créé avec succès."),
                      size_hint=(None, None), size=(400, 200))
        popup.open()


class ExelGeneratorApp(App):
    def build(self):
        return MyWidget()


if __name__ == '__main__':
    ExelGeneratorApp().run()
