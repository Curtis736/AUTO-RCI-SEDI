#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Module pour le traitement des tags personnalisés d'images.
Gère les tags de format §§FOLDER_NAME§§ ou §§FOLDER_NAME_01§§
"""

import os
import re
import Log
import Settings

import ImageReader

class CustomImageTagProcessor:
    """
    Classe pour traiter les tags d'images personnalisés dans les documents Word.
    """
    
    def __init__(self, document, work_dir, container):
        """
        Initialise le processeur de tags personnalisés.
        
        Args:
            document: Document Word à traiter
            work_dir: Répertoire de travail
            container: Conteneur avec les données (SN, etc.)
        """
        self.document = document
        self.work_dir = work_dir
        self.container = container
        self.processed_tags = set()  # Pour éviter les doublons
    
    def find_custom_tags_in_text(self, text):
        """
        Trouve tous les tags personnalisés dans un texte.
        
        Args:
            text (str): Texte à analyser
            
        Returns:
            list: Liste des tags personnalisés trouvés
        """
        custom_tags = []
        i = 0
        
        while i < len(text):
            if text[i:i+2] == "§§":
                # Début d'un tag potentiel
                start = i
                i += 2
                tag_content = ""
                
                while i < len(text) and text[i:i+2] != "§§":
                    tag_content += text[i]
                    i += 1
                
                if i < len(text) and text[i:i+2] == "§§":
                    # Tag complet trouvé
                    custom_tags.append(tag_content)
                    i += 2
                else:
                    # Tag incomplet, continuer la recherche
                    i = start + 1
            else:
                i += 1
        
        return custom_tags
    
    def parse_custom_tag(self, tag_content):
        """
        Analyse un tag personnalisé pour extraire le nom du dossier et le filtre.
        
        Args:
            tag_content (str): Contenu du tag (ex: "FACE_E_01")
            
        Returns:
            tuple: (folder_name, filter_pattern) ou (tag_content, None)
        """
        # Vérifier si le tag contient un filtre (ex: FACE_E_01)
        if "_" in tag_content:
            parts = tag_content.split("_")
            if len(parts) >= 3 and parts[-1].isdigit():
                filter_pattern = parts[-1]
                folder_name = "_".join(parts[:-1])
                return folder_name, filter_pattern
        
        # Pas de filtre, utiliser le tag complet comme nom de dossier
        return tag_content, None
    
    def process_paragraph(self, paragraph):
        """
        Traite un paragraphe pour trouver et traiter les tags personnalisés.
        
        Args:
            paragraph: Paragraphe Word à traiter
        """
        if not paragraph or not hasattr(paragraph, 'text'):
            return
        
        text = paragraph.text
        custom_tags = self.find_custom_tags_in_text(text)
        
        for tag_content in custom_tags:
            # Éviter les doublons
            if tag_content in self.processed_tags:
                continue
            
            # Traiter le tag personnalisé
            self.process_custom_tag(tag_content)
    
    def process_custom_tag(self, tag_content):
        """
        Traite un tag personnalisé spécifique.
        
        Args:
            tag_content (str): Contenu du tag à traiter
        """
        # Analyser le tag
        folder_name, filter_pattern = self.parse_custom_tag(tag_content)
        
        # Construire le chemin du dossier
        folder_path = os.path.join(self.work_dir, folder_name)
        
        # Vérifier si le dossier existe
        if not os.path.exists(folder_path):
            Log.Warning(f"Dossier non trouvé pour le tag §§{tag_content}§§ : {folder_path}")
            return
        
        # Appeler la fonction pour ajouter les images du dossier
        Log.Message(f"Traitement du tag personnalisé §§{tag_content}§§")
        
        try:
            # Importer la fonction AddImagesFromFolder depuis Writer
            from Writer import AddImagesFromFolder
            success = AddImagesFromFolder(folder_path, tag_content, self.container.SN, filter_pattern)
            
            if success:
                # Marquer le tag comme traité
                self.processed_tags.add(tag_content)
            else:
                Log.Warning(f"Échec du traitement du tag §§{tag_content}§§")
            
        except ImportError:
            Log.Error(f"Impossible d'importer AddImagesFromFolder pour traiter le tag §§{tag_content}§§")
        except Exception as e:
            Log.Error(f"Erreur lors du traitement du tag §§{tag_content}§§ : {str(e)}")
    
    def process_all_paragraphs(self):
        """
        Traite tous les paragraphes du document pour les tags personnalisés.
        """
        Log.Message("Recherche de tags d'images personnalisés...")
        
        if not self.document.paragraphs:
            Log.Message("Aucun paragraphe trouvé dans le document ou document non initialisé")
            return
        
        for paragraph in self.document.paragraphs:
            self.process_paragraph(paragraph)
        
        if self.processed_tags:
            Log.Message(f"Tags personnalisés traités : {', '.join(self.processed_tags)}")
        else:
            Log.Message("Aucun tag personnalisé trouvé dans le document")


def process_custom_image_tags(document, work_dir, container):
    """
    Fonction principale pour traiter tous les tags d'images personnalisés.
    
    Args:
        document: Document Word à traiter
        work_dir: Répertoire de travail
        container: Conteneur avec les données
    """
    processor = CustomImageTagProcessor(document, work_dir, container)
    processor.process_all_paragraphs()


# Exemple d'utilisation :
if __name__ == "__main__":
    # Test de la fonction de recherche de tags
    test_text = "Voici un tag §§FACE_E_01§§ et un autre §§CUSTOM_FOLDER§§"
    processor = CustomImageTagProcessor(None, "", None)
    tags = processor.find_custom_tags_in_text(test_text)
    print(f"Tags trouvés : {tags}")
    
    # Test de l'analyse des tags
    for tag in tags:
        folder, filter_pattern = processor.parse_custom_tag(tag)
        print(f"Tag '{tag}' -> Dossier: '{folder}', Filtre: '{filter_pattern}'")