#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Script pour générer un PDF de la documentation AUTO RCI
Utilise les mêmes dépendances que l'application principale
"""

import os
import sys
from datetime import datetime

# Essayer d'importer reportlab
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    REPORTLAB_AVAILABLE = True
except ImportError:
    print("⚠️  ReportLab n'est pas installé. Tentative d'installation...")
    try:
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "reportlab"])
        from reportlab.lib.pagesizes import A4
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import inch
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        REPORTLAB_AVAILABLE = True
        print("✅ ReportLab installé avec succès")
    except Exception as e:
        print(f"❌ Impossible d'installer ReportLab: {e}")
        print("📝 Création d'un fichier texte à la place...")
        REPORTLAB_AVAILABLE = False

def create_documentation_text():
    """Génère un fichier texte de la documentation si PDF impossible"""
    
    txt_filename = f"Documentation_AUTO_RCI_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    
    content = """
GUIDE D'UTILISATION - AUTO RCI
==============================

Générateur Automatique de Documents
Version du {}

Documentation complète pour tous les utilisateurs

TABLE DES MATIÈRES
==================
1. À quoi sert ce logiciel ?
2. Prérequis et installation
3. Première utilisation
4. Configuration détaillée
5. Génération de documents
6. Structure des fichiers
7. Dépannage et problèmes courants
8. Fonctionnalités avancées
9. Support et contact

1. À QUOI SERT CE LOGICIEL ?
============================

AUTO RCI est un logiciel spécialement conçu pour automatiser la création de 
documents de contrôle qualité dans l'industrie. Il permet de générer 
automatiquement des rapports Word à partir de données Excel et de modèles Word.

Principales fonctionnalités :
• Génération automatique de documents RCI (Rapport de Contrôle Interne)
• Génération automatique de documents RE (Rapport d'Essai)
• Génération automatique de documents PV (Procès-Verbal)
• Remplacement automatique de tags dans les modèles Word
• Lecture automatique des données Excel de mesures
• Recherche automatique des fichiers et dossiers
• Validation automatique des mesures
• Gestion des images et graphiques

2. PRÉREQUIS ET INSTALLATION
============================

Logiciels requis :
• Windows 10 ou plus récent
• Microsoft Excel (pour lire les fichiers de mesures)
• Microsoft Word (pour générer les documents)
• Python 3.7 ou plus récent

Installation :
1. Assurez-vous que tous les logiciels requis sont installés
2. Ouvrez le dossier du logiciel AUTO RCI
3. Double-cliquez sur le fichier 'lancer_RCI.bat'
4. L'application se lance automatiquement

3. PREMIÈRE UTILISATION
=======================

Configuration rapide avec 'Find Data' :
L'onglet 'Find Data' permet de configurer automatiquement tous les chemins nécessaires :

1. Entrez le numéro de lancement (ex: LT2500414)
2. Entrez le formulaire (ex: F458 E)
3. Sélectionnez la catégorie (RE, PV, ou RCI)
4. Cliquez sur 'fill'
5. Vérifiez que tous les chemins ont été trouvés

4. CONFIGURATION DÉTAILLÉE
==========================

Onglet 'Paths' :
• Dossier de lancement : Le dossier principal de travail
• Fichier Excel : Le fichier contenant vos mesures
• Modèle Word : Le template pour générer les documents

Onglet 'Générateur fiches RCI/RE' :
• Format du nom de fichier : Comment nommer les documents générés
• Chemin de sortie : Où sauvegarder les documents
• Réaction aux fichiers préexistants : Que faire si un fichier existe déjà

5. GÉNÉRATION DE DOCUMENTS
==========================

Génération d'un document unique :
1. Allez dans l'onglet 'Générateur fiches RCI/RE'
2. Cliquez sur 'générer document'
3. Sélectionnez le SN pour lequel générer le document
4. Le document sera créé dans le dossier de sortie configuré

Génération en lot :
1. Dans l'onglet 'Générateur fiches RCI/RE'
2. Sélectionnez plusieurs SN
3. Le logiciel générera un document pour chaque SN sélectionné

6. STRUCTURE DES FICHIERS
=========================

Structure de dossiers attendue :
Le logiciel s'attend à une structure de dossiers standardisée :

X:/Tracabilite/
└── RCT°-AGS23.159B/
    ├── LT2500414/          ← Dossier de travail
    └── Mesures 23.159.xlsx ← Fichier de mesures

Format du fichier Excel :
Le fichier Excel doit contenir :
• Une première ligne avec les en-têtes contenant des tags entre **
• Une colonne SN obligatoire
• Des données pour chaque numéro de série

Exemple de structure Excel :
| SN | **RL1** | **RL2** | **IL_BEFORE** | **IL_AFTER** |
|----|---------|---------|---------------|--------------|
| 001 | -45.2 | -42.1 | -0.8 | -0.9 |
| 002 | -48.3 | -44.7 | -0.7 | -0.8 |

7. DÉPANNAGE ET PROBLÈMES COURANTS
==================================

Problèmes fréquents et solutions :

Le logiciel ne démarre pas : Vérifiez que Python, Word et Excel sont installés

Erreur 'le dossier n'existe pas' : Vérifiez que le dossier existe et que vous avez les droits d'accès

Erreur 'tag non remplacé' : Vérifiez que votre fichier Excel contient la colonne SN

'Find Data' ne trouve pas les fichiers : Vérifiez la structure de dossiers et les permissions

Excel ne s'ouvre pas : Fermez Excel et relancez l'application

Word ne s'ouvre pas : Vérifiez que Word est installé et que le template n'est pas corrompu

8. FONCTIONNALITÉS AVANCÉES
============================

Validation automatique des mesures :
Le logiciel vérifie automatiquement :
• Les plages de valeurs typiques pour les mesures en dB
• Les spécifications définies dans le fichier Excel
• La cohérence des données

Gestion des images :
Le logiciel peut automatiquement :
• Chercher des images dans des dossiers spécifiques
• Les associer aux SN correspondants
• Les insérer dans les documents

9. SUPPORT ET CONTACT
=====================

Logs et diagnostic :
Les logs sont disponibles dans le dossier 'Logs/' :
• normal.log : Messages généraux
• problems.log : Erreurs et avertissements
• verbose.log : Messages détaillés

En cas de problème :
1. Consultez les logs dans le dossier 'Logs/'
2. Vérifiez que tous les chemins sont corrects
3. Assurez-vous que les fichiers Excel et Word ne sont pas ouverts
4. Contactez le support technique si le problème persiste

FIN DE LA DOCUMENTATION
=======================

Cette documentation a été générée automatiquement par AUTO RCI.
Pour toute question ou suggestion d'amélioration, n'hésitez pas à contacter l'équipe de développement.

""".format(datetime.now().strftime('%d/%m/%Y'))
    
    try:
        with open(txt_filename, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"✅ Documentation texte générée avec succès : {txt_filename}")
        print(f"📁 Fichier créé dans : {os.path.abspath(txt_filename)}")
        return txt_filename
    except Exception as e:
        print(f"❌ Erreur lors de la génération du fichier texte : {str(e)}")
        return None

def create_documentation_pdf():
    """Génère un PDF de la documentation complète"""
    
    if not REPORTLAB_AVAILABLE:
        return create_documentation_text()
    
    # Nom du fichier PDF
    pdf_filename = f"Documentation_AUTO_RCI_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    
    # Créer le document PDF
    doc = SimpleDocTemplate(pdf_filename, pagesize=A4)
    
    # Styles
    styles = getSampleStyleSheet()
    
    # Styles personnalisés
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        spaceAfter=30,
        alignment=TA_CENTER,
        textColor=colors.darkblue
    )
    
    subtitle_style = ParagraphStyle(
        'CustomSubtitle',
        parent=styles['Heading2'],
        fontSize=14,
        spaceAfter=20,
        spaceBefore=20,
        textColor=colors.darkred
    )
    
    section_style = ParagraphStyle(
        'CustomSection',
        parent=styles['Heading3'],
        fontSize=12,
        spaceAfter=15,
        spaceBefore=15,
        textColor=colors.darkgreen
    )
    
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontSize=10,
        spaceAfter=6,
        alignment=TA_JUSTIFY
    )
    
    code_style = ParagraphStyle(
        'CustomCode',
        parent=styles['Code'],
        fontSize=9,
        spaceAfter=6,
        fontName='Courier',
        backColor=colors.lightgrey
    )
    
    # Contenu du PDF
    story = []
    
    # Page de titre
    story.append(Paragraph("GUIDE D'UTILISATION", title_style))
    story.append(Spacer(1, 20))
    story.append(Paragraph("AUTO RCI - Générateur Automatique de Documents", subtitle_style))
    story.append(Spacer(1, 30))
    story.append(Paragraph(f"Version du {datetime.now().strftime('%d/%m/%Y')}", normal_style))
    story.append(Paragraph("Documentation complète pour tous les utilisateurs", normal_style))
    story.append(PageBreak())

    # Table des matières
    story.append(Paragraph("TABLE DES MATIÈRES", subtitle_style))
    story.append(Spacer(1, 10))
    toc_items = [
        "1. À quoi sert ce logiciel ?",
        "2. Prérequis et installation",
        "3. Explication détaillée de chaque onglet",
        "4. Conseils sur la saisie (majuscules, espaces, etc.)",
        "5. Exemples de saisie correcte et incorrecte",
        "6. Génération de documents",
        "7. Structure des fichiers et schéma dossier",
        "8. Erreurs fréquentes et astuces",
        "9. Support et contact"
    ]
    for item in toc_items:
        story.append(Paragraph(f"• {item}", normal_style))
    story.append(PageBreak())

    # Section 1: À quoi sert ce logiciel
    story.append(Paragraph("1. À QUOI SERT CE LOGICIEL ?", subtitle_style))
    story.append(Paragraph("""
    AUTO RCI est un logiciel qui automatise la création de documents de contrôle qualité à partir de fichiers Excel et de modèles Word. Il est conçu pour être utilisé par tous, même sans connaissances informatiques avancées.
    """, normal_style))
    story.append(PageBreak())

    # Section 2: Prérequis et installation
    story.append(Paragraph("2. PRÉREQUIS ET INSTALLATION", subtitle_style))
    story.append(Paragraph("Logiciels requis :", section_style))
    requirements = [
        "• Windows 10 ou plus récent",
        "• Microsoft Excel (pour lire les fichiers de mesures)",
        "• Microsoft Word (pour générer les documents)",
        "• Python 3.7 ou plus récent"
    ]
    for req in requirements:
        story.append(Paragraph(req, normal_style))
    story.append(Paragraph("Installation :", section_style))
    story.append(Paragraph("""
    1. Vérifiez que tous les logiciels ci-dessus sont installés
    2. Ouvrez le dossier du logiciel AUTO RCI
    3. Double-cliquez sur 'lancer_RCI.bat'
    4. L'application se lance automatiquement
    """, normal_style))
    story.append(PageBreak())

    # Section 3: Explication détaillée de chaque onglet
    story.append(Paragraph("3. EXPLICATION DÉTAILLÉE DE CHAQUE ONGLET", subtitle_style))
    # Onglet Paths
    story.append(Paragraph("Onglet 'Paths' (Chemins)", section_style))
    story.append(Paragraph("""
    Cet onglet sert à indiquer au logiciel où trouver vos fichiers :
    - Le dossier principal de travail (où sont vos documents)
    - Le fichier Excel de mesures
    - Le modèle Word à utiliser
    
    Comment l’utiliser :
    - Cliquez sur chaque bouton (“choix chemin mesures”, “choix chemin formulaire”, “choix chemin lancement”)
    - Sélectionnez le bon fichier ou dossier dans la fenêtre qui s’ouvre
    - Vérifiez que le chemin affiché correspond bien à ce que vous voulez
    """, normal_style))
    # Onglet Find Data
    story.append(Paragraph("Onglet 'Find Data'", section_style))
    story.append(Paragraph("""
    Sert à remplir automatiquement les chemins si vous indiquez :
    - Le numéro de lancement (LT)
    - Le nom du formulaire (ex : F458 E)
    - La catégorie (RE, PV, RCI)
    
    Attention aux majuscules et espaces :
    - LT doit toujours commencer par “LT” en majuscules, suivi de chiffres (ex : LT2500414)
    - Formulaire : format “Fxxx E” (ex : F458 E), avec un espace entre le numéro et la lettre
    - Catégorie : choisissez dans la liste, ne tapez pas à la main
    
    Comment l’utiliser :
    - Remplissez chaque champ
    - Cliquez sur “fill”
    - Les chemins se remplissent automatiquement si la structure des dossiers est correcte
    """, normal_style))
    # Onglet Générateur fiches RCI/RE
    story.append(Paragraph("Onglet 'Générateur fiches RCI/RE'", section_style))
    story.append(Paragraph("""
    C’est ici que vous lancez la création des documents Word à partir de vos données.
    
    Comment l’utiliser :
    - Vérifiez que les chemins sont bien remplis (voir onglet Paths)
    - Cliquez sur “générer document”
    - Sélectionnez le ou les SN (numéros de série) à traiter
    - Le logiciel crée les fichiers Word automatiquement
    """, normal_style))
    # Onglet Conversion PDF
    story.append(Paragraph("Onglet 'Conversion PDF'", section_style))
    story.append(Paragraph("""
    Permet de convertir vos fichiers Word ou Excel en PDF.
    
    Comment l’utiliser :
    - Indiquez le dossier à traiter
    - Précisez si besoin un mot-clé obligatoire ou à éviter dans le nom des fichiers
    - Cliquez sur le bouton pour lancer la conversion
    """, normal_style))
    story.append(PageBreak())

    # Section 4: Conseils sur la saisie
    story.append(Paragraph("4. CONSEILS SUR LA SAISIE (MAJUSCULES, ESPACES, ETC.)", subtitle_style))
    story.append(Paragraph("""
    - Respectez les majuscules/minuscules : “LT” et “F” doivent être en majuscules
    - Respectez les espaces : entre “F458” et “E” dans le formulaire, il faut un espace
    - Pas d’espace en trop au début ou à la fin
    - Ne mettez pas d’accent ou de caractère spécial dans les noms de fichiers/dossiers
    - Vérifiez que les chemins affichés sont corrects avant de lancer la génération
    """, normal_style))
    story.append(PageBreak())

    # Section 5: Exemples de saisie correcte et incorrecte
    story.append(Paragraph("5. EXEMPLES DE SAISIE CORRECTE ET INCORRECTE", subtitle_style))
    data = [
        ["Champ", "Exemple correct", "Exemple incorrect"],
        ["LT", "LT2500414", "lt2500414, LT 2500414"],
        ["Formulaire", "F458 E", "f458e, F458E, F 458E"],
        ["Catégorie", "RE", "re, rE, R E"]
    ]
    table = Table(data, colWidths=[80, 150, 200])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
        ('BOTTOMPADDING', (0,0), (-1,0), 8),
        ('BACKGROUND', (0,1), (-1,-1), colors.whitesmoke),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey)
    ]))
    story.append(table)
    story.append(PageBreak())

    # Section 6: Génération de documents
    story.append(Paragraph("6. GÉNÉRATION DE DOCUMENTS", subtitle_style))
    story.append(Paragraph("""
    1. Allez dans l’onglet 'Générateur fiches RCI/RE'
    2. Cliquez sur 'générer document'
    3. Sélectionnez le ou les SN à traiter
    4. Les documents sont créés automatiquement dans le dossier de sortie
    """, normal_style))
    story.append(PageBreak())

    # Section 7: Structure des fichiers et schéma dossier
    story.append(Paragraph("7. STRUCTURE DES FICHIERS ET SCHÉMA DOSSIER", subtitle_style))
    story.append(Paragraph("Structure de dossiers attendue :", section_style))
    structure = [
        "X:/Tracabilite/",
        "└── RCT°-AGS23.159B/",
        "    ├── LT2500414/          ← Dossier de travail",
        "    └── Mesures 23.159.xlsx ← Fichier de mesures"
    ]
    for line in structure:
        story.append(Paragraph(line, code_style))
    story.append(Paragraph("Format du fichier Excel :", section_style))
    story.append(Paragraph("""
    Le fichier Excel doit contenir :
    • Une première ligne avec les en-têtes contenant des tags entre **
    • Une colonne SN obligatoire
    • Des données pour chaque numéro de série
    """, normal_style))
    story.append(Paragraph("Exemple de structure Excel :", normal_style))
    excel_example = [
        "| SN | **RL1** | **RL2** | **IL_BEFORE** | **IL_AFTER** |",
        "|----|---------|---------|---------------|--------------|",
        "| 001 | -45.2 | -42.1 | -0.8 | -0.9 |",
        "| 002 | -48.3 | -44.7 | -0.7 | -0.8 |"
    ]
    for line in excel_example:
        story.append(Paragraph(line, code_style))
    story.append(PageBreak())

    # Section 8: Erreurs fréquentes et astuces
    story.append(Paragraph("8. ERREURS FRÉQUENTES ET ASTUCES", subtitle_style))
    story.append(Paragraph("Erreurs fréquentes :", section_style))
    errors = [
        ("le fichier n’a pas l’extension .docx", "Vous avez choisi un mauvais fichier Word (il doit finir par .docx)"),
        ("le fichier n’a pas l’extension .xlsx ou .xlsm", "Vous avez choisi un mauvais fichier Excel"),
        ("le chemin n’est pas un dossier", "Vous n’avez pas sélectionné un dossier pour le chemin de lancement"),
        ("le chemin vers le LT n’a pas été trouvé", "Le numéro de lancement n’existe pas dans la structure de dossiers attendue")
    ]
    for err, expl in errors:
        story.append(Paragraph(f"<b>{err}</b> : {expl}", normal_style))
    story.append(Paragraph("Astuces pour éviter les problèmes :", section_style))
    tips = [
        "Toujours vérifier les majuscules et les espaces",
        "Ne pas modifier la structure des dossiers sans demander",
        "Fermer Word et Excel avant de lancer le logiciel",
        "Lire les messages d’erreur affichés dans la fenêtre de l’application"
    ]
    for tip in tips:
        story.append(Paragraph(f"• {tip}", normal_style))
    story.append(PageBreak())

    # Section 9: Support et contact
    story.append(Paragraph("9. SUPPORT ET CONTACT", subtitle_style))
    story.append(Paragraph("""
    Pour toute question, demandez à un collègue ou au service informatique.
    Gardez ce guide à portée de main.
    Les logs sont disponibles dans le dossier 'Logs/' :
    • normal.log : Messages généraux
    • problems.log : Erreurs et avertissements
    • verbose.log : Messages détaillés
    """, normal_style))
    story.append(PageBreak())

    # Page de fin
    story.append(Paragraph("FIN DE LA DOCUMENTATION", title_style))
    story.append(Spacer(1, 20))
    story.append(Paragraph("""
    Cette documentation a été générée automatiquement par AUTO RCI.
    Pour toute question ou suggestion d'amélioration, n'hésitez pas à contacter l'équipe de développement.
    """, normal_style))
    
    # Générer le PDF
    try:
        doc.build(story)
        print(f"✅ Documentation PDF générée avec succès : {pdf_filename}")
        print(f"📁 Fichier créé dans : {os.path.abspath(pdf_filename)}")
        return pdf_filename
    except Exception as e:
        print(f"❌ Erreur lors de la génération du PDF : {str(e)}")
        print("📝 Tentative de génération d'un fichier texte...")
        return create_documentation_text()

if __name__ == "__main__":
    print("🔄 Génération de la documentation...")
    result = create_documentation_pdf()
    if result:
        print("🎉 Documentation créée avec succès !")
        print("📖 Vous pouvez maintenant ouvrir le fichier pour consulter la documentation complète.")
    else:
        print("💥 Échec de la génération de la documentation.") 