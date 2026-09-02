import Settings


DEFAULT_THEME = {
    "background_color" : "gray14",
    "outline_color" : "dim gray",
    "entry_color" : "gray25",
    "text_color" : "white smoke",
    "unselected_color" : "gray30",
    "selected_color" : "DodgerBlue4",
    "question_color" : "dark turquoise",
    "error_color" : "tomato",
    "warning_color" : "goldenrod1"
}

DEFAULT_PATHS = {
    "poppler_path": None,
}

DEFAULT_FIELDS = {
    "LT": "",
    "form": "",
    "out_file_category": "RCI",
    "generator_out_filename": "**REF_SEDI**_SN**SN**_RCI",
    "file_reaction_mode": "dupliquer",
    "num_plan": None,
    "TITREPLAN": "Titre du Plan",
    "pdf_input_path": "",
    "pdf_necessary": None,
    "pdf_replace": None,
}

def VerifySettingsCompleteness() :
    Settings.VerifyCompleteness("theme", DEFAULT_THEME)
    Settings.VerifyCompleteness("paths", DEFAULT_PATHS)
    Settings.VerifyCompleteness("fields", DEFAULT_FIELDS)