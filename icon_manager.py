import customtkinter as ctk
from tkinter import filedialog
import tkinter as tk
from PIL import Image, ImageTk
import os
import sys
import json
import shutil
import shlex
from datetime import datetime

try:
    import winshell
    from win32com.client import Dispatch
    WINDOWS_AVAILABLE = True
except ImportError:
    WINDOWS_AVAILABLE = False

if getattr(sys, 'frozen', False):
    BASE_PATH = os.path.dirname(sys.executable)
else:
    BASE_PATH = os.path.dirname(os.path.abspath(__file__))

RESOURCES_DIR = os.path.join(BASE_PATH, "recursos")
STORAGE_FILE = os.path.join(RESOURCES_DIR, "icon_storage.json")
SETTINGS_FILE = os.path.join(RESOURCES_DIR, "settings.json")

os.makedirs(RESOURCES_DIR, exist_ok=True)

COLORS = {
    "primary":        "#6C63FF",
    "primary_hover":  "#5A52D5",
    "primary_light":  "#8B85FF",
    "success":        "#2DD4BF",
    "success_hover":  "#14B8A6",
    "danger":         "#F43F5E",
    "danger_hover":   "#E11D48",
    "warning":        "#F59E0B",
    "bg_dark":        "#0F0F14",
    "bg_card":        "#1A1A24",
    "bg_sidebar":     "#12121A",
    "bg_input":       "#252533",
    "text_primary":   "#F8FAFC",
    "text_secondary": "#94A3B8",
    "text_muted":     "#64748B",
    "border":         "#2D2D3D",
    "border_focus":   "#6C63FF",
}

LANGUAGES = {
    "en": "English",
    "es": "Español",
    "fr": "Français",
    "pt": "Português",
    "de": "Deutsch",
    "it": "Italiano",
}

TRANSLATIONS = {
    "en": {
        "dashboard": "Dashboard",
        "add": "Add",
        "library": "Library",
        "settings": "Settings",
        "welcome": "Welcome to Icon Manager — manage your shortcuts in style.",
        "total_icons": "Total Icons",
        "quick_action": "Quick Action",
        "create_all_label": "Create all",
        "system": "System",
        "available": "Available",
        "win_only": "Windows only",
        "quick_actions": "Quick Actions",
        "add_new_icon": "Add New Icon",
        "view_library": "View Library",
        "create_all_shortcuts": "Create All Shortcuts",
        "recent_icons": "Recent Icons",
        "only_on_windows": "Only available on Windows",
        "n_shortcuts_created": "{n} shortcuts created successfully",
        "add_new_icon_title": "Add New Icon",
        "add_subtitle": "Register an icon with its target to create desktop shortcuts.",
        "shortcut_icon": "Shortcut Icon",
        "formats": "Formats: .ico  .png  .jpg  .jpeg",
        "browse_file": "Browse File",
        "no_file_selected": "No file selected",
        "target": "Target",
        "target_hint": "URL, application, file, or full command",
        "target_placeholder": 'e.g. https://google.com  or  "C:\\Apps\\app.exe" --arg',
        "shortcut_name": "Shortcut Name",
        "name_placeholder": "e.g. Google Chrome, My App…",
        "save_icon": "Save Icon",
        "save_and_create": "Save & Create Shortcut",
        "convert_failed": "Could not convert image to .ico",
        "select_icon_file": "Please select an icon file",
        "enter_target": "Please enter a target",
        "enter_name": "Please enter a shortcut name",
        "icon_saved": "Icon '{name}' saved",
        "icon_library": "Icon Library",
        "create_selected": "Create Selected",
        "create_all_btn": "Create All",
        "search_placeholder": "Search icons…",
        "no_icons_stored": "No icons stored yet",
        "no_results": "No results found",
        "add_first_icon": "Add First Icon",
        "select_at_least_one": "Select at least one icon",
        "n_shortcuts": "{n} shortcuts created",
        "no_icons": "No icons stored",
        "deleted": "'{name}' deleted",
        "delete_icon": "Delete Icon",
        "confirm_delete": "Are you sure you want to delete '{name}'?\nThis can't be undone.",
        "settings_title": "Settings",
        "appearance": "Appearance",
        "app_theme": "App theme",
        "language": "Language",
        "data": "Data",
        "export": "Export",
        "import_btn": "Import",
        "delete_all": "Delete All",
        "about": "About",
        "about_desc": "Fast and clean shortcut manager for Windows.\nBy duohnson.",
        "export_ok": "Data exported successfully",
        "n_imported": "{n} icons imported",
        "all_deleted": "All data deleted",
        "delete_all_title": "Delete All",
        "confirm_delete_all": "Are you sure? All stored icons will be removed.\nThis can't be undone.",
        "cancel": "Cancel",
        "confirm": "Confirm",
        "shortcut_win_only": "Shortcut features are only available on Windows.",
        "icon_not_found": "Icon file '{path}' not found.",
        "shortcut_created": "Shortcut '{name}' created successfully.",
        "shortcut_failed": "Failed to create shortcut: {err}",
        "lang_changed": "Language changed to {lang}",
    },
    "es": {
        "dashboard": "Inicio",
        "add": "Agregar",
        "library": "Biblioteca",
        "settings": "Ajustes",
        "welcome": "Bienvenido a Icon Manager — gestiona tus accesos directos con estilo.",
        "total_icons": "Total de Iconos",
        "quick_action": "Acción Rápida",
        "create_all_label": "Crear todos",
        "system": "Sistema",
        "available": "Disponible",
        "win_only": "Solo Windows",
        "quick_actions": "Acciones Rápidas",
        "add_new_icon": "Agregar Icono",
        "view_library": "Ver Biblioteca",
        "create_all_shortcuts": "Crear Todos los Atajos",
        "recent_icons": "Iconos Recientes",
        "only_on_windows": "Solo disponible en Windows",
        "n_shortcuts_created": "{n} atajos creados exitosamente",
        "add_new_icon_title": "Agregar Nuevo Icono",
        "add_subtitle": "Registra un icono con su destino para crear accesos directos.",
        "shortcut_icon": "Icono del Atajo",
        "formats": "Formatos: .ico  .png  .jpg  .jpeg",
        "browse_file": "Buscar Archivo",
        "no_file_selected": "Ningún archivo seleccionado",
        "target": "Destino",
        "target_hint": "URL, aplicación, archivo o comando completo",
        "target_placeholder": 'ej. https://google.com  o  "C:\\Apps\\app.exe" --arg',
        "shortcut_name": "Nombre del Atajo",
        "name_placeholder": "ej. Google Chrome, Mi App…",
        "save_icon": "Guardar Icono",
        "save_and_create": "Guardar y Crear Atajo",
        "convert_failed": "No se pudo convertir la imagen a .ico",
        "select_icon_file": "Por favor selecciona un archivo de icono",
        "enter_target": "Por favor ingresa un destino",
        "enter_name": "Por favor ingresa un nombre",
        "icon_saved": "Icono '{name}' guardado",
        "icon_library": "Biblioteca de Iconos",
        "create_selected": "Crear Seleccionados",
        "create_all_btn": "Crear Todos",
        "search_placeholder": "Buscar iconos…",
        "no_icons_stored": "Aún no hay iconos guardados",
        "no_results": "Sin resultados",
        "add_first_icon": "Agregar Primer Icono",
        "select_at_least_one": "Selecciona al menos un icono",
        "n_shortcuts": "{n} atajos creados",
        "no_icons": "No hay iconos guardados",
        "deleted": "'{name}' eliminado",
        "delete_icon": "Eliminar Icono",
        "confirm_delete": "¿Estás seguro de eliminar '{name}'?\nEsta acción no se puede deshacer.",
        "settings_title": "Ajustes",
        "appearance": "Apariencia",
        "app_theme": "Tema de la app",
        "language": "Idioma",
        "data": "Datos",
        "export": "Exportar",
        "import_btn": "Importar",
        "delete_all": "Eliminar Todo",
        "about": "Acerca de",
        "about_desc": "Gestor de accesos directos para Windows, rápido y limpio.\nPor duohnson.",
        "export_ok": "Datos exportados exitosamente",
        "n_imported": "{n} iconos importados",
        "all_deleted": "Todos los datos eliminados",
        "delete_all_title": "Eliminar Todo",
        "confirm_delete_all": "¿Estás seguro? Todos los iconos almacenados serán eliminados.\nNo se puede deshacer.",
        "cancel": "Cancelar",
        "confirm": "Confirmar",
        "shortcut_win_only": "Las funciones de acceso directo solo están disponibles en Windows.",
        "icon_not_found": "Archivo de icono '{path}' no encontrado.",
        "shortcut_created": "Acceso directo '{name}' creado exitosamente.",
        "shortcut_failed": "Error al crear acceso directo: {err}",
        "lang_changed": "Idioma cambiado a {lang}",
    },
    "fr": {
        "dashboard": "Tableau de bord",
        "add": "Ajouter",
        "library": "Bibliothèque",
        "settings": "Paramètres",
        "welcome": "Bienvenue sur Icon Manager — gérez vos raccourcis avec style.",
        "total_icons": "Total d'icônes",
        "quick_action": "Action rapide",
        "create_all_label": "Créer tout",
        "system": "Système",
        "available": "Disponible",
        "win_only": "Windows uniquement",
        "quick_actions": "Actions rapides",
        "add_new_icon": "Ajouter une icône",
        "view_library": "Voir la bibliothèque",
        "create_all_shortcuts": "Créer tous les raccourcis",
        "recent_icons": "Icônes récentes",
        "only_on_windows": "Disponible uniquement sur Windows",
        "n_shortcuts_created": "{n} raccourcis créés avec succès",
        "add_new_icon_title": "Ajouter une icône",
        "add_subtitle": "Enregistrez une icône avec sa cible pour créer des raccourcis.",
        "shortcut_icon": "Icône du raccourci",
        "formats": "Formats : .ico  .png  .jpg  .jpeg",
        "browse_file": "Parcourir",
        "no_file_selected": "Aucun fichier sélectionné",
        "target": "Cible",
        "target_hint": "URL, application, fichier ou commande complète",
        "target_placeholder": 'ex. https://google.com  ou  "C:\\Apps\\app.exe" --arg',
        "shortcut_name": "Nom du raccourci",
        "name_placeholder": "ex. Google Chrome, Mon App…",
        "save_icon": "Enregistrer",
        "save_and_create": "Enregistrer et créer",
        "convert_failed": "Impossible de convertir l'image en .ico",
        "select_icon_file": "Veuillez sélectionner un fichier icône",
        "enter_target": "Veuillez entrer une cible",
        "enter_name": "Veuillez entrer un nom",
        "icon_saved": "Icône '{name}' enregistrée",
        "icon_library": "Bibliothèque d'icônes",
        "create_selected": "Créer la sélection",
        "create_all_btn": "Créer tout",
        "search_placeholder": "Rechercher…",
        "no_icons_stored": "Aucune icône enregistrée",
        "no_results": "Aucun résultat",
        "add_first_icon": "Ajouter la première icône",
        "select_at_least_one": "Sélectionnez au moins une icône",
        "n_shortcuts": "{n} raccourcis créés",
        "no_icons": "Aucune icône enregistrée",
        "deleted": "'{name}' supprimé",
        "delete_icon": "Supprimer l'icône",
        "confirm_delete": "Voulez-vous vraiment supprimer '{name}' ?\nCette action est irréversible.",
        "settings_title": "Paramètres",
        "appearance": "Apparence",
        "app_theme": "Thème de l'app",
        "language": "Langue",
        "data": "Données",
        "export": "Exporter",
        "import_btn": "Importer",
        "delete_all": "Tout supprimer",
        "about": "À propos",
        "about_desc": "Gestionnaire de raccourcis pour Windows, rapide et élégant.\nPar duohnson.",
        "export_ok": "Données exportées avec succès",
        "n_imported": "{n} icônes importées",
        "all_deleted": "Toutes les données supprimées",
        "delete_all_title": "Tout supprimer",
        "confirm_delete_all": "Êtes-vous sûr ? Toutes les icônes seront supprimées.\nCette action est irréversible.",
        "cancel": "Annuler",
        "confirm": "Confirmer",
        "shortcut_win_only": "Les raccourcis ne sont disponibles que sur Windows.",
        "icon_not_found": "Fichier icône '{path}' introuvable.",
        "shortcut_created": "Raccourci '{name}' créé avec succès.",
        "shortcut_failed": "Échec de la création du raccourci : {err}",
        "lang_changed": "Langue changée en {lang}",
    },
    "pt": {
        "dashboard": "Painel",
        "add": "Adicionar",
        "library": "Biblioteca",
        "settings": "Configurações",
        "welcome": "Bem-vindo ao Icon Manager — gerencie seus atalhos com estilo.",
        "total_icons": "Total de Ícones",
        "quick_action": "Ação Rápida",
        "create_all_label": "Criar todos",
        "system": "Sistema",
        "available": "Disponível",
        "win_only": "Somente Windows",
        "quick_actions": "Ações Rápidas",
        "add_new_icon": "Adicionar Ícone",
        "view_library": "Ver Biblioteca",
        "create_all_shortcuts": "Criar Todos os Atalhos",
        "recent_icons": "Ícones Recentes",
        "only_on_windows": "Disponível apenas no Windows",
        "n_shortcuts_created": "{n} atalhos criados com sucesso",
        "add_new_icon_title": "Adicionar Novo Ícone",
        "add_subtitle": "Registre um ícone com seu destino para criar atalhos na área de trabalho.",
        "shortcut_icon": "Ícone do Atalho",
        "formats": "Formatos: .ico  .png  .jpg  .jpeg",
        "browse_file": "Procurar Arquivo",
        "no_file_selected": "Nenhum arquivo selecionado",
        "target": "Destino",
        "target_hint": "URL, aplicativo, arquivo ou comando completo",
        "target_placeholder": 'ex. https://google.com  ou  "C:\\Apps\\app.exe" --arg',
        "shortcut_name": "Nome do Atalho",
        "name_placeholder": "ex. Google Chrome, Meu App…",
        "save_icon": "Salvar Ícone",
        "save_and_create": "Salvar e Criar Atalho",
        "convert_failed": "Não foi possível converter a imagem para .ico",
        "select_icon_file": "Selecione um arquivo de ícone",
        "enter_target": "Insira um destino",
        "enter_name": "Insira um nome para o atalho",
        "icon_saved": "Ícone '{name}' salvo",
        "icon_library": "Biblioteca de Ícones",
        "create_selected": "Criar Selecionados",
        "create_all_btn": "Criar Todos",
        "search_placeholder": "Pesquisar ícones…",
        "no_icons_stored": "Nenhum ícone armazenado",
        "no_results": "Nenhum resultado encontrado",
        "add_first_icon": "Adicionar Primeiro Ícone",
        "select_at_least_one": "Selecione pelo menos um ícone",
        "n_shortcuts": "{n} atalhos criados",
        "no_icons": "Nenhum ícone armazenado",
        "deleted": "'{name}' excluído",
        "delete_icon": "Excluir Ícone",
        "confirm_delete": "Tem certeza que deseja excluir '{name}'?\nEsta ação não pode ser desfeita.",
        "settings_title": "Configurações",
        "appearance": "Aparência",
        "app_theme": "Tema do app",
        "language": "Idioma",
        "data": "Dados",
        "export": "Exportar",
        "import_btn": "Importar",
        "delete_all": "Excluir Tudo",
        "about": "Sobre",
        "about_desc": "Gerenciador de atalhos para Windows, rápido e limpo.\nPor duohnson.",
        "export_ok": "Dados exportados com sucesso",
        "n_imported": "{n} ícones importados",
        "all_deleted": "Todos os dados excluídos",
        "delete_all_title": "Excluir Tudo",
        "confirm_delete_all": "Tem certeza? Todos os ícones armazenados serão removidos.\nEsta ação é irreversível.",
        "cancel": "Cancelar",
        "confirm": "Confirmar",
        "shortcut_win_only": "Recursos de atalho disponíveis apenas no Windows.",
        "icon_not_found": "Arquivo de ícone '{path}' não encontrado.",
        "shortcut_created": "Atalho '{name}' criado com sucesso.",
        "shortcut_failed": "Falha ao criar atalho: {err}",
        "lang_changed": "Idioma alterado para {lang}",
    },
    "de": {
        "dashboard": "Übersicht",
        "add": "Hinzufügen",
        "library": "Bibliothek",
        "settings": "Einstellungen",
        "welcome": "Willkommen bei Icon Manager — verwalte deine Verknüpfungen stilvoll.",
        "total_icons": "Gesamte Icons",
        "quick_action": "Schnellaktion",
        "create_all_label": "Alle erstellen",
        "system": "System",
        "available": "Verfügbar",
        "win_only": "Nur Windows",
        "quick_actions": "Schnellaktionen",
        "add_new_icon": "Icon hinzufügen",
        "view_library": "Bibliothek anzeigen",
        "create_all_shortcuts": "Alle Verknüpfungen erstellen",
        "recent_icons": "Letzte Icons",
        "only_on_windows": "Nur unter Windows verfügbar",
        "n_shortcuts_created": "{n} Verknüpfungen erfolgreich erstellt",
        "add_new_icon_title": "Neues Icon hinzufügen",
        "add_subtitle": "Registriere ein Icon mit Ziel, um Desktop-Verknüpfungen zu erstellen.",
        "shortcut_icon": "Verknüpfungssymbol",
        "formats": "Formate: .ico  .png  .jpg  .jpeg",
        "browse_file": "Datei wählen",
        "no_file_selected": "Keine Datei ausgewählt",
        "target": "Ziel",
        "target_hint": "URL, Anwendung, Datei oder vollständiger Befehl",
        "target_placeholder": 'z.B. https://google.com  oder  "C:\\Apps\\app.exe" --arg',
        "shortcut_name": "Verknüpfungsname",
        "name_placeholder": "z.B. Google Chrome, Meine App…",
        "save_icon": "Speichern",
        "save_and_create": "Speichern & erstellen",
        "convert_failed": "Bild konnte nicht in .ico konvertiert werden",
        "select_icon_file": "Bitte wähle eine Icon-Datei",
        "enter_target": "Bitte gib ein Ziel ein",
        "enter_name": "Bitte gib einen Namen ein",
        "icon_saved": "Icon '{name}' gespeichert",
        "icon_library": "Icon-Bibliothek",
        "create_selected": "Auswahl erstellen",
        "create_all_btn": "Alle erstellen",
        "search_placeholder": "Icons suchen…",
        "no_icons_stored": "Noch keine Icons gespeichert",
        "no_results": "Keine Ergebnisse",
        "add_first_icon": "Erstes Icon hinzufügen",
        "select_at_least_one": "Wähle mindestens ein Icon",
        "n_shortcuts": "{n} Verknüpfungen erstellt",
        "no_icons": "Keine Icons gespeichert",
        "deleted": "'{name}' gelöscht",
        "delete_icon": "Icon löschen",
        "confirm_delete": "Möchtest du '{name}' wirklich löschen?\nDies kann nicht rückgängig gemacht werden.",
        "settings_title": "Einstellungen",
        "appearance": "Aussehen",
        "app_theme": "App-Thema",
        "language": "Sprache",
        "data": "Daten",
        "export": "Exportieren",
        "import_btn": "Importieren",
        "delete_all": "Alles löschen",
        "about": "Über",
        "about_desc": "Schneller und sauberer Verknüpfungsmanager für Windows.\nVon duohnson.",
        "export_ok": "Daten erfolgreich exportiert",
        "n_imported": "{n} Icons importiert",
        "all_deleted": "Alle Daten gelöscht",
        "delete_all_title": "Alles löschen",
        "confirm_delete_all": "Bist du sicher? Alle gespeicherten Icons werden entfernt.\nDies kann nicht rückgängig gemacht werden.",
        "cancel": "Abbrechen",
        "confirm": "Bestätigen",
        "shortcut_win_only": "Verknüpfungsfunktionen sind nur unter Windows verfügbar.",
        "icon_not_found": "Icon-Datei '{path}' nicht gefunden.",
        "shortcut_created": "Verknüpfung '{name}' erfolgreich erstellt.",
        "shortcut_failed": "Verknüpfung konnte nicht erstellt werden: {err}",
        "lang_changed": "Sprache geändert zu {lang}",
    },
    "it": {
        "dashboard": "Pannello",
        "add": "Aggiungi",
        "library": "Libreria",
        "settings": "Impostazioni",
        "welcome": "Benvenuto su Icon Manager — gestisci le tue scorciatoie con stile.",
        "total_icons": "Icone totali",
        "quick_action": "Azione rapida",
        "create_all_label": "Crea tutti",
        "system": "Sistema",
        "available": "Disponibile",
        "win_only": "Solo Windows",
        "quick_actions": "Azioni rapide",
        "add_new_icon": "Aggiungi icona",
        "view_library": "Vedi libreria",
        "create_all_shortcuts": "Crea tutte le scorciatoie",
        "recent_icons": "Icone recenti",
        "only_on_windows": "Disponibile solo su Windows",
        "n_shortcuts_created": "{n} scorciatoie create con successo",
        "add_new_icon_title": "Aggiungi nuova icona",
        "add_subtitle": "Registra un'icona con la sua destinazione per creare scorciatoie.",
        "shortcut_icon": "Icona scorciatoia",
        "formats": "Formati: .ico  .png  .jpg  .jpeg",
        "browse_file": "Sfoglia",
        "no_file_selected": "Nessun file selezionato",
        "target": "Destinazione",
        "target_hint": "URL, applicazione, file o comando completo",
        "target_placeholder": 'es. https://google.com  o  "C:\\Apps\\app.exe" --arg',
        "shortcut_name": "Nome scorciatoia",
        "name_placeholder": "es. Google Chrome, La mia App…",
        "save_icon": "Salva icona",
        "save_and_create": "Salva e crea",
        "convert_failed": "Impossibile convertire l'immagine in .ico",
        "select_icon_file": "Seleziona un file icona",
        "enter_target": "Inserisci una destinazione",
        "enter_name": "Inserisci un nome",
        "icon_saved": "Icona '{name}' salvata",
        "icon_library": "Libreria icone",
        "create_selected": "Crea selezionati",
        "create_all_btn": "Crea tutti",
        "search_placeholder": "Cerca icone…",
        "no_icons_stored": "Nessuna icona salvata",
        "no_results": "Nessun risultato",
        "add_first_icon": "Aggiungi prima icona",
        "select_at_least_one": "Seleziona almeno un'icona",
        "n_shortcuts": "{n} scorciatoie create",
        "no_icons": "Nessuna icona salvata",
        "deleted": "'{name}' eliminato",
        "delete_icon": "Elimina icona",
        "confirm_delete": "Sei sicuro di voler eliminare '{name}'?\nQuesta azione è irreversibile.",
        "settings_title": "Impostazioni",
        "appearance": "Aspetto",
        "app_theme": "Tema dell'app",
        "language": "Lingua",
        "data": "Dati",
        "export": "Esporta",
        "import_btn": "Importa",
        "delete_all": "Elimina tutto",
        "about": "Info",
        "about_desc": "Gestore di scorciatoie per Windows, veloce e pulito.\nDi duohnson.",
        "export_ok": "Dati esportati con successo",
        "n_imported": "{n} icone importate",
        "all_deleted": "Tutti i dati eliminati",
        "delete_all_title": "Elimina tutto",
        "confirm_delete_all": "Sei sicuro? Tutte le icone salvate verranno rimosse.\nQuesta azione è irreversibile.",
        "cancel": "Annulla",
        "confirm": "Conferma",
        "shortcut_win_only": "Le scorciatoie sono disponibili solo su Windows.",
        "icon_not_found": "File icona '{path}' non trovato.",
        "shortcut_created": "Scorciatoia '{name}' creata con successo.",
        "shortcut_failed": "Errore nella creazione della scorciatoia: {err}",
        "lang_changed": "Lingua cambiata in {lang}",
    },
}

_current_lang = "en"


def t(key, **kwargs):
    text = TRANSLATIONS.get(_current_lang, TRANSLATIONS["en"]).get(key)
    if text is None:
        text = TRANSLATIONS["en"].get(key, key)
    if kwargs:
        text = text.format(**kwargs)
    return text


def set_lang(code):
    global _current_lang
    if code in TRANSLATIONS:
        _current_lang = code


def load_storage():
    if not os.path.exists(STORAGE_FILE):
        return {}
    try:
        with open(STORAGE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            for name, info in data.items():
                if not os.path.isabs(info.get("ruta", "")):
                    info["ruta"] = os.path.abspath(os.path.join(RESOURCES_DIR, info["ruta"]))
                if "url" in info and "destino" not in info:
                    info["destino"] = info.pop("url")
            return data
    except (json.JSONDecodeError, KeyError):
        return {}


def save_storage(data):
    out = {}
    for name, info in data.items():
        out[name] = {
            "ruta": os.path.relpath(info["ruta"], start=RESOURCES_DIR),
            "destino": info.get("destino", ""),
            "created": info.get("created", datetime.now().isoformat()),
        }
    os.makedirs(RESOURCES_DIR, exist_ok=True)
    with open(STORAGE_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, indent=4, ensure_ascii=False)


def load_settings():
    defaults = {"theme": "dark", "language": "en"}
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                defaults.update(json.load(f))
        except Exception:
            pass
    return defaults


def save_settings(settings):
    os.makedirs(RESOURCES_DIR, exist_ok=True)
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=4)


def convert_to_ico(image_path):
    try:
        output = os.path.splitext(image_path)[0] + ".ico"
        img = Image.open(image_path)
        img.save(output, format="ICO",
                 sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
        return output
    except Exception:
        return None


def create_shortcut(name, icon_path, destination):
    if not WINDOWS_AVAILABLE:
        return False, t("shortcut_win_only")
    try:
        os.makedirs(RESOURCES_DIR, exist_ok=True)

        icon_name = os.path.basename(icon_path)
        res_icon = os.path.join(RESOURCES_DIR, icon_name)
        if not os.path.exists(res_icon):
            shutil.copy(icon_path, res_icon)

        pictures_dir = os.path.join(os.path.expanduser("~"), "Pictures", "recursos")
        os.makedirs(pictures_dir, exist_ok=True)
        final_icon = os.path.join(pictures_dir, icon_name)
        if not os.path.exists(final_icon):
            shutil.copy(res_icon, final_icon)
        final_icon = os.path.abspath(final_icon)

        if not os.path.isfile(final_icon):
            return False, t("icon_not_found", path=final_icon)

        desktop = winshell.desktop()
        shortcut_path = os.path.join(desktop, f"{name}.lnk")

        parts = shlex.split(destination)
        target = parts[0] if parts else destination
        arguments = " ".join(parts[1:]) if len(parts) > 1 else ""

        shell = Dispatch("WScript.Shell")
        sc = shell.CreateShortCut(shortcut_path)
        sc.TargetPath = target
        if arguments:
            sc.Arguments = arguments
        sc.IconLocation = final_icon
        sc.save()

        return True, t("shortcut_created", name=name)
    except Exception as e:
        return False, t("shortcut_failed", err=str(e))


class Toast(ctk.CTkFrame):

    def __init__(self, parent, message, toast_type="info", duration=3000):
        accent_map = {
            "success": (COLORS["success"], "#064E3B"),
            "error":   (COLORS["danger"],  "#4C0519"),
            "info":    (COLORS["primary"], "#1E1B4B"),
            "warning": (COLORS["warning"], "#78350F"),
        }
        accent, bg = accent_map.get(toast_type, accent_map["info"])

        super().__init__(parent, fg_color=bg, corner_radius=12,
                         border_width=1, border_color=accent)

        icon_map = {"success": "", "error": "", "info": "", "warning": ""}
        icon = icon_map.get(toast_type, "")

        ctk.CTkLabel(self, text=icon, font=ctk.CTkFont(size=18, weight="bold"),
                     text_color=accent, width=30).pack(side="left", padx=(15, 5), pady=12)

        ctk.CTkLabel(self, text=message, font=ctk.CTkFont(size=13),
                     text_color=COLORS["text_primary"],
                     wraplength=380).pack(side="left", padx=(5, 15), pady=12, fill="x", expand=True)

        self.place(relx=0.5, rely=0.0, anchor="n", y=-60)
        self._target_y = 20
        self._cur_y = -60
        self._slide_in(duration)

    def _slide_in(self, duration):
        if self._cur_y < self._target_y:
            self._cur_y += 5
            self.place_configure(y=self._cur_y)
            self.after(12, lambda: self._slide_in(duration))
        else:
            self.after(duration, self._slide_out)

    def _slide_out(self):
        if self._cur_y > -70:
            self._cur_y -= 5
            self.place_configure(y=self._cur_y)
            self.after(12, self._slide_out)
        else:
            self.destroy()


class ConfirmDialog(ctk.CTkToplevel):

    def __init__(self, parent, title, message, on_confirm=None):
        super().__init__(parent)
        self.on_confirm = on_confirm
        self.title(title)
        self.geometry("440x210")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 440) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 210) // 2
        self.geometry(f"+{x}+{y}")

        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=30, pady=25)

        ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(size=18, weight="bold"),
                     text_color=COLORS["text_primary"]).pack(anchor="w")

        ctk.CTkLabel(frame, text=message, font=ctk.CTkFont(size=13),
                     text_color=COLORS["text_secondary"],
                     wraplength=380, justify="left").pack(anchor="w", pady=(10, 20))

        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(fill="x")

        ctk.CTkButton(btn_frame, text=t("cancel"), width=110, height=36,
                      fg_color="transparent", border_width=1,
                      border_color=COLORS["border"],
                      text_color=COLORS["text_secondary"],
                      hover_color=COLORS["bg_input"],
                      command=self.destroy).pack(side="right", padx=(10, 0))

        ctk.CTkButton(btn_frame, text=t("confirm"), width=110, height=36,
                      fg_color=COLORS["danger"],
                      hover_color=COLORS["danger_hover"],
                      command=self._do_confirm).pack(side="right")

    def _do_confirm(self):
        if self.on_confirm:
            self.on_confirm()
        self.destroy()


class SidebarButton(ctk.CTkButton):

    def __init__(self, parent, text, icon, command, **kw):
        label = f"  {icon}  {text}" if icon else text
        super().__init__(parent, text=label, command=command,
                         height=44, corner_radius=10,
                         font=ctk.CTkFont(size=14), anchor="w",
                         fg_color="transparent",
                         text_color=COLORS["text_secondary"],
                         hover_color=COLORS["bg_input"], **kw)

    def set_active(self, active):
        if active:
            self.configure(fg_color=COLORS["primary"],
                          text_color=COLORS["text_primary"],
                          hover_color=COLORS["primary_hover"])
        else:
            self.configure(fg_color="transparent",
                          text_color=COLORS["text_secondary"],
                          hover_color=COLORS["bg_input"])


class IconCard(ctk.CTkFrame):

    def __init__(self, parent, name, info, *, on_select=None, on_delete=None, on_create=None):
        super().__init__(parent, fg_color=COLORS["bg_card"], corner_radius=14,
                         border_width=1, border_color=COLORS["border"], cursor="hand2")
        self.name = name
        self.info = info
        self.selected = False
        self._on_select_cb = on_select

        self.configure(width=230, height=210)
        self.pack_propagate(False)

        inner = ctk.CTkFrame(self, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=16, pady=14)

        pf = ctk.CTkFrame(inner, fg_color=COLORS["bg_input"],
                          corner_radius=12, height=64, width=64)
        pf.pack(anchor="center", pady=(0, 10))
        pf.pack_propagate(False)
        self._load_preview(pf, info.get("ruta", ""))

        ctk.CTkLabel(inner, text=name, font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=COLORS["text_primary"]).pack(anchor="center")

        dest = info.get("destino", "")
        if len(dest) > 35:
            dest = dest[:32] + "…"
        ctk.CTkLabel(inner, text=dest, font=ctk.CTkFont(size=11),
                     text_color=COLORS["text_muted"]).pack(anchor="center", pady=(2, 8))

        act = ctk.CTkFrame(inner, fg_color="transparent")
        act.pack(fill="x")

        if on_create:
            ctk.CTkButton(act, text="Create", width=70, height=28, corner_radius=8,
                          fg_color=COLORS["success"], hover_color=COLORS["success_hover"],
                          font=ctk.CTkFont(size=13),
                          command=lambda: on_create(name)).pack(side="left", padx=(0, 4))

        if on_delete:
            ctk.CTkButton(act, text="Delete", width=70, height=28, corner_radius=8,
                          fg_color=COLORS["danger"], hover_color=COLORS["danger_hover"],
                          font=ctk.CTkFont(size=13),
                          command=lambda: on_delete(name)).pack(side="right")

        self.check_var = ctk.BooleanVar(value=False)
        self.checkbox = ctk.CTkCheckBox(act, text="", variable=self.check_var, width=24,
                                        fg_color=COLORS["primary"],
                                        hover_color=COLORS["primary_hover"],
                                        border_color=COLORS["border"], corner_radius=6,
                                        command=self._on_toggle)
        self.checkbox.pack(side="right", padx=(0, 8))

        self.bind("<Button-1>", lambda _: self._toggle())
        inner.bind("<Button-1>", lambda _: self._toggle())

    def _load_preview(self, frame, icon_path):
        try:
            if os.path.isfile(icon_path):
                img = Image.open(icon_path).resize((40, 40), Image.LANCZOS)
                self._photo = ctk.CTkImage(light_image=img, dark_image=img, size=(40, 40))
                ctk.CTkLabel(frame, image=self._photo, text="").pack(expand=True)
            else:
                raise FileNotFoundError
        except Exception:
            ctk.CTkLabel(frame, text="No icon", font=ctk.CTkFont(size=16, weight="bold"),
                         text_color=COLORS["text_muted"]).pack(expand=True)

    def _toggle(self):
        self.check_var.set(not self.check_var.get())
        self._on_toggle()

    def _on_toggle(self):
        self.selected = self.check_var.get()
        border = COLORS["primary"] if self.selected else COLORS["border"]
        self.configure(border_color=border)
        if self._on_select_cb:
            self._on_select_cb()


class DashboardPage(ctk.CTkScrollableFrame):

    def __init__(self, parent, app):
        super().__init__(parent, fg_color="transparent",
                         scrollbar_button_color=COLORS["bg_input"],
                         scrollbar_button_hover_color=COLORS["primary"])
        self.app = app
        self._build()

    def _build(self):
        total = len(self.app.data)

        ctk.CTkLabel(self, text=t("dashboard"), font=ctk.CTkFont(size=28, weight="bold"),
                     text_color=COLORS["text_primary"]).pack(anchor="w", padx=30, pady=(30, 0))

        ctk.CTkLabel(self, text=t("welcome"),
                     font=ctk.CTkFont(size=14),
                     text_color=COLORS["text_secondary"]).pack(anchor="w", padx=30, pady=(4, 0))

        stats = ctk.CTkFrame(self, fg_color="transparent")
        stats.pack(fill="x", padx=30, pady=(25, 0))

        self._stat_card(stats, "", t("total_icons"), str(total),
                       COLORS["primary"]).pack(side="left", padx=(0, 15))

        self._stat_card(stats, "", t("quick_action"), t("create_all_label"),
                       COLORS["success"]).pack(side="left", padx=(0, 15))

        status_text = t("available") if WINDOWS_AVAILABLE else t("win_only")
        status_col = COLORS["success"] if WINDOWS_AVAILABLE else COLORS["warning"]
        self._stat_card(stats, "", t("system"), status_text,
                       status_col).pack(side="left")

        ctk.CTkLabel(self, text=t("quick_actions"), font=ctk.CTkFont(size=18, weight="bold"),
                     text_color=COLORS["text_primary"]).pack(anchor="w", padx=30, pady=(30, 12))

        actions = ctk.CTkFrame(self, fg_color="transparent")
        actions.pack(fill="x", padx=30)

        ctk.CTkButton(actions, text=t("add_new_icon"), height=44, corner_radius=10,
                      fg_color=COLORS["primary"], hover_color=COLORS["primary_hover"],
                      font=ctk.CTkFont(size=14),
                      command=lambda: self.app.show_page("add")).pack(side="left", padx=(0, 12))

        ctk.CTkButton(actions, text=t("view_library"), height=44, corner_radius=10,
                      fg_color=COLORS["bg_card"], hover_color=COLORS["bg_input"],
                      border_width=1, border_color=COLORS["border"],
                      font=ctk.CTkFont(size=14), text_color=COLORS["text_primary"],
                      command=lambda: self.app.show_page("library")).pack(side="left", padx=(0, 12))

        if total > 0:
            ctk.CTkButton(actions, text=t("create_all_shortcuts"), height=44, corner_radius=10,
                          fg_color=COLORS["success"], hover_color=COLORS["success_hover"],
                          font=ctk.CTkFont(size=14),
                          command=self._create_all).pack(side="left")

        if total > 0:
            ctk.CTkLabel(self, text=t("recent_icons"), font=ctk.CTkFont(size=18, weight="bold"),
                         text_color=COLORS["text_primary"]).pack(anchor="w", padx=30, pady=(30, 12))

            recent = ctk.CTkFrame(self, fg_color="transparent")
            recent.pack(fill="x", padx=30, pady=(0, 30))

            for name, info in list(self.app.data.items())[-5:]:
                self._recent_item(recent, name, info).pack(fill="x", pady=(0, 6))

    def _stat_card(self, parent, icon, title, value, accent):
        card = ctk.CTkFrame(parent, fg_color=COLORS["bg_card"],
                           corner_radius=14, width=210, height=105)
        card.pack_propagate(False)
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=18, pady=14)

        top = ctk.CTkFrame(inner, fg_color="transparent")
        top.pack(fill="x")
        ctk.CTkLabel(top, text=icon, font=ctk.CTkFont(size=24)).pack(side="left")
        ctk.CTkLabel(top, text=title, font=ctk.CTkFont(size=12),
                     text_color=COLORS["text_muted"]).pack(side="left", padx=(8, 0))

        ctk.CTkLabel(inner, text=value, font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=accent).pack(anchor="w", pady=(8, 0))
        return card

    def _recent_item(self, parent, name, info):
        item = ctk.CTkFrame(parent, fg_color=COLORS["bg_card"],
                           corner_radius=10, height=50)
        item.pack_propagate(False)
        inner = ctk.CTkFrame(item, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=14, pady=8)

        ctk.CTkLabel(inner, text=name, font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=COLORS["text_primary"]).pack(side="left")

        dest = info.get("destino", "")
        if len(dest) > 60:
            dest = dest[:57] + "…"
        ctk.CTkLabel(inner, text=dest, font=ctk.CTkFont(size=12),
                     text_color=COLORS["text_muted"]).pack(side="right")
        return item

    def _create_all(self):
        if not WINDOWS_AVAILABLE:
            Toast(self.app, t("only_on_windows"), "warning")
            return
        success = 0
        for name, info in self.app.data.items():
            ok, _ = create_shortcut(name, info["ruta"], info.get("destino", ""))
            if ok:
                success += 1
        Toast(self.app, t("n_shortcuts_created", n=success), "success")


class AddPage(ctk.CTkScrollableFrame):

    def __init__(self, parent, app):
        super().__init__(parent, fg_color="transparent",
                         scrollbar_button_color=COLORS["bg_input"],
                         scrollbar_button_hover_color=COLORS["primary"])
        self.app = app
        self.icon_path = ""
        self._build()

    def _build(self):
        ctk.CTkLabel(self, text=t("add_new_icon_title"), font=ctk.CTkFont(size=28, weight="bold"),
                     text_color=COLORS["text_primary"]).pack(anchor="w", padx=30, pady=(30, 0))

        ctk.CTkLabel(self, text=t("add_subtitle"),
                     font=ctk.CTkFont(size=14),
                     text_color=COLORS["text_secondary"]).pack(anchor="w", padx=30, pady=(4, 0))

        card = ctk.CTkFrame(self, fg_color=COLORS["bg_card"], corner_radius=16)
        card.pack(fill="x", padx=30, pady=(25, 30))

        form = ctk.CTkFrame(card, fg_color="transparent")
        form.pack(fill="x", padx=30, pady=30)

        # icon preview + browse
        top = ctk.CTkFrame(form, fg_color="transparent")
        top.pack(fill="x", pady=(0, 20))

        self.preview_frame = ctk.CTkFrame(top, fg_color=COLORS["bg_input"], corner_radius=14,
                                          width=80, height=80, border_width=2,
                                          border_color=COLORS["border"])
        self.preview_frame.pack(side="left")
        self.preview_frame.pack_propagate(False)

        self.preview_label = ctk.CTkLabel(self.preview_frame, text="No preview",
                                          font=ctk.CTkFont(size=14),
                                          text_color=COLORS["text_muted"])
        self.preview_label.pack(expand=True)

        browse_col = ctk.CTkFrame(top, fg_color="transparent")
        browse_col.pack(side="left", padx=(16, 0), fill="y")

        ctk.CTkLabel(browse_col, text=t("shortcut_icon"),
                     font=ctk.CTkFont(size=15, weight="bold"),
                     text_color=COLORS["text_primary"]).pack(anchor="w")

        ctk.CTkLabel(browse_col, text=t("formats"),
                     font=ctk.CTkFont(size=12),
                     text_color=COLORS["text_muted"]).pack(anchor="w", pady=(2, 8))

        ctk.CTkButton(browse_col, text=t("browse_file"), width=170, height=34, corner_radius=8,
                      fg_color=COLORS["primary"], hover_color=COLORS["primary_hover"],
                      font=ctk.CTkFont(size=13),
                      command=self._browse_icon).pack(anchor="w")

        self.path_label = ctk.CTkLabel(form, text="No file selected",
                                       font=ctk.CTkFont(size=12),
                                       text_color=COLORS["text_muted"])
        self.path_label.pack(anchor="w", pady=(0, 18))

        # destination field
        ctk.CTkLabel(form, text="Target", font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=COLORS["text_primary"]).pack(anchor="w")
        ctk.CTkLabel(form, text="URL, application, file, or full command",
                     font=ctk.CTkFont(size=12),
                     text_color=COLORS["text_muted"]).pack(anchor="w", pady=(2, 6))

        self.dest_entry = ctk.CTkEntry(form, height=42, corner_radius=10,
                                       fg_color=COLORS["bg_input"],
                                       border_color=COLORS["border"],
                                       text_color=COLORS["text_primary"],
                                       placeholder_text='e.g. https://google.com  or  "C:\\Apps\\app.exe" --arg',
                                       font=ctk.CTkFont(size=13))
        self.dest_entry.pack(fill="x", pady=(0, 18))

        # name field
        ctk.CTkLabel(form, text="Shortcut Name", font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=COLORS["text_primary"]).pack(anchor="w")

        self.name_entry = ctk.CTkEntry(form, height=42, corner_radius=10,
                                       fg_color=COLORS["bg_input"],
                                       border_color=COLORS["border"],
                                       text_color=COLORS["text_primary"],
                                       placeholder_text="e.g. Google Chrome, My App…",
                                       font=ctk.CTkFont(size=13))
        self.name_entry.pack(fill="x", pady=(6, 22))

        btn_row = ctk.CTkFrame(form, fg_color="transparent")
        btn_row.pack(fill="x")

        ctk.CTkButton(btn_row, text=t("save_icon"), height=44, corner_radius=10,
                      fg_color=COLORS["primary"], hover_color=COLORS["primary_hover"],
                      font=ctk.CTkFont(size=14, weight="bold"),
                      command=self._save_icon).pack(side="left", padx=(0, 12))

        ctk.CTkButton(btn_row, text=t("save_and_create"), height=44, corner_radius=10,
                      fg_color=COLORS["success"], hover_color=COLORS["success_hover"],
                      font=ctk.CTkFont(size=14, weight="bold"),
                      command=self._save_and_create).pack(side="left")

    def _browse_icon(self):
        path = filedialog.askopenfilename(
            filetypes=[("Icon Files", "*.ico"), ("Images", "*.png *.jpg *.jpeg *.bmp"), ("All", "*.*")])
        if not path:
            return
        if not path.lower().endswith(".ico"):
            converted = convert_to_ico(path)
            if converted:
                path = converted
            else:
                Toast(self.app, t("convert_failed"), "error")
                return

        self.icon_path = path
        display = path if len(path) < 60 else "…" + path[-57:]
        self.path_label.configure(text=display, text_color=COLORS["text_primary"])

        try:
            for w in self.preview_frame.winfo_children():
                w.destroy()
            img = Image.open(path).resize((48, 48), Image.LANCZOS)
            self._photo = ctk.CTkImage(light_image=img, dark_image=img, size=(48, 48))
            ctk.CTkLabel(self.preview_frame, image=self._photo, text="").pack(expand=True)
            self.preview_frame.configure(border_color=COLORS["primary"])
        except Exception:
            pass

    def _validate(self):
        if not self.icon_path:
            Toast(self.app, t("select_icon_file"), "error")
            return False
        if not self.dest_entry.get().strip():
            Toast(self.app, t("enter_target"), "error")
            return False
        if not self.name_entry.get().strip():
            Toast(self.app, t("enter_name"), "error")
            return False
        return True

    def _persist(self):
        name = self.name_entry.get().strip()
        icon_name = os.path.basename(self.icon_path)
        dest_icon = os.path.join(RESOURCES_DIR, icon_name)
        if not os.path.exists(dest_icon):
            shutil.copy(self.icon_path, dest_icon)

        self.app.data[name] = {
            "ruta": dest_icon,
            "destino": self.dest_entry.get().strip(),
            "created": datetime.now().isoformat(),
        }
        save_storage(self.app.data)
        return name, dest_icon

    def _save_icon(self):
        if not self._validate():
            return
        name, _ = self._persist()
        Toast(self.app, t("icon_saved", name=name), "success")
        self._clear()

    def _save_and_create(self):
        if not self._validate():
            return
        name, dest_icon = self._persist()
        ok, msg = create_shortcut(name, dest_icon, self.dest_entry.get().strip())
        Toast(self.app, msg, "success" if ok else "error")
        if ok:
            self._clear()

    def _clear(self):
        self.icon_path = ""
        self.dest_entry.delete(0, "end")
        self.name_entry.delete(0, "end")
        self.path_label.configure(text=t("no_file_selected"), text_color=COLORS["text_muted"])
        for w in self.preview_frame.winfo_children():
            w.destroy()
        ctk.CTkLabel(self.preview_frame, text="No preview", font=ctk.CTkFont(size=14),
                     text_color=COLORS["text_muted"]).pack(expand=True)
        self.preview_frame.configure(border_color=COLORS["border"])


class LibraryPage(ctk.CTkFrame):

    def __init__(self, parent, app):
        super().__init__(parent, fg_color="transparent")
        self.app = app
        self.cards: list[IconCard] = []
        self._build()

    def _build(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=30, pady=(30, 0))

        ctk.CTkLabel(header, text=t("icon_library"), font=ctk.CTkFont(size=28, weight="bold"),
                     text_color=COLORS["text_primary"]).pack(side="left")

        batch = ctk.CTkFrame(header, fg_color="transparent")
        batch.pack(side="right")

        ctk.CTkButton(batch, text=t("create_selected"), height=36, corner_radius=8,
                      fg_color=COLORS["success"], hover_color=COLORS["success_hover"],
                      font=ctk.CTkFont(size=13),
                      command=self._create_selected).pack(side="left", padx=(0, 8))

        ctk.CTkButton(batch, text=t("create_all_btn"), height=36, corner_radius=8,
                      fg_color=COLORS["primary"], hover_color=COLORS["primary_hover"],
                      font=ctk.CTkFont(size=13),
                      command=self._create_all).pack(side="left")

        self.search_entry = ctk.CTkEntry(self, height=42, corner_radius=10,
                                         fg_color=COLORS["bg_input"],
                                         border_color=COLORS["border"],
                                         text_color=COLORS["text_primary"],
                                         placeholder_text=t("search_placeholder"),
                                         font=ctk.CTkFont(size=13))
        self.search_entry.pack(fill="x", padx=30, pady=(16, 0))
        self.search_entry.bind("<KeyRelease>", self._on_search)

        self.scroll = ctk.CTkScrollableFrame(self, fg_color="transparent", corner_radius=0,
                                             scrollbar_button_color=COLORS["bg_input"],
                                             scrollbar_button_hover_color=COLORS["primary"])
        self.scroll.pack(fill="both", expand=True, padx=30, pady=(16, 20))

        self._render_cards()

    def _render_cards(self, filter_text=""):
        for w in self.scroll.winfo_children():
            w.destroy()
        self.cards.clear()

        filtered = {
            k: v for k, v in self.app.data.items()
            if filter_text.lower() in k.lower()
            or filter_text.lower() in v.get("destino", "").lower()
        }

        if not filtered:
            empty = ctk.CTkFrame(self.scroll, fg_color="transparent")
            empty.pack(expand=True, pady=60)
            ctk.CTkLabel(empty, text="No items", font=ctk.CTkFont(size=18)).pack()
            msg = t("no_icons_stored") if not filter_text else t("no_results")
            ctk.CTkLabel(empty, text=msg, font=ctk.CTkFont(size=16),
                         text_color=COLORS["text_muted"]).pack(pady=(10, 0))
            if not filter_text:
                ctk.CTkButton(empty, text=t("add_first_icon"), height=40, corner_radius=10,
                              fg_color=COLORS["primary"], hover_color=COLORS["primary_hover"],
                              command=lambda: self.app.show_page("add")).pack(pady=(16, 0))
            return

        row_frame = None
        for i, (name, info) in enumerate(filtered.items()):
            if i % 3 == 0:
                row_frame = ctk.CTkFrame(self.scroll, fg_color="transparent")
                row_frame.pack(fill="x", pady=(0, 12))

            card = IconCard(row_frame, name, info,
                           on_select=self._update_selection,
                           on_delete=self._delete_icon,
                           on_create=self._create_single)
            card.pack(side="left", padx=(0, 12), fill="y")
            self.cards.append(card)

    def _on_search(self, _event=None):
        self._render_cards(self.search_entry.get())

    def _update_selection(self):
        pass

    def _get_selected(self):
        return [c.name for c in self.cards if c.selected]

    def _create_selected(self):
        selected = self._get_selected()
        if not selected:
            Toast(self.app, t("select_at_least_one"), "warning")
            return
        success = 0
        for name in selected:
            info = self.app.data[name]
            ok, _ = create_shortcut(name, info["ruta"], info.get("destino", ""))
            if ok:
                success += 1
        Toast(self.app, t("n_shortcuts", n=success), "success")

    def _create_all(self):
        if not self.app.data:
            Toast(self.app, t("no_icons"), "warning")
            return
        success = 0
        for name, info in self.app.data.items():
            ok, _ = create_shortcut(name, info["ruta"], info.get("destino", ""))
            if ok:
                success += 1
        Toast(self.app, t("n_shortcuts", n=success), "success")

    def _create_single(self, name):
        info = self.app.data.get(name)
        if info:
            ok, msg = create_shortcut(name, info["ruta"], info.get("destino", ""))
            Toast(self.app, msg, "success" if ok else "error")

    def _delete_icon(self, name):
        def do_delete():
            if name in self.app.data:
                del self.app.data[name]
                save_storage(self.app.data)
                Toast(self.app, t("deleted", name=name), "success")
                self._render_cards(self.search_entry.get())

        ConfirmDialog(self.app, t("delete_icon"),
                     t("confirm_delete", name=name),
                     on_confirm=do_delete)

    def refresh(self):
        q = self.search_entry.get() if hasattr(self, "search_entry") else ""
        self._render_cards(q)


class SettingsPage(ctk.CTkScrollableFrame):

    def __init__(self, parent, app):
        super().__init__(parent, fg_color="transparent",
                         scrollbar_button_color=COLORS["bg_input"],
                         scrollbar_button_hover_color=COLORS["primary"])
        self.app = app
        self._build()

    def _build(self):
        ctk.CTkLabel(self, text=t("settings_title"), font=ctk.CTkFont(size=28, weight="bold"),
                     text_color=COLORS["text_primary"]).pack(anchor="w", padx=30, pady=(30, 0))

        self._section(t("appearance"), self._appearance_content)
        self._section(t("language"), self._language_content)
        self._section(t("data"), self._data_content)
        self._section(t("about"), self._about_content)

    def _section(self, title, content_fn):
        card = ctk.CTkFrame(self, fg_color=COLORS["bg_card"], corner_radius=16)
        card.pack(fill="x", padx=30, pady=(20, 0))
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=24, pady=20)
        ctk.CTkLabel(inner, text=title, font=ctk.CTkFont(size=18, weight="bold"),
                     text_color=COLORS["text_primary"]).pack(anchor="w", pady=(0, 12))
        content_fn(inner)

    def _appearance_content(self, parent):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x")
        ctk.CTkLabel(row, text=t("app_theme"), font=ctk.CTkFont(size=14),
                     text_color=COLORS["text_secondary"]).pack(side="left")

        self.theme_menu = ctk.CTkOptionMenu(
            row, values=["Dark", "Light", "System"],
            fg_color=COLORS["bg_input"], button_color=COLORS["primary"],
            button_hover_color=COLORS["primary_hover"],
            dropdown_fg_color=COLORS["bg_card"],
            font=ctk.CTkFont(size=13), width=150,
            command=self._change_theme)
        self.theme_menu.pack(side="right")
        current = self.app.settings.get("theme", "dark").capitalize()
        self.theme_menu.set(current if current in ("Dark", "Light", "System") else "Dark")

    def _language_content(self, parent):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x")
        ctk.CTkLabel(row, text=t("language"), font=ctk.CTkFont(size=14),
                     text_color=COLORS["text_secondary"]).pack(side="left")

        lang_names = list(LANGUAGES.values())
        self.lang_menu = ctk.CTkOptionMenu(
            row, values=lang_names,
            fg_color=COLORS["bg_input"], button_color=COLORS["primary"],
            button_hover_color=COLORS["primary_hover"],
            dropdown_fg_color=COLORS["bg_card"],
            font=ctk.CTkFont(size=13), width=180,
            command=self._change_language)
        self.lang_menu.pack(side="right")

        current_code = self.app.settings.get("language", "en")
        current_name = LANGUAGES.get(current_code, "English")
        self.lang_menu.set(current_name)

    def _data_content(self, parent):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x")

        ctk.CTkButton(row, text=t("export"), height=40, corner_radius=10,
                      fg_color=COLORS["primary"], hover_color=COLORS["primary_hover"],
                      font=ctk.CTkFont(size=13),
                      command=self._export).pack(side="left", padx=(0, 10))

        ctk.CTkButton(row, text=t("import_btn"), height=40, corner_radius=10,
                      fg_color=COLORS["bg_input"], hover_color=COLORS["border"],
                      border_width=1, border_color=COLORS["border"],
                      font=ctk.CTkFont(size=13), text_color=COLORS["text_primary"],
                      command=self._import).pack(side="left", padx=(0, 10))

        ctk.CTkButton(row, text=t("delete_all"), height=40, corner_radius=10,
                      fg_color=COLORS["danger"], hover_color=COLORS["danger_hover"],
                      font=ctk.CTkFont(size=13),
                      command=self._delete_all).pack(side="left")

    def _about_content(self, parent):
        ctk.CTkLabel(parent, text="Icon Manager v2.0",
                     font=ctk.CTkFont(size=15, weight="bold"),
                     text_color=COLORS["primary"]).pack(anchor="w")
        ctk.CTkLabel(parent,
                     text=t("about_desc"),
                     font=ctk.CTkFont(size=13), text_color=COLORS["text_muted"],
                     justify="left").pack(anchor="w", pady=(4, 0))

    def _change_theme(self, value):
        mapping = {"Dark": "dark", "Light": "light", "System": "system"}
        mode = mapping.get(value, "dark")
        ctk.set_appearance_mode(mode)
        self.app.settings["theme"] = mode
        save_settings(self.app.settings)

    def _change_language(self, display_name):
        code = "en"
        for k, v in LANGUAGES.items():
            if v == display_name:
                code = k
                break
        set_lang(code)
        self.app.settings["language"] = code
        save_settings(self.app.settings)
        self.app._refresh_sidebar()
        self.app.show_page("settings")
        Toast(self.app, t("lang_changed", lang=display_name), "success")

    def _export(self):
        path = filedialog.asksaveasfilename(defaultextension=".json",
                                            filetypes=[("JSON", "*.json")],
                                            initialfile="icon_manager_backup.json")
        if path:
            try:
                shutil.copy(STORAGE_FILE, path)
                Toast(self.app, t("export_ok"), "success")
            except Exception as e:
                Toast(self.app, f"Error: {e}", "error")

    def _import(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                imported = json.load(f)
            for name, info in imported.items():
                if not os.path.isabs(info.get("ruta", "")):
                    info["ruta"] = os.path.abspath(os.path.join(RESOURCES_DIR, info["ruta"]))
                if "url" in info and "destino" not in info:
                    info["destino"] = info.pop("url")
                self.app.data[name] = info
            save_storage(self.app.data)
            Toast(self.app, t("n_imported", n=len(imported)), "success")
        except Exception as e:
            Toast(self.app, f"Error: {e}", "error")

    def _delete_all(self):
        def do_delete():
            self.app.data.clear()
            save_storage(self.app.data)
            Toast(self.app, t("all_deleted"), "success")

        ConfirmDialog(self.app, t("delete_all_title"),
                     t("confirm_delete_all"),
                     on_confirm=do_delete)


class IconManagerApp(ctk.CTk):

    NAV_ITEMS = [
        ("dashboard", "dashboard", ""),
        ("add",       "add",       ""),
        ("library",   "library",   ""),
        ("settings",  "settings",  ""),
    ]

    def __init__(self):
        super().__init__()

        self.data = load_storage()
        self.settings = load_settings()

        saved_lang = self.settings.get("language", "en")
        set_lang(saved_lang)

        ctk.set_appearance_mode(self.settings.get("theme", "dark"))
        ctk.set_default_color_theme("blue")

        self.title("Icon Manager")
        self.geometry("1100x740")
        self.minsize(920, 600)
        self.configure(fg_color=COLORS["bg_dark"])

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._current_key = "dashboard"
        self._build_sidebar()
        self._build_content()
        self.show_page("dashboard")

    def _build_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, fg_color=COLORS["bg_sidebar"],
                                    corner_radius=0, width=230)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)

        logo = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        logo.pack(fill="x", padx=18, pady=(24, 30))

        ctk.CTkLabel(logo, text="⬡", font=ctk.CTkFont(size=28, weight="bold"),
                     text_color=COLORS["primary"]).pack(side="left")
        ctk.CTkLabel(logo, text=" Icon Manager", font=ctk.CTkFont(size=18, weight="bold"),
                     text_color=COLORS["text_primary"]).pack(side="left")

        self.nav_buttons: dict[str, SidebarButton] = {}
        for key, tkey, icon in self.NAV_ITEMS:
            btn = SidebarButton(self.sidebar, text=t(tkey), icon=icon,
                               command=lambda k=key: self.show_page(k))
            btn.pack(fill="x", padx=12, pady=(0, 4))
            self.nav_buttons[key] = btn

        self.version_label = ctk.CTkLabel(
            self.sidebar, text="v2.0 — by duohnson",
            font=ctk.CTkFont(size=11),
            text_color=COLORS["text_muted"])
        self.version_label.pack(side="bottom", anchor="w", padx=18, pady=(0, 20))

    def _refresh_sidebar(self):
        for key, tkey, icon in self.NAV_ITEMS:
            btn = self.nav_buttons.get(key)
            if btn:
                label = f"  {icon}  {t(tkey)}" if icon else t(tkey)
                btn.configure(text=label)

    def _build_content(self):
        self.content = ctk.CTkFrame(self, fg_color=COLORS["bg_dark"], corner_radius=0)
        self.content.grid(row=0, column=1, sticky="nsew")
        self.content.grid_rowconfigure(0, weight=1)
        self.content.grid_columnconfigure(0, weight=1)
        self.current_page = None

    def show_page(self, key):
        self._current_key = key
        for k, btn in self.nav_buttons.items():
            btn.set_active(k == key)

        if self.current_page:
            self.current_page.destroy()

        pages = {
            "dashboard": DashboardPage,
            "add":       AddPage,
            "library":   LibraryPage,
            "settings":  SettingsPage,
        }
        cls = pages.get(key, DashboardPage)
        self.current_page = cls(self.content, self)
        self.current_page.grid(row=0, column=0, sticky="nsew")


if __name__ == "__main__":
    app = IconManagerApp()
    app.mainloop()
