"""
Installation et résolution automatique de Poppler (pdf2image).
Poppler est installé dans tools/poppler/ à la première utilisation — aucune config manuelle requise.
"""

import os
import shutil
import zipfile
import urllib.request

import Log
import Settings


POPPLER_VERSION = "24.08.0-0"
POPPLER_URL = (
    f"https://github.com/oschwartz10612/poppler-windows/releases/download/"
    f"v{POPPLER_VERSION}/Release-{POPPLER_VERSION}.zip"
)
BUNDLED_POPPLER_ROOT = os.path.join(Settings.CWD, "tools", "poppler")
INSTALL_LOCK = os.path.join(Settings.CWD, "tools", ".poppler_install.lock")


def is_poppler_bin(path: str) -> bool:
    if not path:
        return False
    return os.path.isfile(os.path.join(path, "pdftoppm.exe"))


def resolve_poppler_bin(candidate: str):
    """
    Accepte un dossier bin Poppler ou la racine d'une archive (Library/bin, bin).
    Retourne le dossier contenant pdftoppm.exe, ou None.
    """
    if not candidate:
        return None

    candidate = os.path.normpath(str(candidate).strip())
    if not os.path.exists(candidate):
        return None

    if os.path.isfile(candidate):
        if os.path.basename(candidate).lower() == "pdftoppm.exe":
            return os.path.abspath(os.path.dirname(candidate))
        return None

    if is_poppler_bin(candidate):
        return os.path.abspath(candidate)

    for sub in ("Library\\bin", "library\\bin", "bin"):
        bin_path = os.path.join(candidate, sub)
        if is_poppler_bin(bin_path):
            return os.path.abspath(bin_path)

    return None


def _iter_directory_poppler_roots(directory: str):
    if not directory or not os.path.isdir(directory):
        return
    try:
        for name in os.listdir(directory):
            if "poppler" not in name.lower():
                continue
            yield os.path.join(directory, name)
    except OSError:
        return


def _iter_shared_poppler_candidates():
    """Cherche Poppler sur le réseau en remontant depuis le dossier du projet."""
    current = Settings.CWD
    seen = set()

    for _ in range(10):
        for root in _iter_directory_poppler_roots(current):
            key = os.path.normcase(root)
            if key not in seen:
                seen.add(key)
                yield root

        for folder_name in ("Outils", "TOOLS", "tools", "THIBAUD", "4_Public", "Production"):
            sibling = os.path.join(current, folder_name)
            for root in _iter_directory_poppler_roots(sibling):
                key = os.path.normcase(root)
                if key not in seen:
                    seen.add(key)
                    yield root

        parent = os.path.dirname(current)
        if parent == current:
            break
        current = parent


def _download_and_install_poppler():
    tools_dir = os.path.join(Settings.CWD, "tools")
    os.makedirs(tools_dir, exist_ok=True)

    if os.path.isfile(INSTALL_LOCK):
        Log.Message("[POPPLER] Installation déjà en cours sur un autre processus, attente...")
        for _ in range(30):
            if resolve_poppler_bin(BUNDLED_POPPLER_ROOT):
                return resolve_poppler_bin(BUNDLED_POPPLER_ROOT)
            if not os.path.isfile(INSTALL_LOCK):
                break
            import time
            time.sleep(2)
        return resolve_poppler_bin(BUNDLED_POPPLER_ROOT)

    try:
        with open(INSTALL_LOCK, "w", encoding="utf-8") as lock_file:
            lock_file.write("installing")

        zip_path = os.path.join(tools_dir, f"Release-{POPPLER_VERSION}.zip")
        extract_temp = os.path.join(tools_dir, "_poppler_download")

        Log.Message(
            f"[POPPLER] Téléchargement automatique de Poppler {POPPLER_VERSION} "
            f"(première utilisation, connexion Internet requise)..."
        )

        urllib.request.urlretrieve(POPPLER_URL, zip_path)

        if os.path.isdir(extract_temp):
            shutil.rmtree(extract_temp, ignore_errors=True)
        os.makedirs(extract_temp, exist_ok=True)

        with zipfile.ZipFile(zip_path, "r") as archive:
            archive.extractall(extract_temp)

        extracted_root = None
        for name in os.listdir(extract_temp):
            full_path = os.path.join(extract_temp, name)
            if os.path.isdir(full_path) and "poppler" in name.lower():
                extracted_root = full_path
                break

        if not extracted_root:
            Log.Error("[POPPLER] Archive téléchargée invalide (dossier poppler introuvable)")
            return None

        if os.path.isdir(BUNDLED_POPPLER_ROOT):
            shutil.rmtree(BUNDLED_POPPLER_ROOT, ignore_errors=True)

        shutil.move(extracted_root, BUNDLED_POPPLER_ROOT)
        shutil.rmtree(extract_temp, ignore_errors=True)

        try:
            os.remove(zip_path)
        except OSError:
            pass

        resolved = resolve_poppler_bin(BUNDLED_POPPLER_ROOT)
        if resolved:
            Log.Message(f"[POPPLER] Installation automatique terminée : {resolved}")
        else:
            Log.Error("[POPPLER] Installation terminée mais pdftoppm.exe introuvable")
        return resolved

    except Exception as e:
        Log.Error(f"[POPPLER] Échec du téléchargement/installation automatique : {e}")
        return None
    finally:
        try:
            os.remove(INSTALL_LOCK)
        except OSError:
            pass


def ensure_poppler_installed():
    """Installe Poppler dans tools/poppler si absent. Retourne le dossier bin ou None."""
    resolved = resolve_poppler_bin(BUNDLED_POPPLER_ROOT)
    if resolved:
        return resolved
    return _download_and_install_poppler()


def get_poppler_bin_path(force_refresh: bool = False, allow_download: bool = True):
    """
    Résout le dossier bin Poppler sans configuration manuelle.
    Ordre : tools/poppler local → copies réseau proches → PATH → auto-téléchargement.
    """
    if not hasattr(get_poppler_bin_path, "_cache"):
        get_poppler_bin_path._cache = None
        get_poppler_bin_path._resolved = False

    if get_poppler_bin_path._resolved and not force_refresh:
        return get_poppler_bin_path._cache

    sources = []

    sources.append(("installation projet", BUNDLED_POPPLER_ROOT))
    sources.append(("dossier tools", os.path.join(Settings.CWD, "tools")))

    for candidate in _iter_shared_poppler_candidates():
        sources.append(("réseau partagé", candidate))

    which = shutil.which("pdftoppm")
    if which:
        sources.append(("PATH système", os.path.dirname(which)))

    for env_key in ("POPPLER_PATH", "POPPLER_HOME"):
        env_val = os.environ.get(env_key, "").strip()
        if env_val:
            sources.append((f"variable {env_key}", env_val))

    try:
        configured = Settings.GetConfigValueString("paths", "poppler_path")
        if configured:
            sources.append(("config paths.json", configured))
    except Exception:
        pass

    legacy_root = r"X:\Production\4_Public\THIBAUD\poppler-24.02.0"
    if os.path.exists(legacy_root):
        sources.append(("chemin réseau historique", legacy_root))

    seen_candidates = set()
    for source_name, candidate in sources:
        key = os.path.normcase(str(candidate))
        if key in seen_candidates:
            continue
        seen_candidates.add(key)

        resolved = resolve_poppler_bin(candidate)
        if resolved:
            get_poppler_bin_path._cache = resolved
            get_poppler_bin_path._resolved = True
            Log.Message(f"[POPPLER] Utilisation de {resolved} (source: {source_name})")
            return resolved

    if allow_download:
        installed = ensure_poppler_installed()
        if installed:
            get_poppler_bin_path._cache = installed
            get_poppler_bin_path._resolved = True
            return installed

    get_poppler_bin_path._cache = None
    get_poppler_bin_path._resolved = True
    Log.Warning(
        "[POPPLER] Poppler introuvable et installation automatique impossible. "
        "Vérifiez la connexion Internet ou exécutez install_modules.bat une fois."
    )
    return None
