

import win32com.client
import pythoncom
import Log
import os
import shutil
import sys

# NETTOYAGE IMMÉDIAT du cache gencache corrompu dès l'importation du module
# Cela évite que le cache corrompu soit utilisé avant même qu'on puisse le nettoyer
def _cleanup_on_import():
    """Nettoie le cache gencache Word dès l'importation du module"""
    try:
        # ÉTAPE 1: Supprimer TOUS les modules win32com.gen_py contenant 000209
        # Le format du nom de module est: win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x7
        # Il faut être très agressif car le module peut être chargé de différentes manières
        modules_to_remove = []
        for module_name in list(sys.modules.keys()):
            if 'win32com.gen_py' in module_name:
                module_lower = module_name.lower()
                # Supprimer TOUS les modules contenant les CLSIDs Word
                # Pattern exact de l'erreur: 00020905-0000-0000-C000-000000000046x0x8x7
                if ('000209' in module_name or 
                    '00020905' in module_name or 
                    '000209ff' in module_lower or
                    '00020900' in module_name or
                    '00020905-0000-0000-c000' in module_lower or
                    '00020905-0000-0000-c000-000000000046' in module_lower or
                    'word' in module_lower):
                    modules_to_remove.append(module_name)
        
        # Supprimer les modules (essayer plusieurs fois pour être sûr)
        for _ in range(3):  # Répéter 3 fois pour être sûr
            for module_name in modules_to_remove:
                try:
                    if module_name in sys.modules:
                        del sys.modules[module_name]
                except:
                    pass
        
        # ÉTAPE 1.5: Supprimer aussi TOUS les modules win32com.gen_py (même sans 000209)
        # Car ils peuvent être interdépendants
        for module_name in list(sys.modules.keys()):
            if 'win32com.gen_py' in module_name and '000209' in module_name:
                try:
                    del sys.modules[module_name]
                except:
                    pass
        
        # ÉTAPE 2: Supprimer les fichiers du cache gencache
        try:
            import win32com.client.gencache
            cache_path = win32com.client.gencache.GetGeneratePath()
            if os.path.exists(cache_path):
                for item in os.listdir(cache_path):
                    item_lower = item.lower()
                    # Supprimer tous les dossiers contenant les CLSIDs de Word
                    if ('000209' in item_lower or 
                        '00020905' in item_lower or 
                        '000209ff' in item_lower or
                        'word' in item_lower):
                        item_path = os.path.join(cache_path, item)
                        try:
                            if os.path.isdir(item_path):
                                shutil.rmtree(item_path)
                            elif os.path.isfile(item_path):
                                os.remove(item_path)
                        except:
                            pass
        except:
            pass
        
        # ÉTAPE 3: Désactiver MakePy ET EnsureDispatch de manière permanente pour cette session
        try:
            import win32com.client.gencache
            
            # Remplacer EnsureModule pour qu'il retourne toujours None
            # Cela force l'utilisation du dispatch dynamique pur
            if hasattr(win32com.client.gencache, 'EnsureModule'):
                original_EnsureModule = win32com.client.gencache.EnsureModule
                def disabled_EnsureModule(*args, **kwargs):
                    return None  # Force le dispatch dynamique
                win32com.client.gencache.EnsureModule = disabled_EnsureModule
                # Stocker la référence originale
                win32com.client.gencache._original_EnsureModule = original_EnsureModule
            
            # Remplacer aussi EnsureDispatch si elle existe
            if hasattr(win32com.client.gencache, 'EnsureDispatch'):
                original_EnsureDispatch = win32com.client.gencache.EnsureDispatch
                def disabled_EnsureDispatch(*args, **kwargs):
                    # Utiliser DispatchEx au lieu de EnsureDispatch
                    import win32com.client
                    return win32com.client.DispatchEx(args[0]) if args else None
                win32com.client.gencache.EnsureDispatch = disabled_EnsureDispatch
                win32com.client.gencache._original_EnsureDispatch = original_EnsureDispatch
        except:
            pass
        
        # ÉTAPE 4: Forcer le garbage collector
        import gc
        gc.collect()
    except:
        pass

# Exécuter le nettoyage immédiatement
_cleanup_on_import()

__wordInstance = None

def _ClearWordGenCache():
    """
    Nettoie complètement le cache gencache pour Word afin d'éviter les erreurs 
    CLSIDToClassMap et MinorVersion après une mise à jour de Windows/Office.
    Supprime les modules du cache ET du sys.modules s'ils sont déjà importés.
    """
    cleaned_modules = False
    cleaned_cache = False
    
    try:
        # ÉTAPE 1: Supprimer TOUS les modules win32com.gen_py de sys.modules
        # (pas seulement ceux de Word, car ils peuvent être interdépendants)
        modules_to_remove = []
        for module_name in list(sys.modules.keys()):
            if 'win32com.gen_py' in module_name:
                # Supprimer tous les modules gencache qui pourraient être corrompus
                if '000209' in module_name or '00020905' in module_name or '000209ff' in module_name.lower() or 'word' in module_name.lower():
                    modules_to_remove.append(module_name)
        
        for module_name in modules_to_remove:
            try:
                del sys.modules[module_name]
                Log.Verbose(f"[WORDCONTROLLER] Module gencache supprimé de sys.modules: {module_name}")
                cleaned_modules = True
            except:
                pass
        
        # ÉTAPE 1.5: Forcer le garbage collector pour libérer les références
        import gc
        gc.collect()
        
        # ÉTAPE 2: Supprimer les fichiers/dossiers du cache gencache
        import win32com.client.gencache
        cache_path = win32com.client.gencache.GetGeneratePath()
        
        if not os.path.exists(cache_path):
            return cleaned_modules or cleaned_cache
        
        # Chercher les dossiers liés à Word dans le cache
        # CLSID de Word.Application: {000209FF-0000-0000-C000-000000000046}
        # Différentes versions de Word peuvent avoir différents CLSIDs
        word_clsids = ["000209FF", "00020905", "00020900", "000209"]  # Différentes versions de Word
        word_keywords = ["word", "000209"]  # Mots-clés pour détecter Word
        
        # Aussi chercher par variantes du nom de fichier généré
        # Format typique: 00020905-0000-0000-C000-000000000046x0x8x7
        
        items_to_remove = []
        for item in os.listdir(cache_path):
            item_lower = item.lower()
            # Chercher les dossiers correspondant aux CLSIDs de Word ou contenant "word"
            should_remove = False
            for clsid in word_clsids:
                if clsid.lower() in item_lower:
                    should_remove = True
                    break
            # Vérifier aussi les mots-clés
            if not should_remove:
                for keyword in word_keywords:
                    if keyword in item_lower:
                        should_remove = True
                        break
            
            if should_remove:
                items_to_remove.append(item)
        
        # Supprimer les items trouvés
        for item in items_to_remove:
            item_path = os.path.join(cache_path, item)
            try:
                if os.path.isdir(item_path):
                    shutil.rmtree(item_path)
                    Log.Verbose(f"[WORDCONTROLLER] Dossier cache Word supprimé: {item}")
                    cleaned_cache = True
                elif os.path.isfile(item_path):
                    os.remove(item_path)
                    Log.Verbose(f"[WORDCONTROLLER] Fichier cache Word supprimé: {item}")
                    cleaned_cache = True
            except PermissionError:
                Log.Warning(f"[WORDCONTROLLER] Impossible de supprimer {item} (fichier verrouillé)")
            except Exception as e:
                Log.Verbose(f"[WORDCONTROLLER] Impossible de supprimer {item}: {str(e)}")
        
        if cleaned_modules or cleaned_cache:
            Log.Message("[WORDCONTROLLER] Cache gencache Word nettoyé (modules et fichiers)")
        
        return cleaned_modules or cleaned_cache
    except Exception as e:
        Log.Verbose(f"[WORDCONTROLLER] Erreur lors du nettoyage du cache: {str(e)}")
        return False

def _CheckAndFixGenCache():
    """
    Vérifie si le cache gencache pour Word est corrompu et le nettoie préventivement.
    Cette fonction détecte les problèmes avant qu'ils ne causent des erreurs.
    """
    try:
        import win32com.client.gencache
        cache_path = win32com.client.gencache.GetGeneratePath()
        
        if not os.path.exists(cache_path):
            return False
        
        # Chercher les modules Word corrompus en vérifiant l'existence du fichier __init__.py
        word_clsids = ["000209FF", "00020905", "00020900"]
        
        for item in os.listdir(cache_path):
            item_lower = item.lower()
            for clsid in word_clsids:
                if clsid.lower() in item_lower:
                    item_path = os.path.join(cache_path, item)
                    if os.path.isdir(item_path):
                        # Vérifier si le module est corrompu (pas de __init__.py ou fichier manquant)
                        init_file = os.path.join(item_path, "__init__.py")
                        if not os.path.exists(init_file):
                            # Cache corrompu, le nettoyer
                            _ClearWordGenCache()
                            return True
                        
                        # Vérifier si CLSIDToClassMap existe dans les fichiers .py
                        for py_file in os.listdir(item_path):
                            if py_file.endswith(".py"):
                                py_path = os.path.join(item_path, py_file)
                                try:
                                    with open(py_path, 'r', encoding='utf-8') as f:
                                        content = f.read()
                                        # Si le fichier référence CLSIDToClassMap mais n'a pas la définition
                                        if 'CLSIDToClassMap' in content and 'CLSIDToClassMap =' not in content:
                                            # Cache probablement corrompu
                                            _ClearWordGenCache()
                                            return True
                                except:
                                    pass
        
        return False
    except Exception as e:
        Log.Verbose(f"[WORDCONTROLLER] Erreur lors de la vérification du cache: {str(e)}")
        return False

def OpenWord() :
    global __wordInstance
    if __wordInstance != None :
        Log.Verbose("[WORDCONTROLLER] Word is already open")
        return True
    try :
        # Nettoyer IMMÉDIATEMENT le cache corrompu AVANT toute tentative
        # Cela évite les erreurs CLSIDToClassMap et MinorVersion après une mise à jour
        # Il faut nettoyer sys.modules aussi car le cache peut être déjà importé
        Log.Verbose("[WORDCONTROLLER] Nettoyage complet du cache gencache (fichiers + modules importés)...")
        _ClearWordGenCache()
        
        # Force la suppression de tout module gencache Word qui pourrait être chargé
        # en recherchant dans sys.modules
        import gc
        gc.collect()  # Force le garbage collector pour libérer les références
        
        # Initialiser COM pour le thread actuel
        try:
            pythoncom.CoInitialize()
        except:
            # COM déjà initialisé, continuer
            pass
        
        # DÉSACTIVER COMPLÈTEMENT MakePy pour forcer le dispatch dynamique pur
        # Cela évite TOUS les problèmes de cache corrompu
        import win32com.client.gencache
        
        # S'assurer que EnsureModule est bien désactivé
        if hasattr(win32com.client.gencache, 'EnsureModule'):
            original_EnsureModule = win32com.client.gencache.EnsureModule
            def disable_makepy(*args, **kwargs):
                return None  # Force le dispatch dynamique
            win32com.client.gencache.EnsureModule = disable_makepy
        else:
            original_EnsureModule = None
        
        try:
            # Approche 1: Utiliser pythoncom directement (bypass complet du cache gencache)
            try:
                from pythoncom import CoCreateInstance, CLSCTX_LOCAL_SERVER
                import pythoncom as pycom
                
                # CLSID de Word.Application: {000209FF-0000-0000-C000-000000000046}
                word_clsid = pycom.MakeIID("{000209FF-0000-0000-C000-000000000046}")
                word_iid = pycom.IID_IDispatch
                
                word_obj = CoCreateInstance(word_clsid, None, CLSCTX_LOCAL_SERVER, word_iid)
                __wordInstance = win32com.client.Dispatch(word_obj)
                Log.Verbose("[WORDCONTROLLER] Word ouvert avec CoCreateInstance direct (bypass cache)")
            except Exception as direct_err:
                # Approche 2: Utiliser DispatchEx avec MakePy désactivé
                try:
                    __wordInstance = win32com.client.DispatchEx("Word.Application")
                    Log.Verbose("[WORDCONTROLLER] Word ouvert avec DispatchEx (MakePy désactivé)")
                except Exception as dispex_err:
                    # Si DispatchEx échoue, essayer Dispatch mais d'abord supprimer complètement le cache
                    Log.Warning(f"[WORDCONTROLLER] DispatchEx échoué: {str(dispex_err)}")
                    Log.Verbose("[WORDCONTROLLER] Nettoyage complet du cache et nouvelle tentative...")
                    
                    # Supprimer TOUT le cache gencache, pas seulement Word
                    try:
                        import win32com.client.gencache
                        cache_path = win32com.client.gencache.GetGeneratePath()
                        if os.path.exists(cache_path):
                            # Supprimer tous les dossiers commençant par les CLSIDs de Word
                            for item in os.listdir(cache_path):
                                item_path = os.path.join(cache_path, item)
                                if os.path.isdir(item_path) and ("000209" in item.lower() or "word" in item.lower()):
                                    try:
                                        shutil.rmtree(item_path)
                                        Log.Verbose(f"[WORDCONTROLLER] Cache supprimé: {item}")
                                    except:
                                        pass
                    except:
                        pass
                    
                    try:
                        # Essayer Dispatch qui peut fonctionner même sans cache
                        __wordInstance = win32com.client.Dispatch("Word.Application")
                        Log.Verbose("[WORDCONTROLLER] Word ouvert avec Dispatch (après nettoyage complet)")
                    except Exception as dispatch_err:
                        # Dernière tentative avec DispatchEx
                        Log.Warning(f"[WORDCONTROLLER] Dispatch échoué: {str(dispatch_err)}")
                        try:
                            __wordInstance = win32com.client.DispatchEx("Word.Application")
                            Log.Verbose("[WORDCONTROLLER] Word ouvert avec DispatchEx (dernière tentative)")
                        except Exception as final_err:
                            Log.Error(f"[WORDCONTROLLER] Impossible de créer l'instance Word: {str(final_err)}")
                            import traceback
                            Log.Error(traceback.format_exc())
                            return False
        finally:
            # Restaurer la fonction EnsureModule originale si elle existe
            if original_EnsureModule is not None:
                try:
                    win32com.client.gencache.EnsureModule = original_EnsureModule
                except:
                    pass
        
        # Définir les propriétés avec gestion d'erreurs
        try:
            __wordInstance.DisplayAlerts = False
        except:
            Log.Warning("[WORDCONTROLLER] Impossible de désactiver les alertes Word")
            
        try:
            __wordInstance.Visible = False
        except:
            Log.Warning("[WORDCONTROLLER] Impossible de cacher la fenêtre Word")
            # Continuer même si on ne peut pas cacher Word
        
        Log.Message("[WORDCONTROLLER] Instance Word créée avec succès")
        return True
    except Exception as e :
        Log.Error("[WORDCONTROLLER] : " + str(e))
        import traceback
        Log.Error(traceback.format_exc())
        return False


def CloseWord() :
    global __wordInstance
    if __wordInstance == None :
        Log.Verbose("[WORDCONTROLLER] Word is not open, cannot close it")
        return True
    try :
        # Fermer tous les documents ouverts
        try:
            for doc in __wordInstance.Documents:
                try:
                    doc.Close(SaveChanges=False)
                except:
                    pass
        except:
            pass
        
        # Fermer Word
        try:
            __wordInstance.Quit(SaveChanges=False)
        except:
            pass
            
    except Exception as e :
        Log.Error("[WORDCONTROLLER] Erreur lors de la fermeture: " + str(e))
        
    finally:
        __wordInstance = None
        # Libérer COM
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        return True

def PrintDocument(path : str) :
    if __wordInstance == None :
        Log.Error("[WORDCONTROLLER] : Word n'est pas ouvert")
        return False
    
    try :
        __wordInstance.PrintOut(path)
        return True
    
    except Exception as e :
        Log.Error("[WORDCONTROLLER] : " + str(e))
        return False

def ConvertToPdfFromPath(path : str) :

    if __wordInstance == None :
        Log.Error("[WORDCONTROLLER] : Word n'est pas ouvert")
        return False

    try :
        path = path.replace("/", "\\")
        document = __wordInstance.Documents.Open(path)
        document.ExportAsFixedFormat(OutputFileName = path.removesuffix(".docx") + ".pdf", ExportFormat=17)
        document.Close(True)
        return True
    except Exception as e :
        Log.Error("[WORDCONTROLLER] : " + str(e))
        return False