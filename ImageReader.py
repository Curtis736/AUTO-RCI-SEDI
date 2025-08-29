import os

import PIL
import PIL.Image
import PIL.ImageChops
from pdf2image import convert_from_path
import ExcelController

import Log
import Settings


CACHE_PATH = "tmpimg"


__cachedImages = []

def Init() :

    os.makedirs(CACHE_PATH, exist_ok=True)

# returns the full path to the temp image folder
def GetAbspathTempFolder() :
    return os.path.join(os.path.abspath(CACHE_PATH))

def Clear() :

    for file in os.scandir(CACHE_PATH) :

        if file.is_file() :
            Log.Verbose(f"ImageReader removed temp file {file.path}")
            os.remove(file.path)

def CacheFromData(imageData : PIL.Image, name : str) :

    path = os.path.join(CACHE_PATH, name) + ".png"
    if os.path.exists(path) :
        path += "_1"
        count = 1
        while os.path.exists(path) and count < 100 :
            count += 1
            path = path[:-2] + str(count)
        
    path = os.path.abspath(path)
    imageData.save(path)
    __cachedImages.append(name)
    Log.Verbose(f"Cached image {path}")
    return path

def GetAbsPathCached(name : str) :

    if not name in __cachedImages : return None
    return os.path.join(Settings.CWD, CACHE_PATH, name) + ".png"


def LoadPdfFile(pathToFile : str) :
    """
    Assumes the path to the file is already verified. It will convert only the first page of the pdf as an image.
    The resulting image is put in the temp folder.
    """

    images = []

    try :
        images = convert_from_path(pathToFile, poppler_path="X:/Production/4_Public/THIBAUD/poppler-24.02.0/Library/bin", use_cropbox=True)
    except Exception as e:
        Log.Error("failed open pdf as images")
        Log.Error(str(e))
        return None
    
    if len(images) == 0 :
        Log.Error("error, no pages in pdf")
        return None
    
    # on prend seulement la première page
    image = images[0]

    fileName = os.path.basename(pathToFile)
    fileName.removesuffix(".pdf")

    try :
        image = TrimWhiteBackground(image)
    except Exception as e :
        Log.Error("impossible de rogner l'image " + pathToFile)

    try :
        return CacheFromData(image, "pdf_" + os.path.basename(pathToFile))
    except Exception as e :
        Log.Error(str(e))

    return None


def TrimWhiteBackground(image):
    # code taken from
    # https://stackoverflow.com/questions/10615901/trim-whitespace-using-pil
    background = PIL.Image.new(image.mode, image.size, image.getpixel((0, image.size[1] - 1)))
    diff = PIL.ImageChops.difference(image, background)
    diff = PIL.ImageChops.add(diff, diff, 2.0, -100)
    bbox = diff.getbbox()
    if bbox:
        return image.crop(bbox)
    else :
        return image
    

"""
Convert the given png / pdf / xlsx into a valid png image path.
Need an absolute path.
"""
def GetImagePath(imageAbsPath : str) :


    if not os.path.isfile(imageAbsPath) :
        Log.Error(f"Le fichier image {imageAbsPath} n'existe pas")
        return None
    
    if imageAbsPath.endswith(".png") :
        return imageAbsPath
    
    if imageAbsPath.endswith(".jpg") :
        return imageAbsPath
    
    if imageAbsPath.endswith(".pdf") :
        return LoadPdfFile(imageAbsPath)
    