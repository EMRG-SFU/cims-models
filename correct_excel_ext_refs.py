#!/bin/python
"""
Excel external link correction.
"""

import openpyxl as op
import os, os.path
import sys
import re
import platform
import argparse
import logging
from datetime import datetime



#############################################################
#############################################################
#############################################################
# Some initial setup. 
#

class UnexpectedFwdSlashes(Exception):
    """
    How to fail from parsing a windows-style path.
    """
    pass

logging.basicConfig(filename="paths_for_token.log", 
                    level=logging.INFO,
                    filemode='w'  # This wipes and restarts the log on each load of this.
                    )
winlogger = logging.getLogger("winlogger")
normlogger = logging.getLogger("normlogger")


#############################################################
#############################################################
#############################################################
# Helper Functions
#
# Splitting operating system file paths into lists of path tokens. 
#
# Only use these on python/os file paths. Don't use on external link 
# path strings (as these don't depend on what flavour os you are using).
#
def iterPath_withRoot(p):
    res = os.path.split(p)
    while res[1] != '':
        yield res[1]
        res = os.path.split(res[0])

    yield res[0]

def iterPath(p):
    res = os.path.split(p)
    while res[1] != '':
        yield res[1]
        res = os.path.split(res[0])

def tokenizePath_withRoot(p):
    return( [z for z in reversed([a for a in iterPath_withRoot(p)])] )

def tokenizePath(p):
    return( [z for z in reversed([a for a in iterPath(p)])] )


#############################################################
#############################################################
#############################################################
# Helper Functions
#
# Tokenizing functions for breaking up paths stored in excel
# external refs.
#
def tokenizeWindows( p, pathToCimsModels ):
    """
    """
    #print(f"calling tokenizeWindows for {p}")
    p2 = re.sub("^file:///", "", p)

    if re.search("/", p2):
        raise UnexpectedFwdSlashes("Path: " + p2)

    p3 = re.split("\\\\", p2)

    if not len(p3) >= 1:
        raise EmptyPathError(f"Path seems empty when parsed: {p}")


    # In windows-style paths, absolute paths always begin with a drive letter followed by
    # a colon. If this is found in the parsed path, report we've found an absolute path.
    if re.match("[a-zA-Z]:", p3[0]):
        isAbsPath = True
    else:
        isAbsPath = False

    winlogger.info(f",{isAbsPath},{p3[0]},{str(p3).replace(',','|')},{p}")

    if isAbsPath:
        try:
            startInd = p3.index('cims-models')
        except:
            raise CIMSModelsNotFound('')
        # Get rid of all the file path items above 
        # and including `cims-models`.
        p4 = p3[(startInd+1):]
    else:
        p4 = p3

    return(p4)

def tokenizeNormal( p, pathToCimsModels ):
    """
    """
    p2 = re.split("/", p)

    if not len(p2) >= 1:
        raise EmptyPathError(f"Path seems empty when parsed: {p}")


    # In normal-style paths, absolut paths always start with the slash path delimiter character.
    # This makes the first element in the parsed path list be ''. If this is the case, report that
    # we've found an absolute path.
    if p2[0] == '':
        isAbsPath = True
    else:
        isAbsPath = False

    normlogger.info(f",{isAbsPath},{p2[0]},{str(p2).replace(',','|')},{p}")

    if isAbsPath:
        try:
            startInd = p2.index('cims-models')
        except:
            raise CIMSModelsNotFound('')
        # Get rid of all the file path items above
        # and including`cims-models`.
        p3 = p2[(startInd+1):]
    else:
        p3 = p2

    return(p3)


#############################################################
#############################################################
#############################################################
# Search for the file referenced by the Excel external ref on
# the filesystem, in order to figure out what the actual path
# on the filesystem is in terms of case-sensitivity. We're 
# having problems relying on any sort of case-sensitivity 
# rules vis-a-vis MacOS,excel, windows, linux, etc.
#
def getSystemPath(cmRoot, filePath, *args, **kwargs):
    """
    `cmRoot`: absolute path to the `cims-models` root on the system
    `filePath`: the name of what the file should be according to the excel external ref, rooted
                at `cims-models` (assuming it includes `cims-models`. Otherwise it should have
                already failed).

    We'll use `os.walk()` to step through the entire hierarchy, and match the path stored
    in the excel external ref, to the path on the system, with both of them forced to lowercase.
    When there's a match, we then take the non-forced system path, and then we use that for what the 
    external ref path SHOULD be.
    """
    try:
        exPathTokens = tokenizeWindows(filePath, cmRoot)
    except UnexpectedFwdSlashes:
        exPathTokens = tokenizeNormal(filePath, cmRoot)

    fullExtRefPath = os.path.abspath(os.path.join(*exPathTokens)).replace("%20", " ")

    if os.path.exists(fullExtRefPath):

        return((True, fullExtRefPath))

    else:

        pathCmpList = []
        for (dirpath, dirnames, filenames) in os.walk(cmRoot):
            for f in filenames:
                fullSysPath = os.path.join(dirpath, f)
                #print(f"Paths same -- {fullExtRefPath.lower() == fullSysPath.lower()}")
                #pathCmpList.append( (fullExtRefPath.lower()==fullSysPath.lower(), fullSysPath, fullExtRefPath) )
                if fullExtRefPath.lower() == fullSysPath.lower():
                    pathCmpList.append(fullSysPath)
        if len(pathCmpList) == 0:
            return((False, None))
        elif len(pathCmpList) == 1:
            return((False, pathCmpList[0]))
        else:
            raise RuntimeError(f"Too many file matches -- {pathCmpList}")
            



class CIMSModelsNotFound(Exception):
    """
    How to fail when `cims-models` is not found within the external link
    path we're trying to correct.
    """
    pass

def repathWindowsRelative( p, pathToCimsModels, num_backups ):
    #print(f"calling repathWindowsRelative for {p}")
    p2 = re.sub("^file:///", "", p)

    if re.search("/", p2):
        raise UnexpectedFwdSlashes("Path: " + p2)

    p3 = re.split("\\\\", p2)

    # Get rid of all the file path items above
    # and including`cims-models`.
    try:
        startInd = p3.index('cims-models')
    except:
        raise CIMSModelsNotFound('')

    p4 = p3[(startInd+1):]

    #p5 = os.path.join( *p4 )
    p5 = "/".join([".." for i in range(0, num_backups)]) + "/" + "/".join(p4)

    return(p5)



def repathNormalRelative( p, pathToCimsModels, num_backups ):
    #print(f"calling repathNormalRelative for {p}")
    p2 = re.split("/", p)

    try:
        startInd = p2.index('cims-models')
    except:
        raise CIMSModelsNotFound('')

    # Get rid of all the file path items above
    # and including`cims-models`.
    p3 = p2[(startInd+1):]

    #p4 = os.path.join( *p3 )
    p4 = "/".join([".." for i in range(0, num_backups)]) + "/" + "/".join(p3)

    return(p4)






def _debug__check_ext_links( xl_path, pathToCimsModels, pathLookupTable, corrExt=None):
    """
    A debugging function for easier access to the internals of the
    path comparisons this module performs.
    """

    print(f"\n\n******* Working on file: {xl_path}\n")
    wb = op.open(xl_path)


    # Here we figure out how many directories down the xl file we're trying to modify is. For the relative external
    # link paths, we'll have to bring them up to the level of the `cims-models` folder first, using `../`, and the number
    # of these `up` thingies we need is the value of `num_backups` below.
    path_tokens = tokenizePath(os.path.abspath(xl_path))
    cimsModels_index = path_tokens.index('cims-models')
    path_trunc = path_tokens[(cimsModels_index+1):]
    num_backups = len(path_trunc) - 1

    ret = []
    for index in range(0, len(wb._external_links)):
        oldPath = wb._external_links[index].file_link.target
        ret.append((oldPath, getSystemPath(pathToCimsModels, oldPath)))

    return(ret)


#############################################################
#############################################################
#############################################################
# Recursively find xlsx files, rooted at `dirPath`
#
def find_excel_files( dirPath ):
    """
    Find every excel file recursively, starting at `dirPath`, and return in a dictionary
    structure, where the keys are the full filepath leading to the excel file, and the items
    in the dict are lists of `file_link.target` found in each excel file.
    """
    #print(f"DirPath: {dirPath}")
    retList = []
    for (dirpath, dirnames, filenames) in os.walk(dirPath):
        #print(f"dirpath: {str(dirpath)}")
        for f in filenames:
            if re.match("^~", f):
                # These are temp files used by excel, when file is open.
                continue
            if os.path.splitext(f)[1] == '.xlsx':
                #print(f"Working on: {f}")
                #wb = op.open(os.path.join(dirpath,f))
                #retDict[os.path.join(dirpath, f)] = [a.file_link.target for a in wb._external_links]
                #wb.close()
                retList.append(os.path.join(dirpath,f))

    return(retList)


#############################################################
#############################################################
#############################################################
# A couple of useful loggers
#
def log_info( msg=None, xl_path_val=None, xl_path=None, xl_path_corrected=None, xl_filename=None, pathOk=None, fileFound=None ):
    with open("excel_external_reference_correction.csv", "a") as f:
        def noneToStr(m):
            return( "" if (m is None) else m )
        f.write(f"OK,{pathOk},{fileFound},{xl_filename},{xl_path_val},{xl_path},{xl_path_corrected},{noneToStr(msg)}\n")

def log_failure( msg=None, xl_path_val=None, xl_path=None, xl_filename=None, pathOk=None, fileFound=None):
    with open("excel_external_reference_correction.csv", "a") as f:
        def noneToStr(m):
            return( "" if (m is None) else m )
        f.write(f"FAIL,{pathOk},{fileFound},{xl_filename},{xl_path_val},{xl_path},,{noneToStr(msg)}\n")

#############################################################
#############################################################
#############################################################
# The main function for link correction.
#

def correct_ext_links(xl_path, pathToCimsModels, pathLookupTable, corrExt=None, dryRun=False):

    wb = op.open(xl_path)

    path_tokens = tokenizePath(os.path.abspath(xl_path))
    cimsModels_index = path_tokens.index('cims-models')
    path_trunc = path_tokens[(cimsModels_index+1):]
    num_backups = len(path_trunc) - 1

    for index in range(0, len(wb._external_links)):
        oldPath = wb._external_links[index].file_link.target
        try:
            try:
                # Tokenizes RELATIVE to cims-models
                oldPathTokens = tokenizeWindows(oldPath, pathToCimsModels)
            except UnexpectedFwdSlashes:
                oldPathTokens = tokenizeNormal(oldPath, pathToCimsModels)

            # Ah... make sure these are dealt with correctly. The relative external links are from
            # the perspective of the excel files, but below this was being dropped into `os.path.abspath`
            # which is from the perspective of where this script is running, so it was backing out too
            # far and failing.
            fullExtRefPath = os.path.abspath(
                os.path.join(
                    os.path.split(xl_path)[0],
                    os.path.join(*oldPathTokens)
            ))
            #fullExtRefPath = os.path.abspath(os.path.join(*oldPathTokens)).replace("%20", " ")
        
        except CIMSModelsNotFound:
            log_failure(xl_path_val=oldPath, xl_path=None, xl_filename=xl_path, msg="cims-models not in path", pathOk=False, fileFound=False)
            continue
        except EmptyPathError:
            log_failure( msg="Path is empty.", xl_path_val=oldPath, xl_path=None, xl_filename=xl_path, pathOk=False, fileFound=False)
            continue
        
        if fullExtRefPath.lower() in pathLookupTable.keys():
            # The path exists on the system, disregarding case sensitivity.
            ### The `fullExtRefPath` is now an absolute path, same as the value in
            ### the pathLookupTable, indexed by the lower() key. This value is what we want,
            ### so use the existing path-tokenizing machinery to convert this path to a 
            ### proper external link relative path.
            pathToUse = pathLookupTable.get(fullExtRefPath.lower())
            try:
                newPath = repathWindowsRelative(pathToUse, pathToCimsModels, num_backups)
            except UnexpectedFwdSlashes:
                try:
                    newPath = repathNormalRelative(pathToUse, pathToCimsModels, num_backups)
                except CIMSModelsNotFound:
                    log_failure(xl_path_val=oldPath, xl_path=pathToUse, xl_filename=xl_path, msg="IMPOSSIBLE. cims-models not in path", pathOk=False, fileFound=True)
                    continue
                except EmptyPathError:
                    log_failure( msg="Path is empty.", xl_path_val=oldPath, xl_path=pathToUse, xl_filename=xl_path, pathOk=False, fileFound=True)
                    continue
            except CIMSModelsNotFound:
                log_failure(xl_path_val=oldPath, xl_path=pathToUse, xl_filename=xl_path, msg="cims-models not in path", pathOk=False, fileFound=True)
                continue
            except EmptyPathError:
                log_failure( msg="Path is empty.", xl_path_val=oldPath, xl_path=pathToUse, xl_filename=xl_path, pathOk=False, fileFound=True)
                continue

            log_info(xl_path_val=oldPath, xl_path=pathToUse, xl_path_corrected=newPath, xl_filename=xl_path, msg="Path OK and File Found")

        else:
            # The path doesn't exist. On the filesystem.
            # In this case we want to prepare it for use by Excel AS IF IT WERE here on the
            # filesystem. I guess in case it's just running late and shows up later? But anyway,
            # we log that this has occurred.
            pathToUse = fullExtRefPath
            try:
                newPath = repathWindowsRelative(pathToUse, pathToCimsModels, num_backups)
            except UnexpectedFwdSlashes:
                try:
                    newPath = repathNormalRelative(pathToUse, pathToCimsModels, num_backups)
                except CIMSModelsNotFound:
                    log_failure(msg="Filepath not found on system AND `cims-models` not in filepath.",xl_path_val=oldPath, xl_path=pathToUse, xl_filename=xl_path, pathOk=False, fileFound=False)
                    continue
                except EmptyPathError:
                    log_failure( msg="Path is empty.", xl_path_val=oldPath, xl_path=pathToUse, xl_filename=xl_path, pathOk=False, fileFound=False)
                    continue
            except CIMSModelsNotFound:
                log_failure(msg="Filepath not found on system AND `cims-models` not in filepath.", xl_path_val=oldPath, xl_path=pathToUse, xl_filename=xl_path, pathOk=False, fileFound=False)
                continue
            except EmptyPathError:
                log_failure( msg="Path is empty.", xl_path_val=oldPath, xl_path=pathToUse, xl_filename=xl_path, pathOk=False, fileFound=False)
                continue

            log_info(xl_path_val=oldPath, xl_path=pathToUse, xl_path_corrected=newPath, xl_filename=xl_path, msg="Path OK but File Not Found")

        if not dryRun:
            wb._external_links[index].file_link.target = newPath
            wb._external_links[index].file_link.id = 'rId1'

    if corrExt is None:
        xl_path_corrected = xl_path
    else:
        temp_splt = os.path.split(xl_path)
        temp_f_splt = os.path.splitext(temp_splt[1])
        xl_path_corrected = os.path.join(temp_splt[0], temp_f_splt[0] + corrExt + temp_f_splt[1])

    if not dryRun:
        wb.save(xl_path_corrected)





def buildLookupTable(cmRoot):
    """
    `cmRoot`: path to the `cims-models` folder on the system. If it's not an absolute path
              it'll be made into one.
    """
    if not os.path.isabs(cmRoot):
        cmRoot = os.path.abspath(cmRoot)
    fpDict = {}
    for (dirpath, dirnames, filenames) in os.walk(cmRoot):
        for f in filenames:
            if os.path.splitext(f)[1] == ".xlsx":
                fullSysPath = os.path.join(dirpath, f)
                fpDict[fullSysPath.lower()] = fullSysPath
            else:
                pass
    return(fpDict)


def correct_list(pathList,
                 pathToCimsModels, 
                 pathLookupTable, 
                 corrExt=None,
                 correctionFunc=correct_ext_links):
    """
    `pathList`: a list of relative paths to xlsx files, rooted at `cims-models`.
    `pathToCimsModels`: (unused, but just in case) Absolute path to `cims-models` dir.
    `corrExt`: suffix to append to corrected file name. If `None`, corrected file will
               be saved under the same name. (Leave `None` unless testing).
    """

    for p in pathList:
        if (not os.path.isfile(p)) or (os.path.splitext(p)[1] != ".xlsx"):
            raise RuntimeError(f"Supplied path is not xlsx file: {p}")
        correctionFunc(p, pathToCimsModels, pathLookupTable, corrExt)


if __name__ == "__main__":

    time_start = datetime.now()

    # Reset the info/failure log files
    with open("excel_external_reference_correction.csv", "w") as f:
        def isMsg(m):
            return( "" if (m is None) else m )
        #f.write(f"OK,{xl_filename},{xl_path},{xl_path_corrected},{isMsg(msg)}\n")
        f.write(f"CorrectionStatus,PathOK,FileFound,Filename,ExtLinkVal,ExtLinkPath,ExtLinkPath_corrected,Message/Error\n")

    whereAreWe = os.path.abspath('.')
    if os.path.split(whereAreWe)[1] != 'cims-models':
        raise RuntimeError("This script needs to be run within the `cims-models` directory")

    # This lookup table contains the paths with cases as they actually are on the system. The keys
    # in this table are these paths.lowered().
    pathLookupTable = buildLookupTable(whereAreWe)

    if len(sys.argv) == 1:
        correct_list(find_excel_files('.'), whereAreWe, pathLookupTable, correctionFunc=correct_ext_links)


    elif len(sys.argv) == 2:
        if os.path.isdir(sys.argv[1]):
            correct_list(find_excel_files(sys.argv[1]), whereAreWe, pathLookupTable, correctionFunc=correct_ext_links)


        elif os.path.isfile(sys.argv[1]):
            correct_ext_links(sys.argv[1], whereAreWe, pathLookupTable, "_corr")

        else:
            raise RuntimeError("Argument seems to be neither a directory nor a file.")
    else:
        # Here we assume that all positional arguments are pathnames pointing to specific 
        # excel files. There can be any number of these. We correct the external references in
        # each of them.
        correct_list(sys.argv[1:], whereAreWe, pathLookupTable, correctionFunc=correct_ext_links)

    time_end = datetime.now()
    print(f"Finished link correction in {time_end - time_start}.")

