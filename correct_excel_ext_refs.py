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


argParser = argparse.ArgumentParser(formatter_class=argparse.RawTextHelpFormatter)

subParsers = argParser.add_subparsers(help='subcommand help', dest="command_name")

parser_correct = subParsers.add_parser('CORRECT', help='Correct Excel External Ref Links\n\n', formatter_class=argparse.RawTextHelpFormatter)
parser_correct.add_argument("-d", "--dry-run", help="Run as normal, but don't actually change external reference links\n  (useful for logging what WOULD be done)\n\n", action="store_true", default=False)
parser_correct.add_argument("-nl", "--no-log", help="Suppress log file generation (excel_external_reference_correction.csv)\n\n", action="store_true", default=False)
parser_correct.add_argument("-s", "--strict", help="Different strictness levels\n  0: log all file path errors but continue\n  1: fail when an unfixable path encountered\n  2: fail when file referred to by path can't be found in `cims-models` folder hierarchy.\n\n", type=int, action="store", default=0, choices=[0,1,2])
parser_correct.add_argument("files", nargs="*", help="If a path to a single folder, process all files in hierarchy rooted there\n  (leave blank to start at `cims-models` top level).\n  Can also be a single xlsx file, or a list of xlsx files to process.\n\n")

parser_verify = subParsers.add_parser('INQUIRE', help='Inquire as to whether Ref Links in given file can be fixed.\n Exit with failure if any are unfixable or not found (depending on strictness)\n   strict=1: Fail if any ext ref link is improper\n   strict=2: Same as 1, but also fail if file not found.\n\n', formatter_class=argparse.RawTextHelpFormatter)
parser_verify.add_argument("-s", "--strict", help="Different strictness levels\n  1: fail when an unfixable path encountered\n  2: fail when file referred to by path can't be found in `cims-models` folder hierarchy.\n\n", type=int, action="store", default=1, choices=[1,2])
parser_verify.add_argument("files", nargs="*", help="If a path to a single folder, process all files in hierarchy rooted there\n  (leave blank to start at `cims-models` top level).\n  Can also be a single xlsx file, or a list of xlsx files to process.\n\n")

parser_verify = subParsers.add_parser('VERIFY', help='Verify that ext ref links are relative and possibly whether the referenced files are found.\n Exit with failure if any are not relative, or not found (depending on strictness)\n   strict=1: Fail if any ext ref link is not a relative path\n   strict=2: Same as 1, but also fail if referenced file not found.\n\n', formatter_class=argparse.RawTextHelpFormatter)
parser_verify.add_argument("-s", "--strict", help="Different strictness levels\n  1: fail when a non-relative path encountered\n  2: fail when referred file can't be found in `cims-models` folder hierarchy.\n\n", type=int, action="store", default=1, choices=[1,2])
parser_verify.add_argument("files", nargs="*", help="If a path to a single folder, process all files in hierarchy rooted there\n  (leave blank to start at `cims-models` top level).\n  Can also be a single xlsx file, or a list of xlsx files to process.\n\n")

parser_inspect = subParsers.add_parser('INSPECT', help='Write given file\'s ext ref links to screen.\n\n', formatter_class=argparse.RawTextHelpFormatter)
parser_inspect.add_argument("files", nargs="*", help="If a path to a single folder, process all files in hierarchy rooted there\n  (leave blank to start at `cims-models` top level).\n  Can also be a single xlsx file, or a list of xlsx files to process.\n\n")



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

# This is an extra level of logging, mostly for debugging how the paths tokens are generated.
# No longer really necessary, but leaving here just in case. To use, uncomment here and then
# the actual logging calls in the `tokenizeWindows` and `tokenizeNormal` functions below.
#
#logging.basicConfig(filename="path_token_info.log", 
#                    level=logging.INFO,
#                    filemode='w'  # This wipes and restarts the log on each load of this.
#                    )
#winlogger = logging.getLogger("winlogger")
#normlogger = logging.getLogger("normlogger")

class CIMSModelsNotFound(Exception):
    """
    How to fail when `cims-models` is not found within the external link
    path we're trying to correct.
    """
    pass

class EmptyPathError(Exception):
    """
    How to fail when the parsed path doesn't contain anything. I don't think this case can
    actually happen anymore, but this exception is still here and handled everywhere.
    """
    pass

class ExtRefFileNotFound(Exception):
    """
    How to fail when the file referred to by an excel external reference isn't found within
    the users `cims-models` file hierarchy. This is only an error at certain levels of 
    strictness.
    """
    pass

class PathNotRelative(Exception):
    """
    How to fail under strict conditions (>=3), when the path encountered in the xlsx file
    isn't relative.
    """
    pass

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

    #winlogger.info(f",{isAbsPath},{p3[0]},{str(p3).replace(',','|')},{p}")

    if isAbsPath:
        try:
            startInd = p3.index('cims-models')
        except:
            raise CIMSModelsNotFound(f"'cims-models' not found in external ref path: {p}")
        # Get rid of all the file path items above 
        # and including `cims-models`.
        p4 = p3[(startInd+1):]
    else:
        p4 = p3

    return((p4, isAbsPath))

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

    #normlogger.info(f",{isAbsPath},{p2[0]},{str(p2).replace(',','|')},{p}")

    if isAbsPath:
        try:
            startInd = p2.index('cims-models')
        except:
            raise CIMSModelsNotFound(f"'cims-models' not found in external ref path: {p}")
        # Get rid of all the file path items above
        # and including`cims-models`.
        p3 = p2[(startInd+1):]
    else:
        p3 = p2

    return((p3, isAbsPath))


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
        exPathTokens, isAbsPath = tokenizeWindows(filePath, cmRoot)
    except UnexpectedFwdSlashes:
        exPathTokens, isAbsPath = tokenizeNormal(filePath, cmRoot)

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
        raise CIMSModelsNotFound(f"'cims-models' not found in external ref path: {p}")

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
        raise CIMSModelsNotFound(f"'cims-models' not found in external ref path: {p}")

    # Get rid of all the file path items above
    # and including`cims-models`.
    p3 = p2[(startInd+1):]

    #p4 = os.path.join( *p3 )
    p4 = "/".join([".." for i in range(0, num_backups)]) + "/" + "/".join(p3)

    return(p4)


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


def check_path_relative(p):
    # Here are all the necessary windows-style filepath conditions, and we want to fail if we
    # find any of these things.
    if re.match("^file:/", p):
        return(False)
    if re.search("\\\\", p):
        return(False)
    if re.match("[a-zA-Z]:", p):
        return(False)

    # So we almost certainly don't have a windows-style filepath, so we split it by forward
    # slashes, and make sure the first item in the list isn't "", as this only happens when the
    # path begins with a forward slash (i.e. is an absolute path). We accept anything else as a
    # valid relative path.

    p2 = re.split("/", p)
    if not len(p2) >= 1:
        return(False)
    if p2[0] == "":
        return(False)

    return(True)



#############################################################
#############################################################
#############################################################
# A couple of useful loggers
#
def log_info( msg=None, xl_path_val=None, xl_path=None, xl_path_corrected=None, xl_filename=None, pathOk=None, fileFound=None ):
    with open("excel_external_reference_correction.csv", "a") as f:
        def noneToStr(m):
            return( "" if (m is None) else m )
        f.write(f"OK,{pathOk},{fileFound},{xl_path_val != xl_path_corrected},{xl_filename},{xl_path_val},{xl_path},{xl_path_corrected},{noneToStr(msg)}\n")

def log_failure( msg=None, xl_path_val=None, xl_path=None, xl_filename=None, pathOk=None, fileFound=None):
    with open("excel_external_reference_correction.csv", "a") as f:
        def noneToStr(m):
            return( "" if (m is None) else m )
        f.write(f"FAIL,{pathOk},{fileFound},True,{xl_filename},{xl_path_val},{xl_path},,{noneToStr(msg)}\n")

#############################################################
#############################################################
#############################################################
# The main function for link correction.
#

def correct_ext_links(xl_path, pathToCimsModels, pathLookupTable, corrExt=None, cmdArgs=None):

    wb = op.open(xl_path)

    path_tokens = tokenizePath(os.path.abspath(xl_path))
    cimsModels_index = path_tokens.index('cims-models')
    path_trunc = path_tokens[(cimsModels_index+1):]
    num_backups = len(path_trunc) - 1

    for index in range(0, len(wb._external_links)):
        oldPath = wb._external_links[index].file_link.target


        # In strictest condition, we want to fail HERE if the path right out of the spreadsheet is not a relative path.
        if hasattr(cmdArgs, 'checkRelative') and cmdArgs.checkRelative and (not check_path_relative(oldPath)):
            raise PathNotRelative(f"Excel file at {xl_path} contains a non-relative path, {oldPath}.")

        try:
            try:
                # Tokenizes RELATIVE to cims-models
                oldPathTokens, isAbsPath = tokenizeWindows(oldPath, pathToCimsModels)
            except UnexpectedFwdSlashes:
                oldPathTokens, isAbsPath = tokenizeNormal(oldPath, pathToCimsModels)

            # Ah... make sure these are dealt with correctly. The relative external links are from
            # the perspective of the excel files, but below this was being dropped into `os.path.abspath`
            # which is from the perspective of where this script is running, so it was backing out too
            # far and failing.

            if isAbsPath:
                fullExtRefPath = os.path.abspath(
                    os.path.join(
                        pathToCimsModels,
                        os.path.join(*oldPathTokens)
                )).replace("%20", " ")
            else:
                fullExtRefPath = os.path.abspath(
                    os.path.join(
                        os.path.split(xl_path)[0],
                        os.path.join(*oldPathTokens)
                )).replace("%20", " ")



            #fullExtRefPath = os.path.abspath(os.path.join(*oldPathTokens)).replace("%20", " ")

            #from IPython import embed; embed(header="check here: ")

        except CIMSModelsNotFound as e:
            if not cmdArgs.no_log:
                log_failure(xl_path_val=oldPath, xl_path=None, xl_filename=xl_path, msg="cims-models not in path", pathOk=False, fileFound=False)
            if cmdArgs.strict >= 1:
                if len(e.args) >= 1:
                    e.args = (e.args[0] + f", Excel file: {xl_path}",) + e.args[1:]
                raise
            continue
        except EmptyPathError:
            if not cmdArgs.no_log:
                log_failure( msg="Path is empty.", xl_path_val=oldPath, xl_path=None, xl_filename=xl_path, pathOk=False, fileFound=False)
            if cmdArgs.strict >= 1 :
                raise
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
                    if not cmdArgs.no_log:
                        log_failure(xl_path_val=oldPath, xl_path=pathToUse, xl_filename=xl_path, msg="IMPOSSIBLE. cims-models not in path", pathOk=False, fileFound=True)
                    if cmdArgs.strict >= 1:
                        if len(e.args) >= 1:
                            e.args = (e.args[0] + f", Excel file: {xl_path}",) + e.args[1:]
                        raise
                    continue
                except EmptyPathError:
                    if not cmdArgs.no_log:
                        log_failure( msg="Path is empty.", xl_path_val=oldPath, xl_path=pathToUse, xl_filename=xl_path, pathOk=False, fileFound=True)
                    if cmdArgs.strict >= 1:
                        raise
                    continue
            except CIMSModelsNotFound:
                if not cmdArgs.no_log:
                    log_failure(xl_path_val=oldPath, xl_path=pathToUse, xl_filename=xl_path, msg="cims-models not in path", pathOk=False, fileFound=True)
                if cmdArgs.strict >= 1:
                    if len(e.args) >= 1:
                        e.args = (e.args[0] + f", Excel file: {xl_path}",) + e.args[1:]
                    raise
                continue
            except EmptyPathError:
                if not cmdArgs.no_log:
                    log_failure( msg="Path is empty.", xl_path_val=oldPath, xl_path=pathToUse, xl_filename=xl_path, pathOk=False, fileFound=True)
                if cmdArgs.strict >= 1:
                    raise
                continue
            
            if not cmdArgs.no_log:
                log_info(xl_path_val=oldPath, xl_path=pathToUse, xl_path_corrected=newPath, xl_filename=xl_path, msg="Path OK and File Found", pathOk=True, fileFound=True)

        else:
            # The path doesn't exist. On the filesystem.
            # In this case we want to prepare it for use by Excel AS IF IT WERE here on the
            # filesystem. I guess in case it's just running late and shows up later? But anyway,
            # we log that this has occurred (if logging enabled).

            # If the strict level is 2, we fail here
            if cmdArgs.strict >= 2:
                raise ExtRefFileNotFound(f"File referred to cannot be found in `cims-models` hierarchy: {fullExtRefPath}")

            pathToUse = fullExtRefPath
            try:
                newPath = repathWindowsRelative(pathToUse, pathToCimsModels, num_backups)
            except UnexpectedFwdSlashes:
                try:
                    newPath = repathNormalRelative(pathToUse, pathToCimsModels, num_backups)
                except CIMSModelsNotFound:
                    if not cmdArgs.no_log:
                        log_failure(msg="Filepath not found on system AND `cims-models` not in filepath.",xl_path_val=oldPath, xl_path=pathToUse, xl_filename=xl_path, pathOk=False, fileFound=False)
                    if cmdArgs.strict >= 1:
                        if len(e.args) >= 1:
                            e.args = (e.args[0] + f", Excel file: {xl_path}",) + e.args[1:]
                        raise
                    continue
                except EmptyPathError:
                    if not cmdArgs.no_log:
                        log_failure( msg="Path is empty.", xl_path_val=oldPath, xl_path=pathToUse, xl_filename=xl_path, pathOk=False, fileFound=False)
                    if cmdArgs.strict >= 1:
                        raise
                    continue
            except CIMSModelsNotFound:
                if not cmdArgs.no_log:
                    log_failure(msg="Filepath not found on system AND `cims-models` not in filepath.", xl_path_val=oldPath, xl_path=pathToUse, xl_filename=xl_path, pathOk=False, fileFound=False)
                if cmdArgs.strict >= 1:
                    if len(e.args) >= 1:
                        e.args = (e.args[0] + f", Excel file: {xl_path}",) + e.args[1:]
                    raise
                continue
            except EmptyPathError:
                if not cmdArgs.no_log:
                    log_failure( msg="Path is empty.", xl_path_val=oldPath, xl_path=pathToUse, xl_filename=xl_path, pathOk=False, fileFound=False)
                if cmdArgs.strict >= 1:
                    raise
                continue
            if not cmdArgs.no_log:
                log_info(xl_path_val=oldPath, xl_path=pathToUse, xl_path_corrected=newPath, xl_filename=xl_path, msg="Path OK but File Not Found", pathOk=True, fileFound=False)

        if not cmdArgs.dry_run:
            wb._external_links[index].file_link.target = newPath
            wb._external_links[index].file_link.id = 'rId1'

    if corrExt is None:
        xl_path_corrected = xl_path
    else:
        temp_splt = os.path.split(xl_path)
        temp_f_splt = os.path.splitext(temp_splt[1])
        xl_path_corrected = os.path.join(temp_splt[0], temp_f_splt[0] + corrExt + temp_f_splt[1])

    if not cmdArgs.dry_run:
        wb.save(xl_path_corrected)
    
    wb.close()


def inspect_ext_links(xl_path, *args, **kwargs):
    """
    Write the full path (`xl_path`) and all the contained external reference
    links to stdOut.
    """
    wb = op.open(xl_path)
    print(f"Excel file: {xl_path}")
    if len(wb._external_links) == 0:
        print("  No external reference links found")
    else:
        for index in range(0, len(wb._external_links)):
            pathToTest = wb._external_links[index].file_link.target
            print(f"  link_{index+1}: {pathToTest}")
    print()



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
            if (os.path.splitext(f)[1] == ".xlsx") or (os.path.splitext(f)[1] == ".xlsb"):
                fullSysPath = os.path.join(dirpath, f)
                fpDict[fullSysPath.lower()] = fullSysPath
            else:
                pass
    return(fpDict)


def iterate_list(pathList,
                 pathToCimsModels, 
                 pathLookupTable, 
                 corrExt=None,
                 correctionFunc=correct_ext_links,
                 cmdArgs=None):
    """
    `pathList`: a list of relative paths to xlsx files, rooted at `cims-models`.
    `pathToCimsModels`: (unused, but just in case) Absolute path to `cims-models` dir.
    `corrExt`: suffix to append to corrected file name. If `None`, corrected file will
               be saved under the same name. (Leave `None` unless testing).
    """

    for p in pathList:
        if (not os.path.isfile(p)) or (os.path.splitext(p)[1] != ".xlsx"):
            raise RuntimeError(f"Supplied path is not xlsx file: {p}")
        correctionFunc(p, pathToCimsModels, pathLookupTable, corrExt, cmdArgs=cmdArgs)


if __name__ == "__main__":

    time_start = datetime.now()

    args = argParser.parse_args()

    ###############################################
    ############################################### INSPECT
    ###############################################

    if args.command_name == "INSPECT":
        if len(args.files) == 0:
            iterate_list(find_excel_files('.'), None, pathLookupTable=None, correctionFunc=inspect_ext_links)


        elif len(args.files) == 1:
            if os.path.isdir(args.files[0]):
                iterate_list(find_excel_files(args.files[0]), None, pathLookupTable=None, correctionFunc=inspect_ext_links)


            elif os.path.isfile(args.files[0]):
                inspect_ext_links(args.files[0])

            else:
                raise RuntimeError("Argument seems to be neither a directory nor a file.")
        else:
            # Here we assume that all positional arguments are pathnames pointing to specific 
            # excel files. There can be any number of these. We correct the external references in
            # each of them.
            iterate_list(args.files, None, pathLookupTable=None, correctionFunc=inspect_ext_links)

    ###############################################
    ############################################### INQUIRE
    ###############################################

    elif args.command_name == "INQUIRE":

        localArgs_dict = {'dry_run':True, 'no_log':True, 'strict':args.strict}
        
        # We need to make this weird object so that we can fake an `argparse` `args` object using this dictionary
        # above. This is one way to do this.
        class DynObj:
            def __init__(self, data):
                for k,v in data.items():
                    setattr(self,k,v)
        localArgs = DynObj(localArgs_dict)

        whereAreWe = os.path.abspath('.')
        if os.path.split(whereAreWe)[1] != 'cims-models':
            raise RuntimeError("This script needs to be run at the top level of the `cims-models` directory")

        # This lookup table contains the paths with cases as they actually are on the system. The keys
        # in this table are these paths.lowered().
        pathLookupTable = buildLookupTable(whereAreWe)

        if len(args.files) == 0:
            iterate_list(find_excel_files('.'), whereAreWe, pathLookupTable, correctionFunc=correct_ext_links, cmdArgs=localArgs)


        elif len(args.files) == 1:
            if os.path.isdir(args.files[0]):
                iterate_list(find_excel_files(args.files[0]), whereAreWe, pathLookupTable, correctionFunc=correct_ext_links, cmdArgs=localArgs)


            elif os.path.isfile(args.files[0]):
                correct_ext_links(args.files[0], whereAreWe, pathLookupTable, cmdArgs=localArgs)

            else:
                raise RuntimeError("Argument seems to be neither a directory nor a file.")
        else:
            # Here we assume that all positional arguments are pathnames pointing to specific 
            # excel files. There can be any number of these. We correct the external references in
            # each of them.
            iterate_list(args.files, whereAreWe, pathLookupTable, correctionFunc=correct_ext_links, cmdArgs=localArgs)

    ###############################################
    ############################################### VERIFY
    ###############################################

    elif args.command_name == "VERIFY":

        localArgs_dict = {'dry_run':True, 'no_log':True, 'strict': args.strict, 'checkRelative':True}

        # We need to make this weird object so that we can fake an `argparse` `args` object using this dictionary
        # above. This is one way to do this.
        class DynObj:
            def __init__(self, data):
                for k,v in data.items():
                    setattr(self,k,v)
        localArgs = DynObj(localArgs_dict)

        whereAreWe = os.path.abspath('.')
        if os.path.split(whereAreWe)[1] != 'cims-models':
            raise RuntimeError("This script needs to be run at the top level of the `cims-models` directory")

        pathLookupTable = buildLookupTable(whereAreWe)

        if len(args.files) == 0:
            iterate_list(find_excel_files('.'), whereAreWe, pathLookupTable, correctionFunc=correct_ext_links, cmdArgs=localArgs)

        elif len(args.files) == 1:
            if os.path.isdir(args.files[0]):
                iterate_list(find_excel_files(args.files[0]), whereAreWe, pathLookupTable, correctionFunc=correct_ext_links, cmdArgs=localArgs)

            elif os.path.isfile(args.files[0]):
                correct_ext_links(args.files[0], whereAreWe, pathLookupTable, cmdArgs=localArgs)

            else:
                raise RuntimeError("Argument seems to be neither a directory nor a file.")

        else:
            # Here we assume that all positional arguments are pathnames pointing to specific 
            # excel files. There can be any number of these. We correct the external references in
            # each of them.
            iterate_list(args.files, whereAreWe, pathLookupTable, correctionFunc=correct_ext_links, cmdArgs=localArgs)

    ###############################################
    ############################################### CORRECT
    ###############################################

    elif args.command_name == "CORRECT":

        # Reset the info/failure log files
        if not args.no_log:
            with open("excel_external_reference_correction.csv", "w") as f:
                def isMsg(m):
                    return( "" if (m is None) else m )
                #f.write(f"OK,{xl_filename},{xl_path},{xl_path_corrected},{isMsg(msg)}\n")
                f.write(f"CorrectionStatus,PathOK,FileFound,CorrectionNeeded,Filename,ExtLinkVal,ExtLinkPath,ExtLinkPath_corrected,Message/Error\n")

        whereAreWe = os.path.abspath('.')
        if os.path.split(whereAreWe)[1] != 'cims-models':
            raise RuntimeError("This script needs to be run at the top level of the `cims-models` directory")

        # This lookup table contains the paths with cases as they actually are on the system. The keys
        # in this table are these paths.lowered().
        pathLookupTable = buildLookupTable(whereAreWe)

        if len(args.files) == 0:
            iterate_list(find_excel_files('.'), whereAreWe, pathLookupTable, correctionFunc=correct_ext_links, cmdArgs=args)


        elif len(args.files) == 1:
            if os.path.isdir(args.files[0]):
                iterate_list(find_excel_files(args.files[0]), whereAreWe, pathLookupTable, correctionFunc=correct_ext_links, cmdArgs=args)


            elif os.path.isfile(args.files[0]):
                correct_ext_links(args.files[0], whereAreWe, pathLookupTable, cmdArgs=args)

            else:
                raise RuntimeError("Argument seems to be neither a directory nor a file.")
        else:
            # Here we assume that all positional arguments are pathnames pointing to specific 
            # excel files. There can be any number of these. We correct the external references in
            # each of them.
            iterate_list(args.files, whereAreWe, pathLookupTable, correctionFunc=correct_ext_links, cmdArgs=args)

        time_end = datetime.now()
        print(f"Finished link correction in {time_end - time_start}.")

    else:
        raise RuntimeError("Incorrect command, must be [CORRECT, VERIFY, INSPECT]")

