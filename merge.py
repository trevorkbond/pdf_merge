from pypdf import PdfMerger, PdfReader
import re, os
import shutil
from os.path import expanduser

def findTemplateFile(files, substring):
    matchingFile = None
    for filename in files:
        if substring in filename:
            matchingFile = filename
            break

    return matchingFile

def find_folder(start_path, target_folder):
    for root, dirs, files in os.walk(start_path):
        if target_folder in dirs:
            return os.path.join(root, target_folder)
    
    return None

def findPDFPage(searchKey, pdf):
    pageFound = -1
    for i in range(0, len(pdf.pages)):
        pageContent = ""
        pageContent += pdf.pages[i].extract_text()
        # content = pageContent.encode('ascii', 'ignore').lower()
        reSearch = re.search(searchKey, pageContent)
        # print(pageContent + "\n")
        if reSearch is not None:
            # print("Found!")
            pageFound = i
            break
    return pageFound

def getKeyList(allFiles, keysAndFiles):
    """Searches for and places list of all files with matching key in dict"""
    for searchKey in keysAndFiles:
        keysAndFiles[searchKey] = [files for files in allFiles if searchKey in files]

def appendPDFs(key, files, offset, merger, scanTemplate):
    """Appends all found files after key's corresponding pdf bookmark"""
    merger.append(scanTemplate, pages = (offset))

    return offset

def getBookmarkDistances(bookmarkDistances):
    """retrieves list of numpages between bookmarks"""
    positions = []
    for position in bookmarkDistances.values():
        positions.append(position)
    distances = [positions[i + 1] - positions[i] for i in range(0, (len(positions) - 1))]
    print(positions)
    distances.insert(0, 1)
    print(distances)
    return distances

def mergePDFs(keysAndFiles, bookmarkLocations, sourceName_bookmarkName):
    i = 0
    merger = PdfMerger()
    for key, files in keysAndFiles.items():
        # print("Merger looping on key: " + key)
        bookmarkName = sourceName_bookmarkName[key]
        bookmarkLocation = bookmarkLocations[bookmarkName]
        # print("name: " + bookmarkName + " location: " + str(bookmarkLocation))
        print("i is " + str(i))
        if key != 'STMT CC':                # STMT CC falls under same bookmark as STMT B, so it would duplicate appending that
            if i >= len(distances):
                print("We actually made it in here")
                merger.append(scanTemplatePDF, pages=(bookmarkLocation, bookmarkLocation + 1))
            else:
                print("Appending " + bookmarkName + ", which was on scanTemplate page " + str(bookmarkLocation) + " and goes until " + str(bookmarkLocation + distances[i]))
                merger.append(scanTemplatePDF, pages=(bookmarkLocation, (bookmarkLocation + distances[i])))
        for file in files:
            merger.append(file)
        i += 1

    return merger

# defining bookmark titles
bookmarkLocations = {
    'Checklist\n& Overview' : 0,
    'Balance Sheet compared to PP' : 0,
    'Reconciliations\nPeriod Reconciliation Reports' : 0,
    'Bank and Credit' : 0,
    'Adjusting Journal' : 0,
    'Journal Entry \nSource Docs' : 0,
    'Payroll Reports' : 0,
    'Depreciation &' : 0,
    'Statements &' : 0,
    'Retained Earnings' : 0,
    'Other Source' : 0,
    'during the period' : 0
}
    
# defining keys to search for
keysAndFiles = {
    "Overview" : [],
    "Financials" : [],
    "REC" : [],
    "STMT B" : [],
    "STMT CC" : [],
    "AJE R" : [],
    "OI AJE" : [],
    "PR" : [],
    "DEP" : [],
    "STMT L" : [],
    "SCH L" : [],
    "OI" : [],
    "CLIENT COM" : []
}

# creating connection between source doc naming and bookmark naming
sourceName_bookmarkName = {
    "Overview" : 'Checklist\n& Overview',
    "Financials" : 'Balance Sheet compared to PP',
    "REC" : 'Reconciliations\nPeriod Reconciliation Reports',
    "STMT B" : 'Bank and Credit',
    "STMT CC" : 'Bank and Credit',
    "AJE R" : 'Adjusting Journal',
    "OI AJE" : 'Journal Entry \nSource Docs',
    "PR" : 'Payroll Reports',
    "DEP" : 'Depreciation &',
    "STMT L" : 'Statements &',
    "SCH L" : 'Retained Earnings',
    "OI" : 'Other Source',
    "CLIENT COM" : 'during the period'
}

TEMPLATE_SUBSTRING = 'Scan Template'

# get list of all files in SCAN folder and find template file
user_home = expanduser("~")
desktop_path = os.path.join(user_home, "Desktop")
target_folder = "SCAN"
folder_path = find_folder(desktop_path, target_folder)
if folder_path:
    folder_path += "/"
    print(f"Found the folder at: {folder_path}")
else:
    print("Folder not found on the desktop.")
    exit()
allFilesNotFull = os.listdir(folder_path)
print('\n'.join(allFilesNotFull))
allFiles = [folder_path + file for file in allFilesNotFull]
print('\n'.join(allFiles))
scanTemplate = findTemplateFile(allFiles, TEMPLATE_SUBSTRING)
scanTemplatePDF = PdfReader(scanTemplate)
# print(scanTemplate)

# find location of all bookmarks in template file
for key in bookmarkLocations:
    bookmarkLocations[key] = findPDFPage(key, scanTemplatePDF)

# get list of all files and associate with key
getKeyList(allFiles, keysAndFiles)
print("Printing keys and file lists")
print(keysAndFiles)

# get distance between bookmarks
distances = getBookmarkDistances(bookmarkLocations)

merged = mergePDFs(keysAndFiles, bookmarkLocations, sourceName_bookmarkName)

merged.write(desktop_path + '/result.pdf')
merged.close()

# os.mkdir(desktop_path + '/Source Docs')
# for file in allFiles:
#     shutil.move(file, desktop_path + '/Source Docs')