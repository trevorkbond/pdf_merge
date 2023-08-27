from pypdf import PdfMerger, PdfReader
import re, os
import tkinter as tk
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

def mergePDFs(keysAndFiles, bookmarkLocations, sourceName_bookmarkName, distances, scanTemplatePDF):
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

# begin functions that trigger PDF merging
def findFolder(folderName, keysAndFiles, bookmarkLocations, sourceName_bookmarkName, template_name, output_name):
    user_home = expanduser("~")
    desktop_path = os.path.join(user_home, "Desktop")
    target_folder = folderName
    folder_path = find_folder(desktop_path, target_folder)
    if folder_path:
        folder_path += "/"
        print(f"Found the folder at: {folder_path}")
        getFileList(folder_path, keysAndFiles, bookmarkLocations, sourceName_bookmarkName, desktop_path, template_name, output_name)
    else:
        print("Folder not found on the desktop.")
        exit()

def getFileList(folder_path, keysAndFiles, bookmarkLocations, sourceName_bookmarkName, desktop_path, template_name, output_name):
    allFilesNotFull = os.listdir(folder_path)
    allFiles = [folder_path + file for file in allFilesNotFull]
    scanTemplate = findTemplateFile(allFiles, template_name)
    scanTemplatePDF = PdfReader(scanTemplate)
    findBookmarks(scanTemplatePDF, keysAndFiles, bookmarkLocations, sourceName_bookmarkName, allFiles, desktop_path, output_name)

def findBookmarks(scanTemplatePDF, keysAndFiles, bookmarkLocations, sourceName_bookmarkName, allFiles, desktop_path, output_name):
    for key in bookmarkLocations:
        bookmarkLocations[key] = findPDFPage(key, scanTemplatePDF)
    getKeyList_Merge(scanTemplatePDF, allFiles, keysAndFiles, bookmarkLocations, sourceName_bookmarkName, desktop_path, output_name)

def getKeyList_Merge(scanTemplatePDF, allFiles, keysAndFiles, bookmarkLocations, sourceName_bookmarkName, desktop_path, output_name):
    """Searches for and places list of all files with matching key in dict"""
    for searchKey in keysAndFiles:
        keysAndFiles[searchKey] = [files for files in allFiles if searchKey in files]
    distances = getBookmarkDistances(bookmarkLocations)
    merged = mergePDFs(keysAndFiles, bookmarkLocations, sourceName_bookmarkName, distances, scanTemplatePDF)
    merged.write(desktop_path + '/' + output_name + '.pdf')
    merged.close()

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

# creating tkinter UI
window= tk.Tk()
window.title("PDF Merger")

window.columnconfigure(0, weight=1, minsize=500)
window.rowconfigure(0, weight=1) # top banner row
window.rowconfigure(1, weight=1) # scan folder name
window.rowconfigure(2, weight=1) # template file name
window.rowconfigure(3, weight=1) # output file name
window.rowconfigure(4, weight=1) # merge PDF button

frm_banner = tk.Frame(master=window, relief=tk.RAISED, borderwidth=1)
lbl_banner = tk.Label(master=frm_banner, text="Welcome to PDF Merger, please enter the information below.")
lbl_banner.pack(padx=10)
frm_banner.grid(row=0, column=0, padx=10, sticky="ew")

frm_scanfolder = tk.Frame(master=window, relief=tk.RAISED, borderwidth=1)
lbl_scanfolder = tk.Label(master=frm_scanfolder, text="Scan folder name: ")
lbl_scanfolder.pack(side=tk.LEFT, padx=5)
ent_scanfolder = tk.Entry(master=frm_scanfolder)
ent_scanfolder.pack(side=tk.LEFT)
frm_scanfolder.grid(row=1, column=0, padx=10, sticky="ew")

frm_templatefile = tk.Frame(master=window, relief=tk.RAISED, borderwidth=1)
lbl_templatefile = tk.Label(master=frm_templatefile, text="Scan Template file name (including .pdf): ")
lbl_templatefile.pack(side=tk.LEFT, padx=5)
ent_templatefile = tk.Entry(master=frm_templatefile)
ent_templatefile.pack(side=tk.LEFT)
frm_templatefile.grid(row=2, column=0, padx=10, sticky="ew")

frm_opfile = tk.Frame(master=window, relief=tk.RAISED, borderwidth=1)
lbl_opfile = tk.Label(master=frm_opfile, text="Desired output file name (without .pdf): ")
lbl_opfile.pack(side=tk.LEFT)
ent_opfile = tk.Entry(master=frm_opfile)
ent_opfile.pack(side=tk.LEFT)
frm_opfile.grid(row=3, column=0, padx=10, sticky="ew")

btn_mergePDFs = tk.Button(master=window, text="Merge PDFs", command=lambda: findFolder(ent_scanfolder.get(), keysAndFiles, bookmarkLocations, sourceName_bookmarkName, ent_templatefile.get(), ent_opfile.get()))
btn_mergePDFs.grid(row=4, column=0)

window.mainloop()
