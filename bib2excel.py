from pybtex.database.input import bibtex
import pybtex.errors
import argparse
import pandas as pd

parser = argparse.ArgumentParser(description='Transform bibtex into excel.')
parser.add_argument('pathtofile', metavar='fname', type=str, nargs=1,
                    help='Path to the bibtex file')

args = parser.parse_args()

pybtex.errors.set_strict_mode(False)

def case_insensitive_unique_list(data):
    seen, result = set(), []
    for item in data:
        if item.lower() not in seen:
            seen.add(item.lower())
            result.append(item)
    return result

if args.pathtofile:
    print(args.pathtofile)

    bparser = bibtex.Parser()
    bibdata = bparser.parse_file(args.pathtofile[0])

    myfield = []

    for bibid in bibdata.entries:
        b = bibdata.entries[bibid].fields
        for field in b:
            myfield.append(field)
    myfield.append("authors")
    myfield = case_insensitive_unique_list(myfield)
    print(myfield)

    mydict = { field : [] for field in myfield}

    for bibid in bibdata.entries:
        b = bibdata.entries[bibid].fields
        for field in myfield:
            if field == "authors":
                try:
                    authors = ""
                    for author in bibdata.entries[bibid].persons["author"]:
                        authors += "%s %s,"%(author.first()[0], author.last()[0])
                    mydict[field].append(authors)
                except:
                    mydict[field].append("")
            else:
                try:
                    mydict[field].append(b[field])
                except:
                    mydict[field].append("")

    df = pd.DataFrame (mydict, columns = myfield)

    df.to_excel("bib2excel.xlsx")