import pdfplumber
import sys
import openpyxl
import argparse
import os
from itertools import groupby
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdftypes import resolve1
from xmp import xmp_to_dict
from collections import defaultdict

# https://towardsdatascience.com/how-to-extract-text-from-pdf-245482a96de7
# https://stackoverflow.com/questions/25864774/how-to-read-the-remaining-of-a-command-line-with-argparse

# reverse if numbers
def revornot(s):
    if s[0].isdigit():
        return s[::-1]
    return s

# top level pdf process
def pdf_process(f):
    w = pdf_detect(f)
    if w:
        if w.startswith('cal'):
            try:
                print(f, w, ":", process_cal(f, w))
            except:
                print(f, w)
        elif w.startswith('leumi'):
            try:
                print(f, w, ":", process_leumi(f, w))
            except:
                print(f, w)

img_size = defaultdict(int)
# detect document type
def pdf_detect(f):
    global img_size
    if not f.endswith("pdf"):
        return None
    fp = open(f, 'rb')
    parser = PDFParser(fp)
    doc = PDFDocument(parser)
    parser.set_document(doc)
    #print(dir(doc))
    #print(doc.xrefs)
    #print(doc.info)  # The "Info" metadata
    #print(doc.catalog)
    #if 'Metadata' in doc.catalog:
    #    metadata = resolve1(doc.catalog['Metadata']).get_data()
    #    print(metadata)  # The raw XMP metadata
    #    print(xmp_to_dict(metadata))
    with pdfplumber.open(f) as pdf:
        first_page = pdf.pages[0]
        if len(first_page.images) > 0:
            w = round(first_page.images[0]['width'], 8)
            h = round(first_page.images[0]['height'], 8)
            i = "{}_{}".format(h, w)
            img_size[i] = img_size[i]+1

        #sys.exit(1)
        s = first_page.extract_words(horizontal_ltr=False)
        if len(s) > 1:
            if s[0]['text'].find('ו8887729-30li.oc.xam') >= 0:
                return "Max"
            if s[-2]['text'] == 'לאומי' and s[-1]['text'] == 'איתך.':
                return "leumi2"
            for (i,ss) in enumerate(s):
                if ss['text'] == 'laC':
                    return "cal({})".format(i)
                    break
                #if s[i]['text'] == 'לכבוד':
                #    print(s[i]['x1'])
                #    return
                if ss['text'] == 'רסמ' or ss['text'] == 'מסר':
                    return "cal2({})".format(i)
                if ss['text'].endswith('-01') and s[i-2]['text']== 'לקוח':
                    return "leumi({})".format(i)
                if ss['text'].endswith('-01') and s[i-1]['text']== 'לקוח':
                    return "leumi({})".format(i)
                #print(ss['text'], ss['text'] == 'לקוח')
                #print(ss['text'], ss['text'].endswith('-01'))
            #else:
            #    for i in range(0,len(s)):
            #        print (i, "**{}**".format(s[i]['text']), s[i]['x0'], s[i]['x1'], s[i]['top'], s[i]['bottom'])
            #    print(f)
            #    sys.exit(1)
            #    return "Unknown"

# return nested 
def group_sort_words(s):
    # group all words on the same line into lists
    words2d = []
    for _, g in groupby(s, lambda x: (x['top'])):
        words2d.append(list(g))
    # sort each line right to left
    for g in words2d:
        g.sort(reverse=True, key=lambda x: x['x1'])
    return words2d

# process CAL credit reports
def process_cal(f, w):
    attrs = {}
    with pdfplumber.open(f) as pdf:
        if w.startswith('cal2'):
            first_page = pdf.pages[1]
        else:
            first_page = pdf.pages[0]
        s = first_page.extract_words(horizontal_ltr=False)
        words2d = group_sort_words(s)
        # filter out anything not right aligned
        words2d = list(filter(lambda x: x[0]['x1'] > 500, words2d))
        # strip everything before לכבוד
        for (i, w) in enumerate(words2d):
            if w[0]['text'].startswith ('לכבוד'):
                words2d = words2d[i:]
                break
        w = words2d[0][0]['text']
        if w == 'לכבוד:':
            attrs['first'] = words2d[1][0]['text']
            attrs['last'] = words2d[1][1]['text']
        else:
            attrs['first'] = w.split(':')[1]
            attrs['last'] = words2d[1][0]['text']
        attrs['addr1'] =  ' '.join([revornot(s['text']) for s in words2d[2]][::-1])
        s = [revornot(s['text']) for s in words2d[3]]
        attrs['city'] = ' '.join(s[0:-1])
        attrs['zip'] =  s[-1]
        #print(words2d[0:5])
        return(attrs)

        #else:
        #    print(f, "is unknown")
    #else:
    #    print(f, "is scanned??")
    #except:
    #    print("Issues with file:", f)
        return
        #print(first_page.extract_tables({"horizontal_strategy":"text", "vertical_strategy":"lines"}))
        for i in range(0,min(len(s), 100)):
            print (i, s[i]['text'], s[i]['x0'], s[i]['y0'], s[i]['x1'], s[i]['y1'])
        #sys.exit(1)

def process_leumi(f, w):
    attrs = {}
    with pdfplumber.open(f) as pdf:
        first_page = pdf.pages[0]
        s = first_page.extract_words(horizontal_ltr=False)
        words2d = group_sort_words(s)
        if words2d[3][1]['text'] == 'לקוח':
            attrs['acct'] = revornot(words2d[3][2]['text'])
            attrs['branch'] = revornot(words2d[3][3]['text'][1:-3])
            attrs['date'] = [d.zfill(2) for d in revornot(words2d[5][-1]['text']).split('/')[::-1]]
        if [s['text'] for s in words2d[6][0:4]] == ['להלן', 'פרוט', 'חשבונך', 'בשקלים:']:
            attrs['page'] = revornot(words2d[7][5]['text']).split('/')[::-1]
        return(attrs)
#
# code starts running here
#
if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='My Balance')
    #parser.add_argument('-s', '--stats', action='store_true', help='Run on all the vocabulary and calculate average number of guesses for a solution')
    parser.add_argument('path', type=str, nargs="+", help='directories or files to parse')
    args = parser.parse_args()
    for p in args.path:
        if os.path.isdir(p):
            for root, dirs, files in os.walk(p):
                for f in files:
                    pdf_process(os.path.join(root, f))
        elif os.path.isfile(p):
            pdf_process(p)
        else:
            print(p, "???")
    for i in img_size:
        print (img_size[i])
    sys.exit(1)

