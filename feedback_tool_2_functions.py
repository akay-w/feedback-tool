# -*- coding: utf-8 -*-
"""
Created on Fri Mar  8 14:56:56 2019

@author: a-whalen
"""

import xml.etree.ElementTree as ET
import os
from collections import namedtuple
from difflib import Differ
import xlsxwriter

Segment = namedtuple('Segment', 'segid srctext tartext')

def get_xliff_list(directory):
    """
    Returns list of sdlxliff files in a given directory,
    including nested files in nested folders.
    
    directory: full path to directory to search for files
    """
    file_list = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith("sdlxliff"):
                file_list.append(os.path.join(os.path.normpath(root), file))
    return file_list

def parse_xliff(filepath):
    """
    Parses xliff file and returns a list of tuples containing the 
    segment id, source text, and target text for each segment.
    
    filepath = complete path to file to be parsed
    """
    tree = ET.parse(filepath)
    root = tree.getroot()

    trans = ""

    for x in root.iter():
        if "trans" in x.tag:
            trans = x.tag
            break
        
    transinfo = []
    
    for x in root.iter(trans):
        if x.attrib.get("translate")=="no":
            continue
        segid = x.attrib.get("id")
        for y in x.iter():
            if "seg-source" in y.tag:
                srctext = "".join(y.itertext())
            if "target" in y.tag:
                tartext = "".join(y.itertext())
        seg = Segment(segid, srctext, tartext)
        transinfo.append(seg)                 
    return transinfo

def get_difference(orig, edit):
    """
    Takes in original text and edited text,
    compares them with a differ
    object, and returns a list of changes
    
    orig: tokenized original text
    edit: tokenized edited text
    """
    d = Differ()
    result = d.compare(orig, edit)
    return [x for x in result]

def analyze(file, origtransinfo, edittransinfo, lang):
    """
    Takes in translation data from original and edited files,
    checks for changes, and returns a list of tuples, with each tuple 
    containing the filename, original source text, original target text,
    edited source text, edited target text, source changes (list),
    and target changes (list)
    
    origtransinfo: list of namedtuples (Segments) containing the 
    segid, origsrctext, and origtartext for each translation segment
    in the original sdlxliff file
    
    edittransinfo: list of namedtuples (Segments) containing the
    segid, origsrctext, and origtartext for each translation segment
    in the edited sdlxliff file
    """
    file_data_list = []
    for seg in range(len(origtransinfo)):
        src_changes = None
        tar_changes = None
        if origtransinfo[seg].segid == edittransinfo[seg].segid:
            origsrc = origtransinfo[seg].srctext
            editsrc = edittransinfo[seg].srctext
            if origsrc != editsrc:
                if lang == 1:
                    src_changes = get_difference(origsrc.split(), editsrc.split())
                else:
                    src_changes = get_difference(origsrc, editsrc)
                
            origtar = origtransinfo[seg].tartext
            edittar = edittransinfo[seg].tartext
            if origtar != edittar:
                if lang == 1:
                    tar_changes = get_difference(origtar.split(), edittar.split())
                else:
                    tar_changes = get_difference(origtar, edittar)
                    
        file_data_list.append((file, origsrc, origtar, editsrc, edittar, src_changes, \
                          tar_changes))
    return file_data_list

def print_changes(changes, orig_flag, red, body, lang):
            """
            Takes in list of changes, and outputs list of 
            formatted text to write in Excel with xlsxwriter if
            there is more than one token, and returns a string if there is
            only one token
            
            changes: list of changes output by differ object
            
            orig_flag: boolean, True for original text and False for edited text.
            Determines whether to search for added or deleted text in changes.
            
            red: xlsxwriter red text format
            
            body: xlsxwriter body text format
            
            lang: integer signifying European (1) or Asian (2) language
            """
            text = ""
            text_to_write = []
            if orig_flag:
                char = "-"
            else:
                char = "+"
                
            for token in changes:
                if token.startswith(char):
                    text = token[2:]
                    text_to_write.append(red)
                    if lang == 1:
                        text_to_write.append(token[2:] + " ")
                    elif lang == 2:
                        text_to_write.append(token[2:])
                elif token.startswith(" "):
                    if lang == 1:
                        text_to_write.append(token[2:] + " ")
                        if text != "":
                            text += " " + token[2:]
                        else:
                            text = token[2:]
                    elif lang == 2:
                        text_to_write.append(token[2:])
                        text += token[2:]
            text_to_write.append(body)
            if len(text_to_write) > 3:
                return text_to_write
            else:
                return text
            
def create_excel(savename, data_list, lang):
    """
    Creates an Excel report of the data, returns nothing.
    
    savename: string, complete filepath and name to save Excel file as
    
    data_list: list of tuples, with each tuple containing the 
    filename (string), original source text (string),
    original target text (string), edited source text (string),
    edited target text (string), source changes (list), and
    target changes (list) for each segment in each file
    
    lang: integer signifying European (1) or Asian (2) language
    """
    wb = xlsxwriter.Workbook(savename)
    ws = wb.add_worksheet()
    
    red = wb.add_format({'font_color': 'red','text_wrap': True, 'border': 1})
    header = wb.add_format({'text_wrap': True, 'align': 'center', 
                                  'bg_color': 'gray', 'border': 1})
    body = wb.add_format({'text_wrap': True, 'border': 1})
    
    ws.set_column('A:F', 45)
    ws.write('A1', "Filename", header)
    ws.write('B1', "Source (Original)",header)
    ws.write('C1', "Source (Edited)",header)
    ws.write('D1', "Target (Original)",header)
    ws.write('E1', "Target (Edited)",header)
    ws.write('F1', "Comments",header)
    
    filecol = 0
    origsrccol = 1
    editsrccol = 2
    origtarcol = 3
    edittarcol = 4
    commentcol = 5
    row = 1
    total_src_equal_flag = True
    
    for segment in data_list:
        file, origsrc, origtar, editsrc, edittar, src_changes, tar_changes = segment
        src_equal_flag = True
        tar_equal_flag = True
        if not origsrc or origsrc.isspace():
            continue
        ws.set_row(row, 100)
        ws.write(row, filecol, file, body)
        if not src_changes:
            ws.write(row, origsrccol, origsrc, body)
            ws.write(row, editsrccol, editsrc, body)
        else:
            text_to_write = print_changes(src_changes, True, red, body, lang)
            if len(text_to_write) > 3:
                ws.write_rich_string(row, origsrccol, *text_to_write)
            else:
                ws.write(row, origsrccol, text_to_write, red)
            text_to_write = print_changes(src_changes, False, red, body, lang)
            if len(text_to_write) > 3:
                ws.write_rich_string(row, editsrccol, *text_to_write)
            else:
                ws.write(row, editsrccol, text_to_write, red)
            src_equal_flag = False
            total_src_equal_flag = False
        if not tar_changes:
            ws.write(row, origtarcol, origtar, body)
            ws.write(row, edittarcol, edittar, body)
        else:
            text_to_write = print_changes(tar_changes, True, red, body, lang)
            if len(text_to_write) > 3:
                ws.write_rich_string(row, origtarcol, *text_to_write)
            else:
                ws.write(row, origtarcol, text_to_write, red)
            text_to_write = print_changes(tar_changes, False, red, body, lang)
            if len(text_to_write) > 3:
                ws.write_rich_string(row, edittarcol, *text_to_write)
            else:
                ws.write(row, edittarcol, text_to_write, red)
            tar_equal_flag = False
        ws.write(row, commentcol, "", body)
        if src_equal_flag and tar_equal_flag:
            ws.set_row(row, None, None, {'hidden': True})
        row += 1
    
    if total_src_equal_flag:
         ws.set_column('C:C', None, None, {'hidden': True})
        
    wb.close()
    return
   