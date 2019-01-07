from docx2txt.docx2txt import *


def numbering(num, lvl, NUM):
    '''
    определение метода нумерации списка и начального элемента, с которого начинается нумерация
    '''
    word_namespace = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    wnum = word_namespace + 'num'
    wabstractNumId = word_namespace + 'abstractNumId'
    wabstractNum = word_namespace + 'abstractNum'
    value = word_namespace + 'val'
    numId = word_namespace + 'numId'
    wlvl = word_namespace + 'lvl'
    ilvl = word_namespace + 'ilvl'
    wstart = word_namespace + 'start'
    wnumFmt = word_namespace + 'numFmt'
    wlvlText = word_namespace + 'lvlText'
    abstractNumId = 'None'
    numer = None
    for num1 in NUM.getiterator(wnum):
        if num1.get(numId) == num:
            abstractNumId = num1.find(wabstractNumId).get(value)
            break
    for abstract in NUM.getiterator(wabstractNum):
        if abstract.get(wabstractNumId) == abstractNumId:
            for tagwlvl in abstract.getiterator(wlvl):
                if tagwlvl.get(ilvl) == lvl:
                    numer = (tagwlvl.find(wnumFmt).get(value), tagwlvl.find(wstart).get(value),
                             tagwlvl.find(wlvlText).get(value))
                    break
            break
    if not numer:
        numer = (None, None, None)
    return numer


def process(docx, img_dir=None):
    text = u''

    # unzip the docx in memory
    # zipf = zipfile.ZipFile(docx)
    # filelist = zipf.namelist()

    # get header text
    # there can be 3 header files in the zip
    # header_xmls = 'word/header[0-9]*.xml'
    # for fname in filelist:
    #     if re.match(header_xmls, fname):
    #         text += xml2text(zipf.read(fname))

    # get main text
    # doc_xml = 'word/document.xml'
    # text += xml2text(zipf.read(doc_xml))
    text += xml2text(docx)

    # get footer text
    # there can be 3 footer files in the zip
    # footer_xmls = 'word/footer[0-9]*.xml'
    # for fname in filelist:
    #     if re.match(footer_xmls, fname):
    #         text += xml2text(zipf.read(fname))

    # if img_dir is not None:
    #     # extract images
    #     for fname in filelist:
    #         _, extension = os.path.splitext(fname)
    #         if extension in [".jpg", ".jpeg", ".png", ".bmp"]:
    #             dst_fname = os.path.join(img_dir, os.path.basename(fname))
    #             with open(dst_fname, "w") as dst_f:
    #                 dst_f.write(zipf.read(fname))
    #
    # zipf.close()
    # return text.strip()
    return text


def xml2text(path):
    zipf = zipfile.ZipFile(path)
    number_xml = 'word/numbering.xml'
    doc_xml = 'word/document.xml'
    try:
        document_numbers = zipf.read(number_xml)
        number_xml = ET.XML(document_numbers)
        nmbr = 'not none'
    except Exception:
        nmbr = None
    document_text = zipf.read(doc_xml)
    zipf.close()
    text = u''
    root = ET.fromstring(document_text)
    j = 0
    lowletter = 1
    previous_list = 1
    lowerRoman = 1
    for n, child in enumerate(root.iter()):
        if child.tag == qn('w:numPr'):
            lvl = child.find(qn('w:ilvl')).get(qn('w:val'))
            num = child.find(qn('w:numId')).get(qn('w:val'))
            if nmbr != None:
                nmbr = numbering(num, lvl, number_xml)
                current_list = int(num)
                # print(num, lvl, nmbr, previous_list, current_list, j)
                if nmbr[0] == 'bullet' or nmbr[0] == 'none':
                    text += "\t" * int(lvl) + str(nmbr[2]) + " "
                    # pass
                elif nmbr[0] == 'decimal':
                    # Детектим первый элемент в новом списке
                    if previous_list < current_list:
                        j = 1
                        previous_list = current_list
                        text += "\t" * int(lvl) \
                                + str(j) + ". "
                        lowletter = 1
                    elif previous_list == current_list and int(lvl) == 0:
                        j += + 1
                        text += "\t" * int(lvl) \
                                + str(j) + ". "
                elif nmbr[0] == 'lowerLetter':
                    if previous_list == current_list and int(lvl) != 0:
                        text += "\t" * int(lvl) + chr(int(nmbr[1]) + lowletter - 2 + ord('a')) + ". "
                        lowletter += 1
                elif nmbr[0] is None:
                    pass
            else:
                pass
        elif child.tag == qn('w:t'):
            t_text = child.text
            text += t_text if t_text is not None else ''
        elif child.tag == qn('w:tab'):
            text += '\t'
        elif child.tag in (qn('w:br'), qn('w:cr')):
            text += '\n'
        elif child.tag == qn('w:p'):
            text += '\n'
        text = ''.join(text)
    return text


xml2text(docx_path)
