import os
import sys
import re
import zipfile
import contextlib
import shutil

def zipfunc(fld2zip,fileName=None):
    shutil.make_archive(os.path.join(os.path.dirname(fld2zip),os.path.basename(fld2zip)), 'zip', fld2zip)
    if fileName == None:
        shutil.move(os.path.join(os.path.dirname(fld2zip),os.path.basename(fld2zip)+".zip"),os.path.join(os.path.dirname(fld2zip),"Temp_"+os.path.basename(fld2zip)+".docx"))
    else:
        shutil.move(os.path.join(os.path.dirname(fld2zip), os.path.basename(fld2zip) + ".zip"),
                    os.path.join(os.path.dirname(fld2zip), os.path.basename(fld2zip) + ".docx"))
    shutil.rmtree(fld2zip)
    return 0


def unzip(zip2ext):
    with contextlib.closing(zipfile.ZipFile(zip2ext, 'r')) as z:
        z.extractall(os.path.join(os.path.dirname(zip2ext),(os.path.basename(zip2ext).upper()).replace(".DOCX","")))

    return os.path.join(os.path.dirname(zip2ext),(os.path.basename(zip2ext).upper()).replace(".DOCX",""))

def extract(myFile):
    global fileCont
    fldPath = unzip(myFile)
    fo = open(myFile+".log","w")
    fo.close()

    xmlFile = os.path.join(fldPath,'word/document.xml')
    with open(xmlFile,"r+") as fo:
        fileCont = fo.read()

    fileCont=re.sub(r'(<w:del [^>]+>)<w:r( [^>]+>((?:(?!</w:r>).)*)</w:r>)','\g<1><w:changed_r\g<2>',fileCont,re.DOTALL)

    for mt in re.finditer('<w:t(?: [^>]+)?>((?:(?!</w:t>).)*)</w:t>',fileCont):
        fullLine=mt.group()
        bk_fullLine = fullLine
        grpStr = mt.group(1)

        if grpStr == " ":
            continue
        if not re.search('\s',grpStr):
            continue

        bk_grpStr = grpStr
        grpStr=re.sub('\s','mergespace',grpStr)
        fullLine=re.sub(re.escape(bk_grpStr),grpStr,fullLine)
        fileCont=fileCont.replace(bk_fullLine,fullLine)



    for mt in re.finditer('<w:pPr><w:pStyle w:val="[^"]+"/>((?:(?!</w:pPr>).)*)</w:pPr><w:r[^>]*>((?:(?!<w:t[^>]*>).)*)<w:t[^>]*>', fileCont):
        mt1 = re.search('(<w:([^>]+) w:val="(?:0|baseline|none)"/>)', str(mt.group(1)))
        if mt1 and str(mt1.group(2)).upper() != "WIDOWCONTROL":
            fileCont = re.sub(str(mt.group()), str(mt.group()) + "mergedocx_start"+(str(mt1.group()).replace("<",'&lt;')).replace('>',"&gt;")+"mergedocx_end", fileCont)


    for mt in re.finditer('(<w:rPr><w:rStyle w:val="[^"]+"/>)((?:(?!</w:rPr>).)*)(</w:rPr><w:t[^>]*>)',fileCont):
        for mt1 in re.finditer('<w:([^>]+) w:val="(?:0|baseline|none)"/>', str(mt.group(2))):
            if mt1 and str(mt1.group(1)).upper() != "WIDOWCONTROL":
                fileCont=re.sub(str(mt.group()),str(mt.group(1))+str(mt.group(3))+"mergedocx_start"+(str(mt1.group()).replace("<",'&lt;')).replace('>',"&gt;")+"mergedocx_end",fileCont,count=1,flags=re.DOTALL)

    fileCont=re.sub(r'(<[^>]+)mergespace([^>]*>)','\g<1> \g<2>',fileCont)
    fileCont=re.sub(r'<w:changed_r([^>]+)>','<w:r\g<1>>',fileCont)
    fo = open(xmlFile,"w")
    fo.write(fileCont)
    fo.close()

    zipfunc(fldPath)

    os.remove(myFile+".log")

    return 0

def final(myFile,Extra=False):
    global fileCont
    fo=open(myFile+".log","w")
    fo.close()
    fldPath = unzip(myFile)
    xmlFile = os.path.join(fldPath, 'word/document.xml')
    with open(xmlFile, "r+") as fo:
        fileCont = fo.read()

    fileCont = re.sub(r'(<w:del [^>]+>)<w:r( [^>]+>((?:(?!</w:r>).)*)</w:r>)', '\g<1><w:changed_r\g<2>', fileCont,re.DOTALL)
    fileCont=re.sub('(mergedocx_start&lt;w:[^ ]+ )</w:t>(?:(?:(?!<w:t[^>]*>).)*)<w:t[^>]*>(w:val="[^"]+"/&gt;mergedocx_end)','\g<1>\g<2>',fileCont)

    while fileCont.count("mergedocx_start") != 0:
        # </w:rPr><w:commentReference w:id="223"/></w:r><w:r><w:t xml:space="preserve
        for mt in re.finditer('(</w:rPr>(?:<w:[^>/<]+/>)?(?:</w:r><w:r ?[^>]*>)?<w:t[^>]*>(?:(?:(?!</w:t>).)*))mergedocx_start((?:(?!mergedocx_end).)*)mergedocx_end',fileCont):
            fileCont=re.sub(re.escape(mt.group()),(str(mt.group(2).replace("&lt;","<").replace("&gt;",">")))+mt.group(1),fileCont,count=1)


        # </w:pPr><w:r w:rsidRPr="006F3A49"><w:t>
        # </w:pPr><w:r><w:sym w:font="Wingdings" w:char="F0D8"/></w:r><w:r><w:t
        for mt in re.finditer('(</w:pPr>(?:<w:r ?[^>]*><w:[^>]+></w:r>)?(?:<w:[^>/<]+/>)?<w:r ?[^>]*>(?:<w:lastRenderedPageBreak/>)?<w:t[^>]*>(?:(?:(?!</w:t>).)*))mergedocx_start((?:(?!mergedocx_end).)*)mergedocx_end',fileCont):
            fileCont = re.sub(re.escape(mt.group()),'<w:rPr>'+(str(mt.group(2).replace("&lt;", "<").replace("&gt;", ">"))) + '</w:rPr>'+mt.group(1), fileCont, count=1)



    fileCont = re.sub(r'<w:changed_r([^>]+)>', '<w:r\g<1>>', fileCont)

    if Extra==True:
        for mt_extra in re.finditer(r'<w:p [^>]*>((?:(?!</w:p>).)*)</w:p>', fileCont, flags=re.DOTALL):
            txtGrp = str(mt_extra.group(0))
            bktxtGrp = txtGrp
            mt1_extra = re.findall(
                '<w:r( [^>]*)?><w:rPr><w:rFonts[^>]+w:ascii="Symbol"[^>]*/></w:rPr><w:t>((?:(?!</w:r>).)*)</w:r>',
                txtGrp, flags=re.DOTALL)
            if len(mt1_extra) > 1:
                txtGrp = re.sub(
                    '<w:r( [^>]*)?><w:rPr><w:rFonts[^>]+w:ascii="Symbol"[^>]*/></w:rPr><w:t>((?:(?!</w:r>).)*)</w:r>',
                    '', txtGrp, count=len(mt1_extra) - 1)
                fileCont = fileCont.replace(bktxtGrp, txtGrp)

    fo = open(xmlFile, "w")
    fo.write(fileCont)
    fo.close()
    zipfunc(fldPath, "FinalFunc")

    os.remove(myFile + ".log")

    return 0


IpFile = sys.argv[1]
function2Exec = sys.argv[2]
if function2Exec=="extract":
    extract(IpFile)
if function2Exec=="final":
    final(IpFile,Extra=False)
if function2Exec=="final_extra":
    final(IpFile,Extra=True)
