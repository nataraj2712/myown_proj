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
    fileCont=re.sub(r'<w:t xml:space="preserve">','<w:t>mergepreservespace',fileCont)

    for mt in re.finditer('<w:pPr><w:pStyle w:val="[^"]+"/>((?:(?!</w:pPr>).)*)</w:pPr><w:r[^>]*>((?:(?!<w:t[^>]*>).)*)<w:t[^>]*>', fileCont):
        mt1 = re.search('(<w:([^>]+) w:val="(?:0|baseline|none)"/>)', str(mt.group(1)))
        if mt1 and str(mt1.group(2)).upper() != "WIDOWCONTROL":
            fileCont = re.sub(str(mt.group()), str(mt.group()) + "mergedocx_start"+(str(mt1.group()).replace("<",'&lt;')).replace('>',"&gt;")+"mergedocx_end", fileCont)


    for mt in re.finditer('(<w:rPr><w:rStyle w:val="[^"]+"/>)((?:(?!</w:rPr>).)*)(</w:rPr><w:t[^>]*>)',fileCont):
        for mt1 in re.finditer('<w:([^>]+) w:val="(?:0|baseline|none)"/>', str(mt.group(2))):
            if mt1 and str(mt1.group(1)).upper() != "WIDOWCONTROL":
                fileCont=re.sub(str(mt.group()),str(mt.group(1))+str(mt.group(3))+"mergedocx_start"+(str(mt1.group()).replace("<",'&lt;')).replace('>',"&gt;")+"mergedocx_end",fileCont,count=1,flags=re.DOTALL)



    fileCont=re.sub(r'<w:changed_r([^>]+)>','<w:r\g<1>>',fileCont)
    fo = open(xmlFile,"w")
    fo.write(fileCont)
    fo.close()

    zipfunc(fldPath)

    os.remove(myFile+".log")

    return 0

def final(myFile):
    global fileCont
    fo=open(myFile+".log","w")
    fo.close()
    fldPath = unzip(myFile)
    xmlFile = os.path.join(fldPath, 'word/document.xml')
    with open(xmlFile, "r+") as fo:
        fileCont = fo.read()

    fileCont = re.sub(r'(<w:del [^>]+>)<w:r( [^>]+>((?:(?!</w:r>).)*)</w:r>)', '\g<1><w:changed_r\g<2>', fileCont,re.DOTALL)
    fileCont = re.sub(r'<w:t>mergepreservespace', '<w:t xml:space="preserve">', fileCont)
    fileCont = re.sub(r'mergepreservespace', '', fileCont)
    # fileCont = re.sub('</w:t></w:r><w:r><w:rPr><w:rStyle w:val="[^>]+"/></w:rPr><w:lastRenderedPageBreak/><w:t>w:','w:',fileCont)
    fileCont=re.sub('(mergedocx_start&lt;w:[^ ]+ )</w:t>(?:(?:(?!<w:t[^>]*>).)*)<w:t[^>]*>(w:val="[^"]+"/&gt;mergedocx_end)','\g<1>\g<2>',fileCont)

    cnt = 1
    while fileCont.count("mergedocx_start") != 0:
        # </w:rPr><w:commentReference w:id="223"/></w:r><w:r><w:t xml:space="preserve
        for mt in re.finditer('(</w:rPr>(?:<w:[^>/<]+/>)?(?:</w:r><w:r ?[^>]*>)?<w:t[^>]*>(?:(?:(?!</w:t>).)*))mergedocx_start((?:(?!mergedocx_end).)*)mergedocx_end',fileCont):
            fileCont=re.sub(re.escape(mt.group()),(str(mt.group(2).replace("&lt;","<").replace("&gt;",">")))+mt.group(1),fileCont,count=1)
            # cnt = cnt + 1
            # print "loop1--->" + str(cnt)
            # fo = open("D:\\121212.xml", "w")
            # fo.write(fileCont)
            # fo.close()

        # </w:pPr><w:r w:rsidRPr="006F3A49"><w:t>
        # </w:pPr><w:r><w:sym w:font="Wingdings" w:char="F0D8"/></w:r><w:r><w:t
        for mt in re.finditer('(</w:pPr>(?:<w:r ?[^>]*><w:[^>]+></w:r>)?(?:<w:[^>/<]+/>)?<w:r ?[^>]*>(?:<w:lastRenderedPageBreak/>)?<w:t[^>]*>(?:(?:(?!</w:t>).)*))mergedocx_start((?:(?!mergedocx_end).)*)mergedocx_end',fileCont):
            fileCont = re.sub(re.escape(mt.group()),'<w:rPr>'+(str(mt.group(2).replace("&lt;", "<").replace("&gt;", ">"))) + '</w:rPr>'+mt.group(1), fileCont, count=1)
            # cnt = cnt + 1
            # print "loop3--->" + str(cnt)
            # fo = open("D:\\121212.xml", "w")
            # fo.write(fileCont)
            # fo.close()
        # fo = open("D:\\121212.xml", "w")
        # fo.write(fileCont)
        # fo.close()


    fileCont = re.sub(r'<w:changed_r([^>]+)>', '<w:r\g<1>>', fileCont)
    # fileCont = re.sub(r'/&gt;&lt;w:', '/><w:', fileCont)
    # fileCont = re.sub(r'(<w:r [^>]+>(?:<w:lastRenderedPageBreak/>)?)(<w:t[^>]*>)mergedocx_start&lt;((?:(?!&gt;mergedocx_end).)*)&gt;mergedocx_end',
    #     '\g<1><w:rPr><\g<3>></w:rPr>\g<2>', fileCont)
    fo = open(xmlFile, "w")
    fo.write(fileCont)
    fo.close()
    zipfunc(fldPath, "FinalFunc")

    os.remove(myFile + ".log")

    return 0


IpFile = sys.argv[1]
function2Exec = sys.argv[2]
if function2Exec == "extract":
    extract(IpFile)
if function2Exec == "final":
    final(IpFile)