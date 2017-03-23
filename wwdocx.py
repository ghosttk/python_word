import sys
import os
import zipfile
from zipfile import ZipFile
import shutil
from lxml import etree
import zip_dir


class docx():
    def add_elems(self,new_elem):
        for elem in self.elems:
            etree.SubElement(elem.getparent(),'{%s}%s'%(nsmap['wp'],new_elem）
        pass
    def remove_one_elem(self,old_elem):
        pass
    def remove_elems(self,elem_to_remove):
        self.find_elems(elem_to_remove)       
        for elem in self.elems:
            elem.getparent().remove(elem)
    def change_one_elem(self,old_elem,new_elem):
        self.add_elems(new_zfname)
        self.remove_elems(elem_to_remove)
        pass
    def __init__(self,fname):
        self.fname=fname
        self.extract_file()
        self.parse_document()
    def extract_file(self):
        f=ZipFile(self.fname)
        rtn=f.extractall(path='tmp')
        f.close()
        return rtn
    def parse_document(self):
        f=open('tmp/word/document.xml')
        tree=etree.parse(f)
        root=tree.getroot()
        self.tree=tree
        self.root=root
        self.nsmap=root.nsmap
        return root
    def find_elems(self,elem):
        elems=self.tree.xpath(elem,namespaces=self.nsmap)
        self.elems=elems
        return elems
    def save_docx(self):
        self.zip_dir('tmp')
        pass
    def zip_dir(self,dirname):
        new_zfname='%s_new.%s'%(self.fname.split('.')[0],'docx')
        filelist = []
        if os.path.isfile(dirname):
            filelist.append(dirname)
        else :
            for root, dirs, files in os.walk(dirname):
                for name in files:
                    filelist.append(os.path.join(root, name))
        zf =ZipFile(new_zfname, "w",)
        for tar in filelist:
            arcname = tar[len(dirname):]
            zf.write(tar,arcname)
        zf.close()
        #shutil.rmtree('tmp')


if __name__=='__main__':
    mydoc=docx('tmp.docx')
    mydoc.find_elems('//w:align')
    mydoc.save_docx()