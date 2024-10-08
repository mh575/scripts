import os
import pymupdf
print(pymupdf.__doc__)
# general idea: if match in common field title: check line and line under, maybe regex between field title x and next field title y? 

'''
TO PULL:
chemical name, synonyms, forumula, flash point, density, melting point, boiling point, molecular weight, pH value, physical state, compatability category(?), UN number, UN pack group, TDG primary, TDG secondary, TDG teritary, GHS classification, GHS hazard statements (pictogram(s) derived from h-codes), GHS signal word, GHS precautionary statement codes, NFPA Diamond, DOT Guide(?)
'''

def main():
    '''iterate over files in sds folder, parse and generate a csv file where each row is a sds file data'''
    assert os.path.exists(os.getcwd()+'/sds_files')
    
    dir = os.fsencode(os.getcwd()+'/sds_files')
    
    for file in os.listdir(dir):
         filename = os.fsdecode(file)
         if filename.endswith('.pdf'):
             doc = pymupdf.open('sds_files/'+filename)
             doc_str = ''
             page = doc.load_page(0)
             for page in doc:
                 print(page.get_text())
                 doc_str += page.get_text()
             print(doc_str)
         else: # file is not a pdf
            if not (filename == '.gitignore'):
                raise ValueError('file extension is not pdf: '+filename)
            else:
                continue

if __name__ == "__main__":
    '''This is executed when run from the command line'''
    main()