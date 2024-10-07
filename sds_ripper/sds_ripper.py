import os
import pymupdf
print(pymupdf.__doc__)
# general idea: if match in common field title: check line and line under, maybe regex between field title x and next field title y? 

'''
TO PULL:
chemical name, synonyms, chemical forumula, flash point, density weight, melting point, boiling point, molecular weight, pH value, unit, physical state, compatability category, special hazard, UN number, UN pack group, TDG primary, TDG secondary, TDG teritary, storage requirements, GHS classification, GHS hazard statements, GHS signal word, GHS precautionary statement codes, GHS pictogram(s), NFPA Diamond, DOT Guide
'''

def main():
    '''iterate over files in sds folder, parse and generate a csv file where each row is a sds file data'''
    # for sds in 

if __name__ == "__main__":
    '''This is executed when run from the command line'''
    main()