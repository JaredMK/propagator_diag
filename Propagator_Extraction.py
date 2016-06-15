import os
import re
import openpyxl
import matplotlib.pyplot as plt
workbook = openpyxl.Workbook()  #creates openpyxl workbook

#TO BE CHANGED
'''folder containing the propagator files'''
propagatorFilesFolder='/Jared_Propa'

#columns for each variable in workbook
colFile='A'
colMolecule='B'
colFullPointGroup='C'
colLargestAbelianSubgroup='D'
colLargestConciseAbelianSubgroup='E'
colCharge='F'
colMultiplicity='G'
colBasis='H'
colOrbital='I'
colHF='J'
colOVGF_A='K'
colOVGF_A_ps='L'
colOVGF_B='M'
colOVGF_B_ps='N'
colOVGF_C='O'
colOVGF_C_ps='P'
colOVGF_Recommended='Q'
colOVGF_Recommended_ps='R'
colP3='S'
colP3_ps='T'
colP3_plus='U'
colP3_plus_ps='V'
colD2='W'
colD2_ps='X'
colOVGF_A_HF='Y'
colOVGF_B_HF='Z'
colOVGF_C_HF='AA'
colOVGF_Recommended_HF='AB'
colP3_HF='AC'
colP3_plus_HF='AD'
colD2_HF='AE'

'''path to this file'''
path=os.path.dirname(os.path.realpath(__file__))
pathorigin=path     #used to save workbook in this location
'''excel file name to open with path'''
#excelFilePathName='/propagatorFilesExcel.xlsx'
excelFilePathName='/Propagator_diag.xlsx'

'''color for graphs'''
primaryColor=(.22, .42, .69)     #blue
secondaryColor=(1.0, .84, 0)    #gold
tertiaryColor=(0, 1, 1) #mystery

row=1

worksheet=workbook.active
worksheet.title="PROPAGATOR"
#creates worksheet Propagator

#add headings to each column
worksheet[colFile+'1']='File'
worksheet[colMolecule+'1']='Molecule'
worksheet[colFullPointGroup+'1']='Full Point Group'
worksheet[colLargestAbelianSubgroup+'1']='Largest Abelian Subgroup'
worksheet[colLargestConciseAbelianSubgroup+'1']='Largest Concise Abelian Subgroup'
worksheet[colCharge+'1']='Charge'
worksheet[colMultiplicity+'1']='Multiplicity'
worksheet[colBasis+'1']='Basis'
worksheet[colOrbital+'1']='Orbital'
worksheet[colHF+'1']='HF'
worksheet[colOVGF_A+'1']='OVGF A'
worksheet[colOVGF_A_ps+'1']='OVGF A PS'
worksheet[colOVGF_B+'1']='OVGF B'
worksheet[colOVGF_B_ps+'1']='OVGF B PS'
worksheet[colOVGF_C+'1']='OVGF C'
worksheet[colOVGF_C_ps+'1']='OVGF C PS'
worksheet[colOVGF_Recommended+'1']='OVGF Recommended'
worksheet[colOVGF_Recommended_ps+'1']='OVGF Recommended PS'
worksheet[colP3+'1']='P3'
worksheet[colP3_ps+'1']='P3 PS'
worksheet[colP3_plus+'1']='P3+'
worksheet[colP3_plus_ps+'1']='P3+ PS'
worksheet[colD2+'1']='D2'
worksheet[colD2_ps+'1']='D2 PS'
worksheet[colOVGF_A_HF+'1']='OVGF A-HF'
worksheet[colOVGF_B_HF+'1']='OVGF B-HF'
worksheet[colOVGF_C_HF+'1']='OVGF C-HF'
worksheet[colOVGF_Recommended_HF+'1']='OVGF-Recommended-HF'
worksheet[colP3_HF+'1']='P3-HF'
worksheet[colP3_plus_HF+'1']='P3+-HF'
worksheet[colD2_HF+'1']='D2-HF'

row_dictionary={}   #row dictionary will contain keys with the form mol_aug_orb and values =[rows]
                    #ex) row_dictionary['CH3F_cc_6'] = {'pVTZ': 147, 'pVDZ': 141, 'pVQZ': 144}


#graph folders
ovgf_minus_hf_graphs='/OVGF-HF_graphs'
p3_minus_hf_graphs='/P3-HF_graphs'
p3p_minus_hf_graphs='/P3+-HF_graphs'
d2_minus_hf_graphs='/D2-HF_graphs'

def run():
    dataExtract(path)
    prepareGraphs()


def prepareGraphs():
    for molaugorb in row_dictionary:
        #find which row each basis is
        print(molaugorb)

        #get values in increasing basis set order
        ovgfA_m_hf_values=[]
        ovgfB_m_hf_values=[]
        ovgfC_m_hf_values=[]
        ovgfRec_m_hf_values=[]
        p3_m_hf_values=[]
        p3p_m_hf_values=[]
        d2_m_hf_values=[]

        xLabels=[]         #holds which basis sets from the ones below this mol_basis_orb has


        if 'pVDZ' in row_dictionary[molaugorb]:
            pvdz_row=row_dictionary[molaugorb]['pVDZ']
            xLabels.append('pVDZ')

            ovgfA_m_hf_values.append(worksheet[colOVGF_A_HF+str(pvdz_row)].value)
            ovgfB_m_hf_values.append(worksheet[colOVGF_B_HF+str(pvdz_row)].value)
            ovgfC_m_hf_values.append(worksheet[colOVGF_C_HF+str(pvdz_row)].value)
            ovgfRec_m_hf_values.append(worksheet[colOVGF_Recommended_HF+str(pvdz_row)].value)
            p3_m_hf_values.append(worksheet[colP3_HF+str(pvdz_row)].value)
            p3p_m_hf_values.append(worksheet[colP3_plus_HF+str(pvdz_row)].value)
            d2_m_hf_values.append(worksheet[colD2_HF+str(pvdz_row)].value)


        if 'pVTZ' in row_dictionary[molaugorb]:
            pvtz_row=row_dictionary[molaugorb]['pVTZ']
            xLabels.append('pVTZ')

            ovgfA_m_hf_values.append(worksheet[colOVGF_A_HF+str(pvtz_row)].value)
            ovgfB_m_hf_values.append(worksheet[colOVGF_B_HF+str(pvtz_row)].value)
            ovgfC_m_hf_values.append(worksheet[colOVGF_C_HF+str(pvtz_row)].value)
            ovgfRec_m_hf_values.append(worksheet[colOVGF_Recommended_HF+str(pvtz_row)].value)
            p3_m_hf_values.append(worksheet[colP3_HF+str(pvtz_row)].value)
            p3p_m_hf_values.append(worksheet[colP3_plus_HF+str(pvtz_row)].value)
            d2_m_hf_values.append(worksheet[colD2_HF+str(pvtz_row)].value)

        if 'pVQZ' in row_dictionary[molaugorb]:
            pvqz_row=row_dictionary[molaugorb]['pVQZ']
            xLabels.append('pVQZ')

            ovgfA_m_hf_values.append(worksheet[colOVGF_A_HF+str(pvqz_row)].value)
            ovgfB_m_hf_values.append(worksheet[colOVGF_B_HF+str(pvqz_row)].value)
            ovgfC_m_hf_values.append(worksheet[colOVGF_C_HF+str(pvqz_row)].value)
            ovgfRec_m_hf_values.append(worksheet[colOVGF_Recommended_HF+str(pvqz_row)].value)
            p3_m_hf_values.append(worksheet[colP3_HF+str(pvqz_row)].value)
            p3p_m_hf_values.append(worksheet[colP3_plus_HF+str(pvqz_row)].value)
            d2_m_hf_values.append(worksheet[colD2_HF+str(pvqz_row)].value)


        if 'pV5Z' in row_dictionary[molaugorb]:
            pv5z_row=row_dictionary[molaugorb]['pV5Z']
            xLabels.append('pV5Z')

            ovgfA_m_hf_values.append(worksheet[colOVGF_A_HF+str(pv5z_row)].value)
            ovgfB_m_hf_values.append(worksheet[colOVGF_B_HF+str(pv5z_row)].value)
            ovgfC_m_hf_values.append(worksheet[colOVGF_C_HF+str(pv5z_row)].value)
            ovgfRec_m_hf_values.append(worksheet[colOVGF_Recommended_HF+str(pv5z_row)].value)
            p3_m_hf_values.append(worksheet[colP3_HF+str(pv5z_row)].value)
            p3p_m_hf_values.append(worksheet[colP3_plus_HF+str(pv5z_row)].value)
            d2_m_hf_values.append(worksheet[colD2_HF+str(pv5z_row)].value)

        l=[]
        q=0
        for w in xLabels:
            l.append(q)
            q+=5

        graph(l,p3_m_hf_values,molaugorb,xLabels, p3_minus_hf_graphs)
        graph(l,p3p_m_hf_values,molaugorb,xLabels, p3p_minus_hf_graphs)
        graph(l,d2_m_hf_values,molaugorb,xLabels, d2_minus_hf_graphs)


        labelA=molaugorb+' OVGF A'
        labelB=molaugorb+' OVGF B'
        labelC=molaugorb+' OVGF C'
        if ovgfRec_m_hf_values==ovgfA_m_hf_values:
            labelA=molaugorb+' OVGF A, Recommended'
        elif ovgfRec_m_hf_values==ovgfB_m_hf_values:
            labelB=molaugorb+' OVGF B, Recommended'
        elif ovgfRec_m_hf_values==ovgfC_m_hf_values:
            labelC=molaugorb+' OVGF C, Recommended'

        graphOVGF(l, ovgfA_m_hf_values,ovgfB_m_hf_values,ovgfC_m_hf_values, molaugorb, xLabels, ovgf_minus_hf_graphs, labelA,labelB,labelC)


def graph(l, values, moleculeLabel, xLabels, graphFolder):

        plt.plot(l, values, color=primaryColor, lw=2, ls='-', marker='s', label=moleculeLabel)
        plt.xticks(l, xLabels, rotation = '30', ha='right')
        plt.margins(0.09, 0.09)
        plt.subplots_adjust(bottom=0.2, top=0.85)
        plt.ylabel('values')
        plt.legend(loc='upper center', bbox_to_anchor=(.5, 1.2), numpoints = 1, shadow=True, ncol=3)
        plt.grid(True)
        if not os.path.exists(path + '/ALL GRAPHS' +graphFolder):
            os.makedirs(path + '/ALL GRAPHS' +graphFolder)
        plt.savefig(path + '/ALL GRAPHS' +graphFolder+ '/' + str(moleculeLabel) + '.eps')
        plt.close()


def graphOVGF(l, valuesA, valuesB, valuesC, moleculeLabel,xLabels, graphFolder, labelA,labelB,labelC):

    plt.plot(l, valuesA, color=primaryColor, lw=2, ls='-', marker='s', label=labelA)
    plt.plot(l, valuesB, color=secondaryColor, lw=2, ls='-', marker='o', label=labelB)
    plt.plot(l, valuesC, color=tertiaryColor, lw=2, ls='-', marker='<', label=labelC)

    plt.xticks(l, xLabels, rotation = '30', ha='right')
    plt.margins(0.09, 0.09)
    plt.subplots_adjust(bottom=0.2, top=0.85)
    plt.ylabel('values')
    plt.legend(loc='upper center', bbox_to_anchor=(.5, 1.2), numpoints = 1, shadow=True, ncol=3)
    plt.grid(True)
    if not os.path.exists(path + '/ALL GRAPHS' +graphFolder):
        os.makedirs(path + '/ALL GRAPHS' +graphFolder)
    plt.savefig(path + '/ALL GRAPHS' +graphFolder+ '/' + str(moleculeLabel) + '.eps')
    plt.close()


def writeDataToExcel(worksheet, row, fileInformation,orbital,hf,ovgf_a, ovgf_a_ps, ovgf_b,\
            ovgf_b_ps,ovgf_c, ovgf_c_ps,ovgf_recommend,ovgf_recommended_ps,\
            p3,p3_plus,d2, ovgf_a_hf,ovgf_b_hf,ovgf_c_hf,\
            ovgf_recommend_hf,p3_hf,p3_plus_hf,d2_hf,molecule,charge,multiplicity,\
            basis,fullPointGroup,largestAbelianSubgroup,largestConciseAbelianSubgroup,p3_ps,p3_plus_ps,d2_ps):
    '''writesDataToExcel takes is called by dataExtract. It takes in the variables found in
    data extraction and writes it into the openpyxl workbook'''

    worksheet[colFile+str(row)]=fileInformation
    worksheet[colOrbital+str(row)]=orbital
    worksheet[colHF+str(row)]=hf
    worksheet[colOVGF_A+str(row)]=ovgf_a
    worksheet[colOVGF_B+str(row)]=ovgf_b
    worksheet[colOVGF_C+str(row)]=ovgf_c
    worksheet[colOVGF_A_ps+str(row)]=ovgf_a_ps
    worksheet[colOVGF_B_ps+str(row)]=ovgf_b_ps
    worksheet[colOVGF_C_ps+str(row)]=ovgf_c_ps
    worksheet[colOVGF_Recommended+str(row)]=ovgf_recommend
    worksheet[colOVGF_Recommended_ps+str(row)]=ovgf_recommended_ps
    worksheet[colP3+str(row)]=p3
    worksheet[colP3_ps+str(row)]=p3_ps
    worksheet[colP3_plus+str(row)]=p3_plus
    worksheet[colP3_plus_ps+str(row)]=p3_plus_ps
    worksheet[colD2+str(row)]=d2
    worksheet[colD2_ps+str(row)]=d2_ps
    worksheet[colOVGF_A_HF+str(row)]=ovgf_a_hf
    worksheet[colOVGF_B_HF+str(row)]=ovgf_b_hf
    worksheet[colOVGF_C_HF+str(row)]=ovgf_c_hf
    worksheet[colOVGF_Recommended_HF+str(row)]=ovgf_recommend_hf
    worksheet[colP3_HF+str(row)]=p3_hf
    worksheet[colP3_plus_HF+str(row)]=p3_plus_hf
    worksheet[colD2_HF+str(row)]=d2_hf
    worksheet[colMolecule+str(row)]=molecule
    worksheet[colCharge+str(row)]=charge
    worksheet[colMultiplicity+str(row)]=multiplicity
    worksheet[colBasis+str(row)]=basis
    worksheet[colFullPointGroup+str(row)]=fullPointGroup
    worksheet[colLargestAbelianSubgroup+str(row)]=largestAbelianSubgroup
    worksheet[colLargestConciseAbelianSubgroup+str(row)]=largestConciseAbelianSubgroup

def numberOfBasisSets(logarray):
    '''returns a list of the split log arrays by basis set. length is number of basis sets'''
    commandLocation=[]
    logsToReturn=[]
    x=0
    while x < len(logarray):
        if logarray[x] =='corrections':
            commandLocation.append(x)
        x+=1
    commandLocation.append(len(logarray))
    x=0
    logsToReturn.append(logarray[:commandLocation[0]])
    #the first log in the array is from the start of the file to the first keyword
    while x< len(commandLocation)-1:
        b=logarray[commandLocation[x]:commandLocation[x+1]]
        logsToReturn.append(b)
        x+=1
    return logsToReturn

def dataExtract(path):
    '''Main function in script. Calls other functions. Takes in path of this file and extracts
    data from the log files folder. Then calls functions above to add data to the openpyxl file'''

    row=2

    #data extraction from log files
    logFiles=[]

    for path, subdirs, files in os.walk(path+propagatorFilesFolder):
        for name in files:
            if os.path.join(path, name)[len(os.path.join(path, name))-\
            4:len(os.path.join(path, name))]=='.log':
                logFiles.append(os.path.join(path, name))

    for currentFile in logFiles:
        log = open(currentFile, 'r').read()
        splitLog = re.split(r'[\\\s]\s*', log)

        #fileInformation needs to be added to excel
        fileInformation=currentFile

        firstSplitLog=numberOfBasisSets(splitLog)[0]
        x=0
        while x<len(firstSplitLog):
            if firstSplitLog[x]=='Stoichiometry':
                molecule=firstSplitLog[x+1]
            if firstSplitLog[x]=='Charge':
                charge=firstSplitLog[x+2]
            if firstSplitLog[x]=='Multiplicity':
                multiplicity=firstSplitLog[x+2]
            if firstSplitLog[x]=='Standard' and firstSplitLog[x+1]=='basis:':
                basis=firstSplitLog[x+2]
            if firstSplitLog[x]=='Full':
                fullPointGroup=firstSplitLog[x+3]
            if firstSplitLog[x]=='Largest' and firstSplitLog[x+1]=='Abelian':
                largestAbelianSubgroup=firstSplitLog[x+3]
            if firstSplitLog[x]=='concise':
                largestConciseAbelianSubgroup=firstSplitLog[x+3]

            x+=1

        for splitlog in numberOfBasisSets(splitLog)[1:]:
            #NUMBEROFSPLITS will return where in log file it needs to be split for basis sets
            #textFile(log)   #text file will return each log split by basis set because some aren't
            #print(splitlog)

            x=0
            while x<len(splitlog):

                if splitlog[x]=='Method' and splitlog[x+1]=='Orbital':

                    orbital=splitlog[x+6]
                    hf=float(splitlog[x+7])
                    ovgf_a=float(splitlog[x+8])
                    ovgf_a_ps=float(splitlog[x+9])
                    ovgf_b=float(splitlog[x+13])
                    ovgf_b_ps=float(splitlog[x+14])
                    ovgf_c=float(splitlog[x+18])
                    ovgf_c_ps=float(splitlog[x+19])

                if splitlog[x]=='recommended':
                    ovgf_recommend=float(splitlog[x+8])
                    ovgf_recommended_ps=float(splitlog[x+9])


                if splitlog[x]=='Converged' and splitlog[x+1]=='3rd' and splitlog[x+2]=='order':
                    p3=float(splitlog[x+7])
                    p3_ps=float(splitlog[x+9])

                    try:
                        p3_plus=float(splitlog[x+17])
                        p3_plus_ps=float(splitlog[x+19])
                    except:
                        p3_plus=None
                        p3_plus_ps=None

                if splitlog[x]=='Converged' and splitlog[x+1]=='second':
                    d2=float(splitlog[x+6])
                    d2_ps=float(splitlog[x+8])

                if x==len(splitlog)-1:
                    ovgf_a_hf = ovgf_a-hf
                    ovgf_b_hf=ovgf_b-hf
                    ovgf_c_hf=ovgf_c-hf
                    ovgf_recommend_hf=ovgf_recommend-hf

                    p3_hf=p3-hf
                    if p3_plus!=None:
                        p3_plus_hf=p3_plus-hf
                    else:
                        p3_plus_hf=None
                    d2_hf=d2-hf

                x+=1
            #data stored in variables is input to writesDataToExcel
            writeDataToExcel(worksheet, row, fileInformation,orbital,hf,ovgf_a, ovgf_a_ps, ovgf_b,\
            ovgf_b_ps,ovgf_c, ovgf_c_ps,ovgf_recommend,ovgf_recommended_ps,\
            p3,p3_plus,d2, ovgf_a_hf,ovgf_b_hf,ovgf_c_hf,\
            ovgf_recommend_hf,p3_hf,p3_plus_hf,d2_hf,molecule,charge,multiplicity,basis,\
            fullPointGroup,largestAbelianSubgroup,largestConciseAbelianSubgroup,p3_ps,p3_plus_ps,d2_ps)


            if basis[0]=='A':
                aug_dict='aug'
            elif basis[0]=='C':
                aug_dict='cc'


            dictionaryKey = molecule+'_'+aug_dict+'_'+str(orbital)      #key for molecule,aug,orbital
            basisDictionary={basis[len(basis)-4:len(basis)]:row}                                 #def is dictionary with basis:row
            if dictionaryKey not in row_dictionary:
                row_dictionary[dictionaryKey]=basisDictionary
            elif dictionaryKey in row_dictionary:
                row_dictionary[dictionaryKey][basis[len(basis)-4:len(basis)]]=row



            row+=1

    workbook.save(pathorigin + excelFilePathName)     #saves file
    print(row_dictionary)
