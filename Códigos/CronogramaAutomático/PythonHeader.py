#Referência para nomes de colunas planilha MEDIÇÕES:
#REFERÊNCIA - CONTRATO - ITEM - QUANTIDADE - MÊS REFERÊNCIA
#IMPORTANTE - o ITEM vem como str e no formato 1.1, 1.2, 2.1, 2.2, como na planilha
#___________________________________________________________________________________________________________________________________________________________________________________________________________
#_______________________________CRIANDO CONSTANTES_________________________________________________________________________________________________________________________________________________________#
#__________________________________________________________________________________________________________________________________________________________________________________________________________#
refPl=xl("$C$5") #pega a cél de referência da planilha (uma divisão ou gerência)(pra achar ela na planilha de medição)
ncPl=xl("$D$5") #pega a cél Nº Contrato da planilha (mesma coisa da de cima)
fullMed=pd.DataFrame(xl("_xlfn.PQSOURCE(\"5fb4d1bb-17a2-47b5-bc6a-40d2e78c149c\")")) #Pega a Planilha medições (da query) (lembrando q são só aquelas colunas lá em cima e SÓ AS MEDIÇÕES APROVADAS)
MedFiltered=fullMed[(fullMed['REFERÊNCIA'] == refPl) & (fullMed['CONTRATO'] == ncPl)] #Filtra a medições a partir da bacia e do nº do contrato
mesesmedidos=len(MedFiltered.groupby('MÊS REFERÊNCIA').nunique())#conta quantos meses foram medidos
mespulado=int() #inicia o mês pulado pra não trocar de valor enquanto roda
#__________________________________________________________________________________________________________________________________________________________________________________________________________#
#______________#SE CONTRATO COMEÇA NO DIA 15 COLOQUE True
#______________#SE ELE COMEÇAR NO DIA 1 COLOQUE False
meiodomes=False #<------------------------------AQUI
#__________________________________________________________________________________________________________________________________________________________________________________________________________#
""
