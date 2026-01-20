#___________________________________________________________________________VARIÁVEIS POR CÉLULA___________________________________________________________________________________________________________#
#__________________________________________________________________________________________________________________________________________________________________________________________________________#
mes=int(xl("C$11")) #pega o mês ref do cálculo                                                            #____________ESSES CARAS PRECISAM PUXAR VALORES                                                       
item=xl("$A13") #pega o item ref (usado para filtrar planilha de medição)                                 #____________ESTÁTICOS NO EXCEL E PRECISAM DAS REFERÊNCIAS
datames=xl("C$12") #pega o mes (em data) do valor (usado para filtrar planilha de medição)                #____________ENTÃO NÃO MUDA A FORMATAÇÃO SENÃO DA ERRO
precouni=xl("'Planilha '!$E13") #pega preço unitário ref ao quantitativo                                  #____________E COM ELES ASSIM VOCE PODE DAR COPIA COLA EM
quantitativomes=0.0 #usado para os calculos                                                               #____________QUALQUER LUGAR DA PLANILHA
quantitativomed=MedFiltered[(MedFiltered['ITEM']==item)]['QUANTIDADE'].sum() #pega a soma dos quantitativos medidos para esse item
totalmeses=max(xl("$C$11:$N$11", headers=True))-mesesmedidos#pega o valor máx dentro do nº de meses (max 12 meses) e subtrai o número de meses que foram medidos
totalquant=float(xl("'Planilha '!$D13"))-float(quantitativomed) #pega o total de quantitativos da outra planilha7
if mes ==1:
    mespulado=0 #reinicia a variavel mespulado para nao errar a conta no 1o mes
quo=totalquant//(totalmeses-mespulado) #calcula o quociente entre o quantitativo e o total de meses (se quo=0 tem menos quant que o total de meses, se quo>=1 tem mais quant do que o total de meses)
retorno="" #usado para retornar o valor depois de calculado
#_____________________________________________________________________RECÁLCULO DOS TOTAIS_________________________________________________________________________________________________________________#
#__________________________________________________________________________________________________________________________________________________________________________________________________________#
if meiodomes==True and MedFiltered[(MedFiltered['ITEM']==item)]['QUANTIDADE'].empty==True and mes!=1:                 #Check se precisa recalcular os valores totais                       
    totalquant=totalquant-(quo/2)      #diminui o total de quant pelo que vai ser colocado no 1o mes         #(recalcula se contrato começa no meio do mês e se não tiveram medições)
    mespulado=1                        #conta como mes pulado para arrumar os calculos dos outros meses
    quo=totalquant//(totalmeses-mespulado)      #recalcula o quo baseado no novo valor dos meses              
#______________________________________________________________________CÁLCULO DO VALOR FINAL______________________________________________________________________________________________________________#    #__________________________________________________________________________________________________________________________________________________________________________________________________________#                                              
if datames<MedFiltered[(MedFiltered['ITEM']==item)]['MÊS REFERÊNCIA'].min():    #check se tiveram meses medidos APÓS o mês em questão (pq n da pra colocar quantitativos no passado)
    quantitativomes=0                                                           #entao se n foi medido nada no mes ele fica 0
    mespulado=mes                                                               #e depois conta quantos foram pra tirar do total de meses inputados 👌
    
elif MedFiltered[(MedFiltered['ITEM']==item)&(MedFiltered['MÊS REFERÊNCIA']==datames)]['QUANTIDADE'].empty == False: #check se tem medição do item no mês em questão
    quantitativomes=MedFiltered[(MedFiltered['ITEM']==item)&(MedFiltered['MÊS REFERÊNCIA']==datames)]['QUANTIDADE'].sum()    #se tiver coloca o valor medido q(≧▽≦q)
    
else:                                        #se nao foi pulado nem medido, faz a conta do input
    match quo:                               #inicia o switch case baseado no quo:
        case 0:                              #se quo == 0 significa que tem menos quantitativo do que meses na planilha (entao a conta é ou 1 ou 0 (ou em alguns casos 1.5 se o valor do 1o mes for metade)
            if 0<(mes-mesesmedidos-mespulado)<=totalquant:                #se o mês em questão for menor que o numero de quantitativos
                if meiodomes==True and (mes-mesesmedidos-mespulado)==totalquant and MedFiltered[(MedFiltered['ITEM']==item)]['QUANTIDADE'].empty==True: #check se o 1o mes foi tirado metade do valor
                    if mes==totalmeses:    #check se ultimo mes
                        quantitativomes=1.5 #porque se for tem que colocar de volta aqui 👈(ﾟヮﾟ👈)
                    else:
                        quantitativomes=0.5 #se nao for coloca no mes apos o ultimo input
                else:
                    quantitativomes=1               #se nao foi tirado metade do 1o mes imprime um quantitativo 
            else:                                    #se o mês em questao for maior que o numero de quantitativos                            
                quantitativomes=0                      #nao coloca nada pq ja foram colocados todos os quantitativos
        case n if quo>0:                       #se quo > 0 significa que tem mais quantitativos que meses
            if (mes-mesesmedidos)!=totalmeses:                #se não for o ultimo mês
                quantitativomes=quo          #imprime quo quantitativos (lembra: quo é o valor espalhado dos quantitativos divido pelos meses que precisam dar input)
            else:                              #se for o ultimo mes
                quantitativomes=quo+(totalquant-(quo*(totalmeses-mespulado)))    #retorna o quantitativo dos meses + o resto
#_____________________________________________________________________________OUTPUT_______________________________________________________________________________________________________________________#          #__________________________________________________________________________________________________________________________________________________________________________________________________________#
if meiodomes==True and mes==1 and MedFiltered[(MedFiltered['ITEM']==item)]['QUANTIDADE'].empty==True: #ultimo check, se é o 1o mes e comeca na metade e n teve medição
    retorno = quantitativomes/2*precouni #divide o valor na metade (👉ﾟヮﾟ)👉
else: 
    retorno = quantitativomes*precouni #coloca o valor que ia colocar mesmo
#__________________________________________________________________________________________________________________________________________________________________________________________________________#
retorno #imprime o valor na celula
