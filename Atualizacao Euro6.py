# importar connection mocha 

import py3270 

import time 

import pandas as pd 

# Conectar no Mocha

em = py3270.Emulator(visible=True) 

em.connect('CPUA') 

# Usario e Senha 
time.sleep(2)

em.fill_field(11,60,'f036370',7) # primeiro linha depois coluna

em.fill_field(12,60,'antonio1',8) #o 8 representa o nº de caracteres

Euro6 = pd.read_excel(r"C:\Users\RCAMINA\Desktop\EURO6\teste_VBA.xlsx",header = None,  skiprows = 1,sheet_name = "Planilha1")

# Cria-se uma lista com todos os AFAB's 

list_AFAB = []
i = 0
for i in range(0,len(Euro6[12])):
    var = str(Euro6[12][i])#.astype(str)
    #print(var)
    oficial = var#[: - 2]
    if oficial != "n":
        list_AFAB.append(oficial)
    
print(list_AFAB)
print(len(list_AFAB))

# Commandos gerais 

em.send_enter() 

#em.send_enter()

em.send_string('3',2,15) 

em.send_enter() 

time.sleep(0.5) # = espere 0.5 secundos 

em.send_pf4() 

em.send_string('C',20,19) 

em.send_enter() 

time.sleep(0.5)

em.send_string('PRO-PMENU',24,32) 

em.send_enter() 

time.sleep(0.5)

em.send_string("L",6,18) 

em.send_enter()
 

list1 = [] # Crie uma lista 
list2 = []

# Iteração

#33 = len(Euro6[12]

for iteration in range (0,len(list_AFAB)): 

    AFAB = list_AFAB[iteration] # atribui cada afab da lista na variavel AFAB
    
    em.send_string(AFAB,4,16) # digita essa variável no campo de coordenada (4,16) -> linha coluna     

    em.send_enter() # pressiona ENTER
    
    BM = em.string_get(10,25,9).strip() # pega o boumster do veículo
    
    QVV = em.string_get(10,40,14).strip() # pega a string da coordenada (10,40). A string coletada tem tamanho de 14 chars
    
    FZ = em.string_get(11,27,7).strip()
    if FZ != '':
        FZ = int(FZ)
    
    NP = em.string_get(12,5,13).strip().replace(".","",2).replace("/","")  
    if NP != '':
        NP = int(NP)
        
    
    em.send_string("e",10,2) # para entrar e ver os agregados do veiculo
    
    em.send_enter()
    
    time.sleep(0.5) # = espere 0.5 secundos
    
    while True:
        
        if em.string_found(11,5,"H"):
            eixo_tra1 = em.string_get(11,60,10)
            fz_t1 = em.string_get(11,51,6)
            if fz_t1 != '      ':
                fz_t1 = int(fz_t1)
            BMH = em.string_get(11,20,9).replace(" ","").replace(".","")
            break
        elif em.string_found(14,5,"H"):
            eixo_tra1 = em.string_get(14,60,10)
            fz_t1 = em.string_get(14,51,6)
            if fz_t1 != '      ':
                fz_t1 = int(fz_t1)
            BMH = em.string_get(14,20,9).replace(" ","").replace(".","")
            break
        elif em.string_found(17,5,"H"):
            eixo_tra1 = em.string_get(17,60,10)
            fz_t1 = em.string_get(17,51,6)
            if fz_t1 != '      ':
                fz_t1 = int(fz_t1)
            BMH = em.string_get(17,20,9).replace(" ","").replace(".","")
            break
        elif em.string_found(20,5,"H"):
            eixo_tra1 = em.string_get(20,60,10)
            fz_t1 = em.string_get(20,51,6)
            if fz_t1 != '      ':
                fz_t1 = int(fz_t1)
            BMH = em.string_get(20,20,9).replace(" ","").replace(".","")
            break
        if em.string_found(23,2,"F"):
            eixo_tra1 = ""
            fz_t1 = ""
            BMH = ""
            break
        em.send_pf8()
        
    time.sleep(0.5) # = espere 0.5 secundos
    
    while True:
        
        if em.string_found(11,5,"J"):
            eixo_tra2 = em.string_get(11,60,10)
            fz_t2 = em.string_get(11,51,6)
            if fz_t2 != '      ':
                fz_t2 = int(fz_t2)
            BMJ = em.string_get(11,20,9).replace(" ","").replace(".","")
            break
        elif em.string_found(14,5,"J"):
            eixo_tra2 = em.string_get(14,60,10)
            fz_t2 = em.string_get(14,51,6)
            if fz_t2 != '      ':
                fz_t2 = int(fz_t2)
            BMJ = em.string_get(14,20,9).replace(" ","").replace(".","")
            break
        elif em.string_found(17,5,"J"):
            eixo_tra2 = em.string_get(17,60,10)
            fz_t2 = em.string_get(17,51,6)
            if fz_t2 != '      ':
                fz_t2 = int(fz_t2)
            BMJ = em.string_get(17,20,9).replace(" ","").replace(".","")
            break
        elif em.string_found(20,5,"J"):
            eixo_tra2 = em.string_get(20,60,10)
            fz_t2 = em.string_get(20,51,6)
            if fz_t2 != '      ':
                fz_t2 = int(fz_t2)
            BMJ = em.string_get(20,20,9).replace(" ","").replace(".","")
            break
        if em.string_found(23,2,"F"):
            eixo_tra2 = ""
            fz_t2 = ""
            BMJ = ""
            break
        em.send_pf8()
        
    time.sleep(0.5) # = espere 0.5 secundos

    
    while True:
        
        if em.string_found(11,5,"V"):
            eixo_di1 = em.string_get(11,60,10)
            fz_d1 = em.string_get(11,51,6)
            if fz_d1 != '      ':
                fz_d1 = int(fz_d1)
            BMV = em.string_get(11,20,9).replace(" ","").replace(".","")
            break
        elif em.string_found(14,5,"V"):
            eixo_di1 = em.string_get(14,60,10)
            fz_d1 = em.string_get(14,51,6)
            if fz_d1 != '      ':
                fz_d1 = int(fz_d1)
            BMV = em.string_get(14,20,9).replace(" ","").replace(".","")
            break
        elif em.string_found(17,5,"V"):
            eixo_di1 = em.string_get(17,60,10)
            fz_d1 = em.string_get(17,51,6)
            if fz_d1 != '      ':
                fz_d1 = int(fz_d1)
            BMV = em.string_get(17,20,9).replace(" ","").replace(".","")
            break
        elif em.string_found(20,5,"V"):
            eixo_di1 = em.string_get(20,60,10)
            fz_d1 = em.string_get(20,51,6)
            if fz_d1 != '      ':
                fz_d1 = int(fz_d1)
            BMV = em.string_get(20,20,9).replace(" ","").replace(".","")
            break
        if em.string_found(23,2,"F"):
            eixo_di1 = ""
            fz_d1 = ""
            BMV = ""
            break
        em.send_pf8()
    
    time.sleep(0.5) # = espere 0.5 secundos
    
    while True:
        
        if em.string_found(11,5,"W"):
            eixo_di2 = em.string_get(11,60,10)
            fz_d2 = em.string_get(11,51,6) 
            if fz_d2 != '      ':
                fz_d2 = int(fz_d2)
            BMW = em.string_get(11,20,9).replace(" ","").replace(".","")
            break
        elif em.string_found(14,5,"W"):
            eixo_di2 = em.string_get(14,60,10)
            fz_d2 = em.string_get(14,51,6)
            if fz_d2 != '      ':
                fz_d2 = int(fz_d2)
            BMW = em.string_get(14,20,9).replace(" ","").replace(".","")
            break
        elif em.string_found(17,5,"W"):
            eixo_di2 = em.string_get(17,60,10)
            fz_d2 = em.string_get(17,51,6)
            if fz_d2 != '      ':
                fz_d2 = int(fz_d2)
            BMW = em.string_get(17,20,9).replace(" ","").replace(".","")
            break
        elif em.string_found(20,5,"W"):
            eixo_di2 = em.string_get(20,60,10)
            fz_d2 = em.string_get(20,51,6)
            if fz_d2 != '      ':
                fz_d2 = int(fz_d2)
            BMW = em.string_get(20,20,9).replace(" ","").replace(".","")
            break
        if em.string_found(23,2,"F"):
            eixo_di2 = ""
            fz_d2 = ""
            BMW = ""
            break
        em.send_pf8()
        
    time.sleep(0.5) # = espere 0.5 secundos
    
    em.send_pf7()
    
    while True:
        
        
        if em.string_found(11,5,"G"):
            cambio = em.string_get(11,60,10)
            fz_cambio = em.string_get(11,51,6) 
            if fz_cambio != '      ':
                fz_cambio = int(fz_cambio)
            BMcambio = em.string_get(11,20,9).replace(" ","").replace(".","")
            break
        elif em.string_found(14,5,"G"):
            cambio = em.string_get(14,60,10)
            fz_cambio = em.string_get(14,51,6)
            if fz_cambio != '      ':
                fz_cambio = int(fz_cambio)
            BMcambio = em.string_get(14,20,9).replace(" ","").replace(".","")
            break
        elif em.string_found(17,5,"G"):
            cambio = em.string_get(17,60,10)
            fz_cambio = em.string_get(17,51,6)
            if fz_cambio != '      ':
                fz_cambio = int(fz_cambio)
            BMcambio = em.string_get(17,20,9).replace(" ","").replace(".","")
            break
        elif em.string_found(20,5,"G"):
            cambio = em.string_get(20,60,10)
            fz_cambio = em.string_get(20,51,6)
            if fz_cambio != '      ':
                fz_cambio = int(fz_cambio)
            BMcambio = em.string_get(20,20,9).replace(" ","").replace(".","")
            break
        if em.string_found(23,2,"F"):
            cambio = ""
            fz_cambio = ""
            BMcambio = ""
            break
        em.send_pf8()
        
    list1.append((QVV,NP,FZ,AFAB)) # Cria uma lista com as variaveis
    list2.append((BMV,eixo_di1,fz_d1,BMW,eixo_di2,fz_d2,BMH,eixo_tra1,fz_t1,BMJ,eixo_tra2,fz_t2,BMcambio,cambio,fz_cambio))
    em.send_pf(12)

    output_table = pd.DataFrame(list1,columns=["QVV","NP","FZ","AFAB"])# Cria um dataframe com a lista1 com as colunas seguintes... 
output_table2 = pd.DataFrame(list2,columns=["BM","V","FZ","BM","W","FZ","BM","H","FZ","BM","J","FZ","BM","G","FZ"])

em.terminate() # fecha a janela do emulador

with pd.ExcelWriter(r"C:\Users\RCAMINA\Desktop\EURO6\resposta_robo.xlsx") as writer:
    output_table.to_excel(writer, sheet_name = "Sheet1")
    output_table2.to_excel(writer, sheet_name = "Sheet2")                             