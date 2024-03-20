import pandas as pd
import regex as re
import math

df = pd.read_excel("623 - Acesso Equipamentos.xls", engine="xlrd")

new_df = df.assign(Resultado="", Valores = "")

def deb_or_cred(row):
    
    type_row = ""

    if not math.isnan(row[5]):
        type_row = "Débito"
        return type_row
    else:
        type_row = "Crédito"
        return type_row

def verify_row_accept(array_row_accept, row):

    type_row = deb_or_cred(row=row)

    if len(array_row_accept) == 0: 

        if type_row == "Débito":
            new_df["Resultado"][row[0]] = "Falta nota"
            new_df["Valores"].iat[row[0]] = row[5]
        elif type_row == "Crédito":
            new_df["Resultado"][row[0]] = "Falta pagamento"
            new_df["Valores"].iat[row[0]] = row[6]

for row in new_df.itertuples():

    control_row_accept = []
    control_row_search_more_than_one_deb = []

    list_nf = re.findall("NF (\d+)", row[3]) or re.findall("NF. (\d+)", row[3]) or re.findall("NF- (\d+)", row[3]) or re.findall("DOC (\d+)", row[3]) or re.findall("NF -(\d+)", row[3]) or re.findall("NF  (\d+)", row[3]) 

    if len(list_nf) > 0:

        if list_nf[0][0:4] == "2022":

            nf_sem_2022 = list_nf[0][4:]

            type_row = deb_or_cred(row=row)

            if type_row == "Débito":

                for row_deb2 in new_df.itertuples():

                    if row[0] != row_deb2[0]:

                        type_row2 = deb_or_cred(row=row_deb2)

                        if type_row2 == "Débito":

                            list_nf_row_deb2 = re.findall("NF (\d+)", row_deb2[3]) or re.findall("NF. (\d+)", row_deb2[3]) or re.findall("NF- (\d+)", row_deb2[3]) or re.findall("DOC (\d+)", row_deb2[3]) or re.findall("NF -(\d+)", row_deb2[3]) or re.findall("NF  (\d+)", row_deb2[3]) 

                            if len(list_nf_row_deb2) > 0:

                                if list_nf_row_deb2[0][0:4] == "2022":

                                    nf_deb2_sem_2022 = list_nf_row_deb2[0][4:]

                                    if nf_sem_2022 == nf_deb2_sem_2022:

                                        valor_row_deb2 = row_deb2[5]

                                        control_row_search_more_than_one_deb.append(valor_row_deb2)

                                    else:
                                        continue
                                else:

                                    if nf_sem_2022 == list_nf_row_deb2[0]:

                                        valor_row_deb2 = row_deb2[5]

                                        control_row_search_more_than_one_deb.append(valor_row_deb2)

                                    else:
                                        continue    
                            else:
                                continue
                            
                        else:
                            continue
                    else:
                        continue
                control_row_search_more_than_one_deb.append(row[5])       
                if len(control_row_search_more_than_one_deb) > 1:
                    i = 0
                    size_array = len(control_row_search_more_than_one_deb) - 1
                    value_portion_deb = 0
                    while i <= size_array:
                        value_portion_deb += control_row_search_more_than_one_deb[i]
                        i+=1

                    value_portion_deb_round = round(value_portion_deb, 2)

                elif len(control_row_search_more_than_one_deb) == 1:
                    value_portion_deb = control_row_search_more_than_one_deb[0]  
                    value_portion_deb_round = round(value_portion_deb, 2)

                for row_cred in new_df.itertuples():

                    if row[0] != row_cred[0]:

                        type_row2 = deb_or_cred(row=row_cred)

                        if type_row2 == "Crédito":

                            list_nf_row_cred = re.findall("NF (\d+)", row_cred[3]) or re.findall("NF. (\d+)", row_cred[3]) or re.findall("NF- (\d+)", row_cred[3]) or re.findall("DOC (\d+)", row_cred[3]) or re.findall("NF -(\d+)", row_cred[3]) or re.findall("NF  (\d+)", row_cred[3]) 

                            if len(list_nf_row_cred) > 0:

                                if list_nf_row_cred[0][0:4] == "2022":

                                    nf_cred_sem_2022 = list_nf_row_cred[0][4:]

                                    if nf_sem_2022 == nf_cred_sem_2022:

                                        control_row_accept.append("Achei")
                                        
                                        valor_row_cred = row_cred[6]

                                        result = value_portion_deb_round - valor_row_cred

                                        if result == 0.0:
                                        
                                            new_df["Resultado"].iat[row[0]] = "OK"
                                            new_df["Valores"].iat[row[0]] = result

                                        elif result > 0.0:
                           
                                            if type_row == "Crédito":        
                                                new_df["Resultado"].iat[row[0]] = "OK - Falta pagamento de parcelas"
                                                new_df["Valores"].iat[row[0]] = result
                                            else:
                                                new_df["Resultado"].iat[row[0]] = "OK - Pagamento com juros"
                                                new_df["Valores"].iat[row[0]] = result  

                                        elif result < 0.0:
                                        
                                            if type_row == "Crédito":        
                                                new_df["Resultado"].iat[row[0]] = "OK - Falta valor de nota"
                                                new_df["Valores"].iat[row[0]] = result
                                            else:
                                                new_df["Resultado"].iat[row[0]] = "OK - Falta pagamento de parcelas"
                                                new_df["Valores"].iat[row[0]] = result    

                                    else:
                                        continue
                                else:

                                    if nf_sem_2022 == list_nf_row_cred[0]:

                                        control_row_accept.append("Achei")

                                        valor_row_cred = row_cred[6]

                                        result = value_portion_deb_round - valor_row_cred

                                        if result == 0.0:
                                        
                                            new_df["Resultado"].iat[row[0]] = "OK"
                                            new_df["Valores"].iat[row[0]] = result

                                        elif result > 0.0:
                           
                                            if type_row == "Crédito":        
                                                new_df["Resultado"].iat[row[0]] = "OK - Falta pagamento de parcelas"
                                                new_df["Valores"].iat[row[0]] = result
                                            else:
                                                new_df["Resultado"].iat[row[0]] = "OK - Pagamento com juros"
                                                new_df["Valores"].iat[row[0]] = result  

                                        elif result < 0.0:
                                        
                                            if type_row == "Crédito":        
                                                new_df["Resultado"].iat[row[0]] = "OK - Falta valor de nota"
                                                new_df["Valores"].iat[row[0]] = result
                                            else:
                                                new_df["Resultado"].iat[row[0]] = "OK - Falta pagamento de parcelas"
                                                new_df["Valores"].iat[row[0]] = result 

                                    else:
                                        continue    
                            else:
                                continue
                            
                        else:
                            continue
                    else:
                        continue
            else:

                for row_deb in new_df.itertuples():

                    if row[0] != row_deb[0]:

                        type_row2 = deb_or_cred(row=row_deb)

                        if type_row2 == "Débito":

                            list_nf_row_deb = re.findall("NF (\d+)", row_deb[3]) or re.findall("NF. (\d+)", row_deb[3]) or re.findall("NF- (\d+)", row_deb[3]) or re.findall("DOC (\d+)", row_deb[3]) or re.findall("NF -(\d+)", row_deb[3]) or re.findall("NF  (\d+)", row_deb[3]) 

                            if len(list_nf_row_deb) > 0:

                                if list_nf_row_deb[0][0:4] == "2022":

                                    nf_deb_sem_2022 = list_nf_row_deb[0][4:]

                                    if nf_sem_2022 == nf_deb_sem_2022:

                                        valor_row_deb = row_deb[5]

                                        control_row_search_more_than_one_deb.append(valor_row_deb)

                                        control_row_accept.append("Achei")

                                    else:
                                        continue
                                else:

                                    if nf_sem_2022 == list_nf_row_deb[0]:

                                        valor_row_deb = row_deb[5]

                                        control_row_search_more_than_one_deb.append(valor_row_deb)

                                        control_row_accept.append("Achei")

                                    else:
                                        continue    
                            else:
                                continue
                            
                        else:
                            continue
                    else:
                        continue         
                if len(control_row_search_more_than_one_deb) > 1:
                    i = 0
                    size_array = len(control_row_search_more_than_one_deb) - 1
                    value_portion_deb = 0
                    while i <= size_array:
                        value_portion_deb += control_row_search_more_than_one_deb[i]
                        i+=1

                    value_portion_deb_round = round(value_portion_deb, 2)

                    value_main_row_cred = row[6]   

                    result = value_main_row_cred - value_portion_deb_round

                    if result == 0.0:
                                        
                        new_df["Resultado"].iat[row[0]] = "OK"
                        new_df["Valores"].iat[row[0]] = result

                    elif result > 0.0:
                           
                        if type_row == "Crédito":        
                            new_df["Resultado"].iat[row[0]] = "OK - Falta pagamento de parcelas"
                            new_df["Valores"].iat[row[0]] = result
                        else:
                            new_df["Resultado"].iat[row[0]] = "OK - Pagamento com juros"
                            new_df["Valores"].iat[row[0]] = result  

                    elif result < 0.0:
                                        
                        if type_row == "Crédito":        
                            new_df["Resultado"].iat[row[0]] = "OK - Falta valor de nota"
                            new_df["Valores"].iat[row[0]] = result
                        else:
                            new_df["Resultado"].iat[row[0]] = "OK - Falta pagamento de parcelas"
                            new_df["Valores"].iat[row[0]] = result

                elif len(control_row_search_more_than_one_deb) == 1:
                    value_portion_deb = control_row_search_more_than_one_deb[0]

                    value_portion_deb_round = round(value_portion_deb, 2)

                    value_main_row_cred = row[6]    

                    result = value_main_row_cred - value_portion_deb_round

                    if result == 0.0:
                                        
                        new_df["Resultado"].iat[row[0]] = "OK"
                        new_df["Valores"].iat[row[0]] = result

                    elif result > 0.0:
                           
                        if type_row == "Crédito":        
                            new_df["Resultado"].iat[row[0]] = "OK - Falta pagamento de parcelas"
                            new_df["Valores"].iat[row[0]] = result
                        else:
                            new_df["Resultado"].iat[row[0]] = "OK - Pagamento com juros"
                            new_df["Valores"].iat[row[0]] = result  

                    elif result < 0.0:
                                        
                        if type_row == "Crédito":        
                            new_df["Resultado"].iat[row[0]] = "OK - Falta valor de nota"
                            new_df["Valores"].iat[row[0]] = result
                        else:
                            new_df["Resultado"].iat[row[0]] = "OK - Falta pagamento de parcelas"
                            new_df["Valores"].iat[row[0]] = result  
                            
        else:
            
            type_row = deb_or_cred(row=row)

            if type_row == "Débito":

                for row_deb2 in new_df.itertuples():

                    if row[0] != row_deb2[0]:

                        type_row2 = deb_or_cred(row=row_deb2)

                        if type_row2 == "Débito":

                            list_nf_row_deb2 = re.findall("NF (\d+)", row_deb2[3]) or re.findall("NF. (\d+)", row_deb2[3]) or re.findall("NF- (\d+)", row_deb2[3]) or re.findall("DOC (\d+)", row_deb2[3]) or re.findall("NF -(\d+)", row_deb2[3]) or re.findall("NF  (\d+)", row_deb2[3]) 

                            if len(list_nf_row_deb2) > 0:

                                if list_nf_row_deb2[0][0:4] == "2022":

                                    nf_deb2_sem_2022 = list_nf_row_deb2[0][4:]

                                    if list_nf[0] == nf_deb2_sem_2022:

                                        valor_row_deb2 = row_deb2[5]

                                        control_row_search_more_than_one_deb.append(valor_row_deb2)

                                    else:
                                        continue
                                else:
                                    if list_nf[0] == list_nf_row_deb2[0]:

                                        valor_row_deb2 = row_deb2[5]

                                        control_row_search_more_than_one_deb.append(valor_row_deb2)

                                    else:
                                        continue    
                            else:
                                continue
                            
                        else:
                            continue
                    else:
                        continue
                control_row_search_more_than_one_deb.append(row[5])        
                if len(control_row_search_more_than_one_deb) > 1:
                    i = 0
                    size_array = len(control_row_search_more_than_one_deb) - 1
                    value_portion_deb = 0
                    while i <= size_array:
                        value_portion_deb += control_row_search_more_than_one_deb[i]
                        i+=1
                    value_portion_deb_round = round(value_portion_deb, 2)
                elif len(control_row_search_more_than_one_deb) == 1:
                    value_portion_deb = control_row_search_more_than_one_deb[0]  
                    value_portion_deb_round = round(value_portion_deb, 2)
                for row_cred in new_df.itertuples():

                    if row[0] != row_cred[0]:

                        type_row2 = deb_or_cred(row=row_cred)

                        if type_row2 == "Crédito":

                            list_nf_row_cred = re.findall("NF (\d+)", row_cred[3]) or re.findall("NF. (\d+)", row_cred[3]) or re.findall("NF- (\d+)", row_cred[3]) or re.findall("DOC (\d+)", row_cred[3]) or re.findall("NF -(\d+)", row_cred[3]) or re.findall("NF  (\d+)", row_cred[3]) 

                            if len(list_nf_row_cred) > 0:

                                if list_nf_row_cred[0][0:4] == "2022":

                                    nf_cred_sem_2022 = list_nf_row_cred[0][4:]

                                    if list_nf[0] == nf_cred_sem_2022:

                                        control_row_accept.append("Achei")

                                        valor_row_cred = row_cred[6]

                                        result = value_portion_deb_round - valor_row_cred

                                        if result == 0.0:
                                        
                                            new_df["Resultado"].iat[row[0]] = "OK"
                                            new_df["Valores"].iat[row[0]] = result

                                        elif result > 0.0:
                           
                                            if type_row == "Crédito":        
                                                new_df["Resultado"].iat[row[0]] = "OK - Falta pagamento de parcelas"
                                                new_df["Valores"].iat[row[0]] = result
                                            else:
                                                new_df["Resultado"].iat[row[0]] = "OK - Pagamento com juros"
                                                new_df["Valores"].iat[row[0]] = result  

                                        elif result < 0.0:
                                        
                                            if type_row == "Crédito":        
                                                new_df["Resultado"].iat[row[0]] = "OK - Falta valor de nota"
                                                new_df["Valores"].iat[row[0]] = result
                                            else:
                                                new_df["Resultado"].iat[row[0]] = "OK - Falta pagamento de parcelas"
                                                new_df["Valores"].iat[row[0]] = result    

                                    else:
                                        continue
                                else:

                                    if list_nf[0] == list_nf_row_cred[0]:

                                        control_row_accept.append("Achei")

                                        valor_row_cred = row_cred[6]

                                        result = value_portion_deb_round - valor_row_cred

                                        if result == 0.0:
                                        
                                            new_df["Resultado"].iat[row[0]] = "OK"
                                            new_df["Valores"].iat[row[0]] = result

                                        elif result > 0.0:
                           
                                            if type_row == "Crédito":        
                                                new_df["Resultado"].iat[row[0]] = "OK - Falta pagamento de parcelas"
                                                new_df["Valores"].iat[row[0]] = result
                                            else:
                                                new_df["Resultado"].iat[row[0]] = "OK - Pagamento com juros"
                                                new_df["Valores"].iat[row[0]] = result  

                                        elif result < 0.0:
                                        
                                            if type_row == "Crédito":        
                                                new_df["Resultado"].iat[row[0]] = "OK - Falta valor de nota"
                                                new_df["Valores"].iat[row[0]] = result
                                            else:
                                                new_df["Resultado"].iat[row[0]] = "OK - Falta pagamento de parcelas"
                                                new_df["Valores"].iat[row[0]] = result 

                                    else:
                                        continue    
                            else:
                                continue
                            
                        else:
                            continue
                    else:
                        continue

            
            else:

                for row_deb in new_df.itertuples():

                    if row[0] != row_deb[0]:

                        type_row2 = deb_or_cred(row=row_deb)

                        if type_row2 == "Débito":

                            list_nf_row_deb = re.findall("NF (\d+)", row_deb[3]) or re.findall("NF. (\d+)", row_deb[3]) or re.findall("NF- (\d+)", row_deb[3]) or re.findall("DOC (\d+)", row_deb[3]) or re.findall("NF -(\d+)", row_deb[3]) or re.findall("NF  (\d+)", row_deb[3]) 

                            if len(list_nf_row_deb) > 0:

                                if list_nf_row_deb[0][0:4] == "2022":

                                    nf_deb_sem_2022 = list_nf_row_deb[0][4:]

                                    if list_nf[0] == nf_deb_sem_2022:

                                        valor_row_deb = row_deb[5]

                                        control_row_search_more_than_one_deb.append(valor_row_deb)

                                        control_row_accept.append("Achei")

                                    else:
                                        continue
                                else:

                                    if list_nf[0] == list_nf_row_deb[0]:

                                        valor_row_deb = row_deb[5]

                                        control_row_search_more_than_one_deb.append(valor_row_deb)

                                        control_row_accept.append("Achei")

                                    else:
                                        continue    
                            else:
                                continue
                            
                        else:
                            continue
                    else:
                        continue       
                if len(control_row_search_more_than_one_deb) > 1:
                    i = 0
                    size_array = len(control_row_search_more_than_one_deb) - 1
                    value_portion_deb = 0
                    while i <= size_array:
                        value_portion_deb += control_row_search_more_than_one_deb[i]
                        i+=1

                    value_portion_deb_round = round(value_portion_deb, 2)

                    value_main_row_cred = row[6]    

                    result = value_main_row_cred - value_portion_deb_round

                    if result == 0.0:
                                        
                        new_df["Resultado"].iat[row[0]] = "OK"
                        new_df["Valores"].iat[row[0]] = result

                    elif result > 0.0:
                           
                        if type_row == "Crédito":        
                            new_df["Resultado"].iat[row[0]] = "OK - Falta pagamento de parcelas"
                            new_df["Valores"].iat[row[0]] = result
                        else:
                            new_df["Resultado"].iat[row[0]] = "OK - Pagamento com juros"
                            new_df["Valores"].iat[row[0]] = result  

                    elif result < 0.0:
                                        
                        if type_row == "Crédito":        
                            new_df["Resultado"].iat[row[0]] = "OK - Falta valor de nota"
                            new_df["Valores"].iat[row[0]] = result
                        else:
                            new_df["Resultado"].iat[row[0]] = "OK - Falta pagamento de parcelas"
                            new_df["Valores"].iat[row[0]] = result

                elif len(control_row_search_more_than_one_deb) == 1 :
                    value_portion_deb = control_row_search_more_than_one_deb[0]

                    value_portion_deb_round = round(value_portion_deb, 2)
                    
                    value_main_row_cred = row[6]    

                    result = value_main_row_cred - value_portion_deb_round

                    if result == 0.0:
                                        
                        new_df["Resultado"].iat[row[0]] = "OK"
                        new_df["Valores"].iat[row[0]] = result

                    elif result > 0.0:
                           
                        if type_row == "Crédito":        
                            new_df["Resultado"].iat[row[0]] = "OK - Falta pagamento de parcelas"
                            new_df["Valores"].iat[row[0]] = result
                        else:
                            new_df["Resultado"].iat[row[0]] = "OK - Pagamento com juros"
                            new_df["Valores"].iat[row[0]] = result  

                    elif result < 0.0:
                                        
                        if type_row == "Crédito":        
                            new_df["Resultado"].iat[row[0]] = "OK - Falta valor de nota"
                            new_df["Valores"].iat[row[0]] = result
                        else:
                            new_df["Resultado"].iat[row[0]] = "OK - Falta pagamento de parcelas"
                            new_df["Valores"].iat[row[0]] = result

        verify_row_accept(array_row_accept=control_row_accept, row=row)                         
            
    else:
        new_df["Resultado"][row[0]] = "Não localizei Número de nota"

        type_row = deb_or_cred(row=row)

        if type_row == "Débito":
            new_df["Valores"].iat[row[0]] = row[5]
        elif type_row == "Crédito":
            new_df["Valores"].iat[row[0]] = row[6]


new_df.to_excel("razão.xlsx", index=False)
