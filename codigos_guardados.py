try:
        tabela = pd.read_excel(r'C:\Users\Customer\Desktop\nw_codigos\app\excell\dados_dia.xlsx', sheet_name=data)
        faturamento_por_barbeiro = tabela[['profissional','entrada']].groupby('profissional').sum()
        x = faturamento_por_barbeiro.to_dict() 
        lista_raiz = x['entrada']
        #dividindo o dict 'lista_raiz' em 2 listas
        lista_barbeiros = []
        lista_barbeiros_fat = []
        for barbeiro, barbeiro_fat in lista_raiz.items():
            lista_barbeiros.append(barbeiro)
            lista_barbeiros_fat.append(barbeiro_fat)

        #esvaziando as 2 listas dissolvendo-as em 1 única lista decrescente
        lista_final = []
        while True:
            if len(lista_barbeiros) == 0:
                break
            maior_fat = max(lista_barbeiros_fat)
            indice_do_maior_fat = lista_barbeiros_fat.index(maior_fat)
            barbeiro = lista_barbeiros[indice_do_maior_fat]
            #adicionando o item na 'lista intermediaria'
            lista_intermediaria = []
            lista_intermediaria.append(barbeiro) 
            lista_intermediaria.append(maior_fat)
            #adicionando o item na 'lista final'
            lista_final.append(lista_intermediaria)
            #removendo o item das 2 listas
            lista_barbeiros.remove(barbeiro)
            lista_barbeiros_fat.remove(maior_fat)
        return lista_final
    except:
        return ['Erro no cálculo']