import math
import pandas as pd
from tkinter import filedialog, messagebox, Tk

#perguntas que precisam ser corrigidas colocando um "usa" na frente (cat. microcomputadores)
usa_em = ['Em casa:',
          'No trabalho:',
          'Na escola:',
          'Em outros locais:',
          'Para trabalhos profissionais:',
          'Para trabalhos escolares:',
          'Para entretenimento (músicas, vídeos, redes sociais, etc)',
          'Para comunicação por email:',
          'Para operações bancárias:',
          'Para compras na internet:'] #perguntas que precisam ser corrigidas colocando um "usa" na frente

def mensagemErro():
    '''Função para exibir mensagem de erro'''
    messagebox.showerror('Ocorreu um erro!','O arquivo selecionado não é uma planilha de respostas do questionário!\n\nPor favor, tente novamente!')

def filtrarFechada(dados_pergunta):
    ''' Essa função tem o papel de filtrar células vazias das pergunts aberta
        irá retornar uma lista corrigida, com apenas as respostas (células que tem conteúdo)

        Argumento: dados_pergunta --> Coluna de repostas
    '''
    valores = {}

    for resposta in dados_pergunta:
        if isinstance(resposta, str):
            # Verifica se a resposta é uma string vazia ou apenas contém espaços em branco
            if resposta.strip() == '':
                continue  # Ignora células vazias
        elif isinstance(resposta, float):# Verifica se a resposta é NaN
            if math.isnan(resposta):
                continue #Ignora NaN
        
        if resposta not in valores:
            valores[resposta] = 1 #adiciona um para a qntd de vezes respondida (caso não tenha aparecido antes)
        else:
            valores[resposta] +=1 #soma um caso já tenha aparecido antes

    return valores

def filtrarNuvem(dados_nuvem): #filtra células vazias
    ''' Essa função tem o papel de filtrar células vazias das perguntas de nuvem
        irá retornar uma lista corrigida, com apenas as respostas (células que tem conteúdo)

        Argumento: dados_pergunta --> Coluna de repostas
    '''
    valores = [] #lista (ainda) vazia dos dados filtrados

    for resposta in dados_nuvem:
        if pd.notna(resposta): 
            if isinstance(resposta,str): # verificando se a resposta não é NaN
                if resposta.strip() == '': # verifica se a resposta é uma string msm ou é apenas uma célula vazia
                    continue
            valores.append(str(resposta)) #após a filtragem, adc na lista
    
    

    return valores

def formatar_porcentagem(respostas):
    def calculo_prcnt(pct):
        total = sum(respostas)
        valor = int(round(pct*total/100))
        return '{p:.1f}%({v:d})'.format(p=pct,v=valor)
    return calculo_prcnt

def processar_moradia(dados_moradia):
    '''Essa é a função que processa o tempo de moradia, transformando meses em anos!
    
    Param dados_moradias -> valores originais da coluna'''
    
    respostas = {}
    menoscincoanos = 0
    cincodezanos = 0
    onzevinteanos = 0
    maisvinte = 0
    for meses in dados_moradia:
        anos = meses // 12
        if anos < 5:
            menoscincoanos += 1
        elif anos >= 5 and anos <=10:
            cincodezanos += 1
        elif anos >= 11 and anos <= 20:
            onzevinteanos += 1
        else:
            maisvinte +=1
    respostas['Menos de 5 anos'] = menoscincoanos
    respostas['5 a 10 anos'] = cincodezanos
    respostas['11 a 20 anos'] = onzevinteanos
    respostas['Mais de 20 anos'] = maisvinte
    
    return respostas

def processar_idade(dados_anonasc):
    from datetime import datetime as dt #bibl. de importar o ano
    ano_atual = dt.now().year
  
    respostas = {}
    dezoito25 = 0
    vinteseis35 = 0             
    trintaseis45 = 0
    maisquarentaseis = 0
    #filtrando:
    for ano_nascimento in dados_anonasc:
        idade = ano_atual - ano_nascimento
        if idade>=18 and idade<=25:
            dezoito25+=1
        elif idade>=26 and idade<=35:
            vinteseis35+=1
        elif idade>=36 and idade<=45:
            trintaseis45+=1
        else:
            maisquarentaseis+=1
    #Jogando numa lista:
    respostas['18 a 25 anos'] = dezoito25
    respostas['26 a 35 anos'] = vinteseis35
    respostas['36 a 45 anos'] = trintaseis45
    respostas['Mais de 46 anos'] = maisquarentaseis

    return respostas

def corrigir_respostas(vlrs):
    """Função responsável pela correção de erros ortográficos, textos oversized ou imprecisos NAS RESPOSTAS"""  
    def corrigir_salariominimo(x,y):#retornar strings do salário minimo corrigidas
        return f'De {x} a {y} Salários Mínimos'
    
    corrigida = {}

    for chave, resposta in vlrs.items(): #iterando as labels (chaves)
        #Condicionais da categoria TECNOLOGIA:
        if chave == 'Nenhum (só ouvi falar)':
            chave = 'Nenhum'
        
        elif chave == 'Pouco ("Me viro")':
            chave = 'Pouco'  #corrigindo oversize
        
        elif chave == 'Intermediário (já tenho certa habilidade e experiência)':
            chave = 'Intermediário'
        
        elif chave == 'Avançado (Sei como usar as funcionalidades mais complexas)':
            chave = 'Avançado'
        
        elif chave == 'Muito Avançado (sou capaz de prestar concursos)':
            chave = 'Muito Avançado'

        # Condicionais da pergunt de RENDA MENSAL:
        elif chave == 'De 5 SALÁRIOS MÍNIMOS (R$ 6.660,00) a 10 SALÁRIOS MÍNIMOS (R$ 13.200,00)':
            chave = corrigir_salariominimo(5,10)
        
        elif chave == 'De 1 SALÁRIO MÍNIMO (R$ 1.320,00) a 2  SALÁRIOS MÍNIMOS (R$ 2.640,00)':
            chave = corrigir_salariominimo(1,2) #como é tudo a msm coisa, coloquei nessa função
        
        elif chave == 'De 2 SALÁRIOS MÍNIMOS (R$ 2.640,00) a 5 SALÁRIOS MÍNIMOS (R$ 6.660,00)':
            chave = corrigir_salariominimo(2,5)

        elif chave == 'Até 1 SALÁRIO MÍNIMO (R$ 1.320,00)':
            chave = 'Até 1 Salário Mínimo'

        elif chave == 'De 10 SALÁRIOS MÍNIMOS (R$ 13.200,00) a 20 SALÁRIOS MÍNIMOS (R$ 26.400,00)':
            chave = corrigir_salariominimo(10,20)
        
        #Categoria Trabalho
        elif chave == 'Sou registrado em empresa pública (Federal, Estadual, Municipal)':
            chave = 'Sou registrado em empresa pública'

        #Categoria Pessoal
        elif chave == 'Tenho e é uma plano individual':
            chave = 'Tenho e é um plano individual.'

        elif chave == 'Ensino Fundamental II (6º ao 9º Ano)':
            chave = 'Ensino Fundamental II'
        
        elif chave == 'Ensino Fundamental I (1º ao 5º Ano)':
            chave = 'Ensino Fundamental I'
        
        elif chave == 'Ensino Médio incompleto (1º e 2º anos)':
            chave = 'Ensino Médio incompleto'
        
        elif chave == 'Não calculo, mas sempre que sobra tempo estou vendo notícias na internet':
            chave = 'Não calculo, mas sempre que me sobra tempo eu leio'

        elif chave == 'Melhorar-me como pessoa para bons relacionamentos futuros':
            chave = 'Melhorar-me como pessoa'
        
        elif chave == 'Obter um (ou mais um) diploma de nível superior':
            chave = 'Obter um diploma de nível superior'

        elif chave == 'Com amigos (compartilhando despesas) ou de favor, república.':
            chave = 'Com amigos, ou de favor'

        # da pergunta: "Como conheceu a FATEC?"
        elif chave == 'Indicação (de familiar/amigo/patrão/colegas de trabalho, etc)':
            chave = 'Indicação de conhecidos'
        
        elif chave == 'Cartaz (de divulgação de vestibulares)':
            chave = 'Cartaz de divulgação'
        
        elif chave == 'Propaganda feita pela FATEC na escola que estudava':
            chave = 'Propaganda da FATEC, na minha escola'
        
        elif chave == 'Propaganda feita pela FATEC na minha empresa':
            chave = 'Propaganda da FATEC, na minha empresa'

        elif chave == 'Eu mesmo procurei por uma instituição de ensino na internet':
            chave = 'Procura própria'
        
        #da pergunta : "Qual sua expectativ após se formar? "
        elif chave == 'Ingressar na carreira acadêmica(se envolver profissionalmente no campo da academia, que engloba atividades relacionadas ao ensino, pesquisa e publicação em instituições educacionais)':
            chave = 'Ingressar na carreira acadêmica'

        
        
        corrigida[chave] = resposta
    
    return corrigida

def corrigir_colunas(pergunta_titulo):
    #Categoria informática
    if pergunta_titulo == 'Windows':
        pergunta_titulo = 'Como classifica seus conhecimentos no Windows?'

    elif pergunta_titulo == 'Linux':
        pergunta_titulo = 'Como classifica seus conhecimentos do Linux?'

    elif pergunta_titulo == 'Editores de Texto (Word, Writer, etc)':
        pergunta_titulo = 'Como classifica seus conhecimentos em Editores de texto?'

    elif pergunta_titulo == 'Planilhas Eletrônicas (Excell, Calc, etc)':
        pergunta_titulo = 'Como classifica seus conhecimentos em Planilhas eletrônicas?'

    elif pergunta_titulo == 'Apresentadores (Powerpoint, Impress, Prezzi, etc)':
        pergunta_titulo = 'Como classifica seus conhecimentos em Softwares de apresentação?'

    elif pergunta_titulo == 'Sistema de Gestão Empresarial':
        pergunta_titulo = 'Como classifica seus conhecimentos em ERPs?'
    

    #Microcomputadores
    if pergunta_titulo == 'Frequência:':
        pergunta_titulo = 'Qual a frequência?'
    
    elif pergunta_titulo in usa_em:
        pergunta_titulo = 'Usa ' + pergunta_titulo


    #Fontes de informação
    if pergunta_titulo == 'Internet':
        pergunta_titulo = 'Quanto você se informa pela Internet?'
    
    elif pergunta_titulo == 'TV':
        pergunta_titulo = 'Quanto você se informa pela TV?'
    
    elif pergunta_titulo == 'Jornais':
        pergunta_titulo = 'Quanto você se informa por Jornais?'
    
    elif pergunta_titulo == 'Revistas':
        pergunta_titulo = 'Quanto você se informa por revistas?'
    
    elif pergunta_titulo == 'Rádio2':
        pergunta_titulo = 'Quanto você se informa pelo rádio?'
    
    elif pergunta_titulo == 'Feed das Redes Sociais (Instagram, Youtube, TikTok, Twitter).':
        pergunta_titulo = 'Quanto você se informa pelas Redes Sociais?'
    
    elif pergunta_titulo == 'Conversas informais com amigos':
        pergunta_titulo = 'Quanto você se informa através de conversas com amigos?'
    
    return pergunta_titulo

def empilhar_coluna(coluna1,coluna2):
    """Função para empilhar colunas duplicadas!
    Parâmetros:
    coluna1 -> primeira coluna a ser empilhada
    coluna2 -> segunda coluna a ser empilhada
    """
    valores = {}
    for resposta in coluna1:
        if isinstance(resposta, str):
            # Verifica se a resposta é uma string vazia ou apenas contém espaços em branco
            if resposta.strip() == '':
                continue  # Ignora células vazias
        elif isinstance(resposta, float):# Verifica se a resposta é NaN
            if math.isnan(resposta):
                continue #Ignora NaN
        
        if resposta not in valores: #após a verificação, começa a adição na lista
            valores[resposta] = 1
        else:
            valores[resposta] +=1


    for resposta2 in coluna2: #não era necessário que a var. de iteração fosse 'resposta2', poderia ser 'resposta' como no primeiro, porém fiz isso para não me confundir!
        if isinstance(resposta2, str):
            #verifica se a resposta é uma string vazia ou apenas contém espaços em branco
            if resposta2.strip() == '':
                continue  #ignora células vazias

        elif isinstance(resposta2, float):#Verifica se a resposta é NaN
            if math.isnan(resposta2):
                continue #ignora NaN

        if resposta2 == 'Tenho e é uma plano individual':
            resposta2 = 'Tenho e é um plano individual.' #corrigindo um erro ortográfico que gera problema no gráfico.
        
        if resposta2 not in valores.keys():
            valores[resposta2] = 1
        else:
            valores[resposta2] +=1

    return valores 



    

