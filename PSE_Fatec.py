from tkinter import *
from tkinter import filedialog
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.widgets import Button as btn
import funcoesImportantes as fi

botao_sair = None #var. global do botao_sair, a qual necessita ser de eixo global para que as funções não percam
#o contexto (e assim o botão pare de funcionar)
botao_proximo = None
botao_anterior = None
temaatual = '' #var. que será usada junto ao aux para contextualização
aux=0 #auxiliar que vai ser responsavel por posição de indice

def proximo(evento):
    global aux
    global temaatual
    global botao_proximo

    botao_proximo = None#"Zerando" o botão próximo
    
    plt.close()
    aux+=1 #ao clicar em prox, a var. aux irá aumentar, fazendo assim pular para a próxima pergunta (a menos que seja a última)
    if aux==(len(temaatual)):
        if 'Você Trabalha atualmente ?' in temaatual:
            trat_nuvem(df_planilha['Qual empresa que você está contratado agora?'])
            #caso esteja na última pergunta d categoria emprego, a nuvem de palavras é aberta!
        else:

            zerar_aux() #rezerando o aux para não dar problemas
        
            reabrir_interface()
    else:
        executar(temaatual)
def anterior(evento):
    global aux
    global botao_anterior
    botao_anterior = None #Zerando o botão "anterior"

    plt.close()
    aux -= 1
    if aux == -1:

        zerar_aux()

        reabrir_interface()
    else:    
        executar(temaatual) #a variável temaatual tem armazenada a categoria que está sendo vista atualmente pelo usuário,
                            #que irá dnv ao executar(), porém dessa vez com o aux maior (ou seja, a próxima pergunta!)
def controle_grafico(evento):
    if evento.key == 'right': #se pressionada a seta direita, executa a função 'proximo'
        proximo(None)
    elif evento.key == 'left': #se pressionanada a seta esquerda, executa a função 'anterior'
        anterior(None)
def fechar_interface():
    janela.withdraw() #fechando a janela temporariamente
def reabrir_interface():
    janela.deiconify()
def zerar_aux():
    global aux
    aux=0
def executar(tema):
    global aux
    global temaatual
    if aux==0: #só vai "fechar" a GUI se for a primeira pergunta 
        
        janela.withdraw() #fechando temporariamente a GUI
    
    temaatual=tema
    
    exibirgrafico(df_planilha[tema[aux]], temaatual) #tema correspondente, e aux corresponde a pos. da pergunta, temaatual mantém o contexto
def sair(evento):                                                                                    
    global botao_sair
    zerar_aux()
    plt.close()
    reabrir_interface()


while True: #loop infinito até upar o arquivo correto dos dados

    arquivo_planilha = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])

    xlsx = pd.ExcelFile(arquivo_planilha)#Abrindo o arquivo xlsx (read only)

    # Verificando se 'Sheet1' está na lista de planilhas
    if 'Sheet1' not in xlsx.sheet_names:
        fi.mensagemErro()
    else:
        df_planilha = pd.read_excel(arquivo_planilha, sheet_name='Sheet1')
        
        colunas = df_planilha.head(0) #pegando os nomes das colunas, para uma segunda verificação

        if 'Planilhas Eletrônicas (Excell, Calc, etc)' in colunas: #uma das colunas presentes no questionario
            break
        else:
            fi.mensagemErro()

#Perguntas por categoria
pergnts_pessoais = [
    'Qual o seu curso ?',
    'Qual o Período que você cursa? ',
    'Em qual estado do Brasil você nasceu ?',
    'Qual a sua cidade de Residência ? (Em qual cidade você mora?)',
    'Qual o seu gênero?',
    'Qual é o seu estado civil ?',
    'Ano de nascimento',
    'Você é portador de alguma deficiência? ',
    'Qual/Quais? (Pode selecionar mais de uma, se for o caso)',
    'Quantos FILHOS você tem ?',
    'Com quem você mora atualmente ?',
    'Quantas pessoas (incluindo você) moram no seu domicílio ?',
    'Qual a situação do domicílio onde mora?',
    'Quanto tempo de moradia nesse domícilio?',
    'Qual a faixa de RENDA MENSAL da sua família (em Salários Mínimos) ?',
    'Você tem plano de saúde privado?',
    'Qual o grau de escolaridade do seu pai?',
    'Qual o grau de escolaridade da sua mãe?',
    'Na sua vida escolar, você sempre estudou:',
    'Não considerando os livros escolares, quantos livros você lê por ano (em média)?',
    'Você lê jornal? ',
    'Se você lê livros literários, qual o gênero preferido?',
    'Você dedica parte do seu tempo para atividades voluntárias?',
    'Se você frequenta alguma religião, qual religião você frequenta?',
    'Você costuma utilizar de fontes de entretenimento cultural?',
    'Com qual frequência você lê jornal?',
    'Como conheceu a FATEC FRANCA?',
    'Qual sua expectativa após se formar ',
    'Qual sua maior expectativa quanto ao curso?',
    'Você já fez algum curso técnico?',
    'Qual o meio de transporte você usa para vir à escola ?'
]
pergnts_casa = [
    'Televisor',
    'Video Cassete e/ou DVD',
    'Rádio',
    'Automóvel (Carro)',
    'Motocicleta (moto)',
    'Máquina de Lavar Roupa e (ou) tanquinho',
    'Geladeira e (ou) Freezer para alimentos',
    'Smartphone (ANDRÓID ou IOS).',
    'Computador de mesa (desktop)',
    'Notebook (laptop)',
    'Telefone Fixo',
    'Internet pessoal',
    'TV por assinatura (TV via satélite com programação fixa)',
    'Empregada Mensalista',
    'Plataformas de Streaming (Netflix, HBO Max, Amazon Prime, etc...)'
]
pergnts_emprego = [
    'Você Trabalha atualmente ?',
    'Qual seu vínculo com o emprego?',
    'Qual a área do seu trabalho',
    'Qual seu horário de trabalho?']
pergnts_microcomputador = [
    'Você usa microcomputadores? ',
    'Frequência:',
    'Em casa:',
    'No trabalho:',
    'Na escola:',
    'Em outros locais:',
    'Para trabalhos profissionais:',
    'Para trabalhos escolares:',
    'Para entretenimento (músicas, vídeos, redes sociais, etc)',
    'Para comunicação por email:',
    'Para operações bancárias:',
    'Para compras na internet:']
pergnts_informatica = [
    'Windows',
    'Linux',
    'Editores de Texto (Word, Writer, etc)',
    'Planilhas Eletrônicas (Excell, Calc, etc)',
    'Apresentadores (Powerpoint, Impress, Prezzi, etc)',
    'Sistema de Gestão Empresarial',
    'Como você classifica seu conhecimento em informática?']
pergnts_idiomas = [
    'Qual seu nível no idioma Inglês?',
    'Qual seu nível no idioma Espanhol?',
    'Fala algum outro idioma?',
    'Qual outro idioma?',
    'Qual seu nível neste idioma?']
pergnts_fontinf = [
    'Internet',
    'TV',
    'Jornais',
    'Revistas',
    'Rádio2',
    'Feed das Redes Sociais (Instagram, Youtube, TikTok, Twitter).',
    'Conversas informais com amigos']
assuntos_le = df_planilha['Quais os assuntos que mais lê?']
escolheu_curso = df_planilha['Por que escolheu este curso?']
sonhos = df_planilha['Escreva algumas linhas sobre sua história e seus sonhos de vida.']
ent_cultural = df_planilha['Quais fontes de ENTRETENIMENTO CULTURAL você usa?*']

#perguntas que precisam de correção ortográfica ou alteração:
perguntas_a_corrigir = ['Como você classifica seu conhecimento em informática?',
    'Windows',
    'Linux',
    'Editores de Texto (Word, Writer, etc)',
    'Planilhas Eletrônicas (Excell, Calc, etc)',
    'Apresentadores (Powerpoint, Impress, Prezzi, etc)',
    'Sistema de Gestão Empresarial',
    'Qual a faixa de RENDA MENSAL da sua família (em Salários Mínimos) ?',
    'Frequência:',
    'Em casa:',
    'No trabalho:',
    'Na escola:',
    'Em outros locais:',
    'Para trabalhos profissionais:',
    'Para trabalhos escolares:',
    'Para entretenimento (músicas, vídeos, redes sociais, etc)',
    'Para comunicação por email:',
    'Para operações bancárias:',
    'Para compras na internet:',
    'Internet',
    'TV',
    'Jornais',
    'Revistas',
    'Rádio2',
    'Feed das Redes Sociais (Instagram, Youtube, TikTok, Twitter).',
    'Conversas informais com amigos'
    ]

def exibirgrafico(pergunta_dados, temaatual):#função para tratamento das perg. fechadas, o parametro de tema auxiliará no avanço ou retrocesso dos gráfico
    global botao_sair
    global botao_proximo
    global botao_anterior
    global aux

    #vendo qual é a pergunta
    pergunta_titulo = temaatual[aux]

    fig,ax = plt.subplots() #criando os eixos e a figura do gráfico

    cores_cps = ['#FF4C4D','#A0C340', '#4C7EFF','#FFC24C','#8A29E6','#FF4CA2','#00D8E8','#42C36C',
                '#E945F0','#F99B45'] #paleta de cores do CPS (Cores auxiliares)

    #Conferindo se os dados precisam ser tratados antes de virar gráfico
    if pergunta_titulo == 'Quanto tempo de moradia nesse domícilio?':
        valores = fi.processar_moradia(pergunta_dados)
    elif pergunta_titulo == 'Ano de nascimento':
        valores = fi.processar_idade(pergunta_dados)
        pergunta_titulo = 'Idade dos entrevistados'
    else:
        valores = fi.filtrarFechada(df_planilha[pergunta_titulo])
    
    #empilhando respostas de colunas duplicadas:
    if pergunta_titulo == 'Rádio2':
        valores = fi.empilhar_coluna(df_planilha[pergunta_titulo],df_planilha['Rádio3'])
    elif pergunta_titulo == 'Revistas':
        valores = fi.empilhar_coluna(df_planilha[pergunta_titulo],df_planilha['Revistas2'])
    elif pergunta_titulo == 'Feed das Redes Sociais (Instagram, Youtube, TikTok, Twitter).':
        valores = fi.empilhar_coluna(df_planilha[pergunta_titulo],df_planilha['Feed das Redes Sociais (Instagram, Youtube, TikTok, Twitter).2'])
    elif pergunta_titulo == 'Internet':
        valores = fi.empilhar_coluna(df_planilha[pergunta_titulo],df_planilha['Internet2'])
    elif pergunta_titulo == 'TV':
        valores = fi.empilhar_coluna(df_planilha[pergunta_titulo],df_planilha['TV2'])
    elif pergunta_titulo == 'Conversas informais com amigos':
        valores = fi.empilhar_coluna(df_planilha[pergunta_titulo],df_planilha['Conversas informais com amigos2'])
    elif pergunta_titulo == 'Você tem plano de saúde privado?':
        valores = fi.empilhar_coluna(df_planilha[pergunta_titulo],df_planilha['Você tem plano de saúde privado?2'])
    
    #tratando erros ortográfico e reduzindo enunciados grandes EM RESPOSTAS
    valores = fi.corrigir_respostas(valores)
    
    #tratando erros ortográfico e reduzindo enunciados grandes EM PERGUNTAS (Enunciados)
    if pergunta_titulo in perguntas_a_corrigir:
        pergunta_titulo = fi.corrigir_colunas(pergunta_titulo)

    ax.pie(
        valores.values(),
        labels=valores.keys(),
        autopct=fi.formatar_porcentagem(valores.values()),
        startangle=140,
        pctdistance=0.85,
        colors=cores_cps
    )
    ax.axis('equal')
    ax.set_title(pergunta_titulo)
    fig.set_facecolor('#DADADA')

    fig.canvas.mpl_connect('key_press_event',controle_grafico) #detectando click das setas, para então dar trigg no evento
    fig.canvas.manager.full_screen_toggle() #deixa em tela cheia
    
    #                        CRIANDO um botão "Sair" na janela
    
    x_coord_cntesquerdo = 0.05  #coordenadas x e y para ficar no cnt superior esquerdo
    y_coord_cima = 0.85 
    alt_botao = 0.05     #<-- tamanho do botão
    larg_botao = 0.1
    botao_sair = btn(fig.add_axes([x_coord_cntesquerdo, y_coord_cima, larg_botao, alt_botao]),'Voltar ao menu', 
                    color='#ef0404',hovercolor='#9c0c0c')
    botao_sair.on_clicked(sair)


    #anotações de "avançar e voltar"
    x_voltar = 0.65
    x_avancar = 8.7      #coordenadas x e y das anotações
    y_cnt_inferior = -16

    plt.annotate('Avançar', xy=(x_avancar, y_cnt_inferior), xycoords='axes fraction', fontsize=12,
            horizontalalignment='right', verticalalignment='top') #texto anotado ao lado do gráfico
    plt.annotate('Voltar', xy=(x_voltar, y_cnt_inferior), xycoords='axes fraction', fontsize=12,
            horizontalalignment='left', verticalalignment='top')
    plt.annotate('Use as setas ou clique nos botões para passar os gráficos!', xy=(4.7,-17), xycoords='axes fraction',
                 fontsize=12, horizontalalignment='center',verticalalignment='bottom',color='#7E0000', weight='bold') #tutorial, presente no centro inferior.

    x_setadireita = 0.875
    x_setaesquerda = 0.11       #coordenadas x e y dos botões
    y_cnt_inferior_img = 0.06
    alt_seta = 0.05
    larg_seta = 0.04          #tamanho das setas
    
    #colocando as imagens no gráfico
    botao_proximo = btn(fig.add_axes([x_setadireita,y_cnt_inferior_img, larg_seta,alt_seta]),'→',color='#9a9fa2',
                        hovercolor='#666666')
    botao_proximo.on_clicked(proximo)
    
    botao_anterior = btn(fig.add_axes([x_setaesquerda,y_cnt_inferior_img,larg_seta,alt_seta]),'←',color='#9a9fa2',
                         hovercolor='#666666')
    botao_anterior.on_clicked(anterior)

    plt.show() #mostra o gráfico

def trat_nuvem(dados_nuvem): #função de tratamento das abertas

    from wordcloud import WordCloud as wc #função para criar uma nuvem de palavras ;
    from nltk.corpus import stopwords as sw # bibl. para contar palavras em texto ;
    from nltk.tokenize import word_tokenize as wt #tokenização de palavras ; 
    from collections import Counter #função contadora , semelhante ao ntlk porém mais simples.
    
    global botao_sair
    global aux  #importando botoes do esc. global para que não se perca o contexto
    global botao_proximo
    global botao_anterior


    dados_filtrados = fi.filtrarNuvem(dados_nuvem)
    titulo_pergunta = dados_nuvem.name #nome da coluna

    if titulo_pergunta != 'Qual empresa que você está contratado agora?':
        janela.withdraw() #fechar a interface temporariamente (a menos que seja a nuvem de empresas)
        
    #Perguntas onde era possível mais de uma resposta também são tratadas em nuvem. (As repostas múltiplas são separadas por ponto e vírgula)
    trat_pontovirgula = ['Quais fontes de ENTRETENIMENTO CULTURAL você usa?*', 'Quais os assuntos que mais lê?', 'Por que escolheu este curso?']

    fig,ax = plt.subplots()

    if titulo_pergunta in trat_pontovirgula:
        #tratamento das nuvens de multipla escolha (possuem ponto e virgula)
            
            dados_filtrados = fi.filtrarNuvem(dados_nuvem)

            texto_completo = ';'.join(dados_filtrados)  #juntando todas as respostas em uma única string
            frases = texto_completo.split(';')  #separanto tudo em frases (com base no ponto e vírgula)
            frases = [frase.strip() for frase in frases]  #removendo espaços em branco em excesso(gerados pela separação em ' ; ')

            contador_frases = Counter(frases)  #usando Counter para contar a frequência das frases

            # Teste para escolher o titulo 
            if titulo_pergunta == 'Quais os assuntos que mais lê?':
                titulo_nuvem = 'Os Assuntos Jornalísticos mais lidos!' #titulo_nuvem refere-se ao titulo da nuvem de palvras
            elif titulo_pergunta == 'Por que escolheu este curso?':
                titulo_nuvem = '"Porquê você escolheu este curso?"' 
            elif titulo_pergunta == 'Quais fontes de ENTRETENIMENTO CULTURAL você usa?*':
                titulo_nuvem = 'As fontes culturais mais consumidas!'

            nuvem_palavras = wc(width=800, height=600, background_color='white').generate_from_frequencies(
                contador_frases)
    
    else:
#                       PROCESSANDO OS DADOS - Versão sem ponto e virgula
        texto_completo = ' '.join(dados_filtrados)
        palavras = wt(texto_completo.lower()) #wt separa o texto palavra por palavra, aqui eu tbm coloquei em minusculo
                                            #para não dar nenhuma incongruência
        
        palavras_sem_stopword = [palavra for palavra in palavras if palavra in palavra not in sw.words('Portuguese')]
        # remoção de stopwords (do nosso idioma) -> 'stopwords' são palavras muito comuns, como artigos e preposições              
        # e por isso não devem ser contabilizadas.     Nesta linha eu ja declarei a variável palavra de uma vez e usei ela
        # nesse loop de for das palavras.
        

        #Definindo titulo
        if titulo_pergunta == 'Escreva algumas linhas sobre sua história e seus sonhos de vida.':
            titulo_nuvem = 'As palavras mais presentes nos nossos sonhos!'
        else:
            titulo_nuvem = 'Nuvem de palavras: empresas que empregram fatecanos!'
        
        #                   CRIANDO A NUVEM DE PALAVRAS
        nuvem_palavras = wc(width=800, height=600, background_color='white').generate(' '.join(palavras_sem_stopword))
        #esta é a variavel da nuvem de palavras, primeiro voce deve dar a resolução dela e a cor de fundo, para depois
        #passar o texto a ser usado (com o .generate), outra coisa importante é usar o ' ' (espaço) com join,
        #se não as palvras ficam sem espaço umas das outras.

    #                        CRIANDO um botão "Sair" na janela
    
    x_coord_cntesquerdo = 0.05  #coordenadas x e y para ficar no cnt superior esquerdo
    y_coord_cima = 0.85 
    alt_cancelar = 0.05     #<-- tamanho do botão
    larg_cancelar = 0.1
    botao_sair = btn(fig.add_axes([x_coord_cntesquerdo, y_coord_cima, larg_cancelar, alt_cancelar]),'Voltar ao menu', 
                    color='#ef0404',hovercolor='#9c0c0c')
    botao_sair.on_clicked(sair)

    #pegando a nuvem e colocando numa janela do pyplot (ax)
    ax.imshow(nuvem_palavras, interpolation='bilinear') #comando sobre oq será exibido na janela do plot
    ax.axis('off') #desativa os eixos
    ax.set_title(titulo_nuvem) #coloca o titulo

    fig.canvas.manager.full_screen_toggle()
    plt.show()

#                                   CRIANDO INTERFACE GRÁFICA

janela = Tk() #criando a janela;
janela.title('PSE GRUPO IV - ADS II Noturno 23/2')# colocando o titulo;

imagem_fundo = PhotoImage(file='_imagens\\background.png') #carregando as imagens de fundo;
imagem_pessoais = PhotoImage(file='_imagens\\botao_perguntaspessoais.png')
imagem_culturais = PhotoImage(file='_imagens\\botao_fontesculturais.png')
imagem_domiciliares = PhotoImage(file='_imagens\\botao_domiciliares.png')
imagem_informatica = PhotoImage(file='_imagens\\botao_informatica.png')
imagem_sonhos = PhotoImage(file='_imagens\\botao_sonhos.png')
imagem_microcomputadores = PhotoImage(file='_imagens\\botao_microcomputadores.png')
imagem_fontesinformacao = PhotoImage(file='_imagens\\botao_fontesinformacao.png')
imagem_noticias = PhotoImage(file='_imagens\\botao_noticias.png')
imagem_emprego = PhotoImage(file='_imagens\\botao_emprego.png')
imagem_idiomas = PhotoImage(file='_imagens\\botao_idiomas.png')
imagem_escolheucurso = PhotoImage(file='_imagens\\botao_curso.png')

canvas = Canvas(janela, width=imagem_fundo.width(), height=imagem_fundo.height()) #criando um canvas na janela, com a mesma resolução da imagem;
canvas.pack()
canvas.create_image(0,0, anchor=NW, image=imagem_fundo) #coloca a imagem no canvas, com posição padrão (0,0) e ancoragem padrão (Northwest);

#BOTÕES DO MENU:
x_colunaesquerda = 130  #Variaveis de coordenada
x_colunadireita = 360
y_linha1 = 175
y_linha2 = (y_linha1 + 80)
y_linha3 = (y_linha2 + 80)      #   "padding" de 80px pra cada botão
y_linha4 = (y_linha3 + 80)      
y_linha5 = (y_linha4 + 80)
y_linha6 = (y_linha5 + 65)      #exceto o último, que é um botão menor
x_cental = 245

botao1 = Button(janela, text='PERGUNTAS PESSOAIS',command=lambda: executar(pergnts_pessoais))#<-- definindo o botão... lambda faz com que ele só seja executado ao clicar;
canvas.create_window(x_colunaesquerda,y_linha1,window=botao1) #posição do botão valor de x é o quanto ele está para direita, 
botao1.config(image=imagem_pessoais)                     # y é para baixo

botao2 = Button(janela, text='FONTES DE INFORMAÇÃO',command=lambda: executar(pergnts_fontinf))
canvas.create_window(x_colunadireita,y_linha1,window=botao2) #os dois primeiros parametros sao cord x e coord y,
botao2.config(image=imagem_fontesinformacao)            #o terceiro é falando que a "janela" a ser criada na GUI é o botao

botao3 = Button(janela, text='NUVEM DE SONHOS',command=lambda: trat_nuvem(sonhos))
canvas.create_window(x_colunaesquerda,y_linha2,window=botao3)
botao3.config(image=imagem_sonhos)

botao4 = Button(janela, text='MICROCOMPUTADOR',command=lambda: executar(pergnts_microcomputador))
canvas.create_window(x_colunadireita,y_linha2,window=botao4)
botao4.config(image=imagem_microcomputadores)

botao5 = Button(janela, text='PERGUNTAS DOMICILIARES',command=lambda: executar(pergnts_casa))
canvas.create_window(x_colunaesquerda,y_linha3,window=botao5)
botao5.config(image=imagem_domiciliares)

botao6 = Button(janela, text='ENTRETENIMENTO FAVORITO',command=lambda: trat_nuvem(ent_cultural))
canvas.create_window(x_colunadireita,y_linha3,window=botao6)
botao6.config(image=imagem_culturais)

botao7 = Button(janela, text='CONHECIMENTOS EM INFORMÁTICA', command= lambda: executar(pergnts_informatica))
canvas.create_window(x_colunaesquerda,y_linha4,window=botao7)
botao7.config(image=imagem_informatica)

botao8 = Button(janela, text='NOTÍCIAS MAIS LIDAS',command=lambda: trat_nuvem(assuntos_le))
canvas.create_window(x_colunadireita,y_linha4,window=botao8)
botao8.config(image=imagem_noticias)

botao9 = Button(janela, text='EMPREGO', command=lambda: executar(pergnts_emprego))
canvas.create_window(x_colunaesquerda,y_linha5,window=botao9)
botao9.config(image=imagem_emprego)

botao10 = Button(janela,text='IDIOMAS', command=lambda: executar(pergnts_idiomas))
canvas.create_window(x_colunadireita,y_linha5,window=botao10)
botao10.config(image=imagem_idiomas)

botao11 = Button(janela,text='CURSO', command=lambda: trat_nuvem(escolheu_curso))
canvas.create_window(x_cental,y_linha6,window=botao11)
botao11.config(image=imagem_escolheucurso)


janela.mainloop() #abre a janela (GUI)