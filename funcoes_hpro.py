import pyautogui as p
from time import sleep
import pandas as pd
import subprocess
from openpyxl import Workbook, load_workbook




# Funçao para hotkeys
def hotkeys(x, z=''):
    y = p.hotkey(x,z)
    sleep(0.5)
    return y

#################################### ABRINDO HPRO e FAZENDO LOGIN ####################################
# Abrir HPRO
def abrir_hpro():
    abrir = subprocess.Popen("C:\ClientTeste\PH2.EXE")
    sleep(15) # Tempo para carregar HPRO
    return abrir


# Mover, Clicar, Escrever -> Usuário e Senha e apertar Enter // IMUTÁVEL
def cadastrar_usuario_senha_apertar_enter():
    p.moveTo(866,559)
    p.click(866,559)
    p.typewrite("matos")
    hotkeys('tab')
    p.typewrite("123456")
    hotkeys('enter')
    sleep(2)
    
#####################################################################################################



#################################### DENTRO DO SITEMA ####################################
def requisitar_item_estoque():
    # Mover, Clicar -> Estoque
    # def mover_clicar_estoque():
    p.moveTo(223, 33)
    p.click(223,33)
    # mover_clicar_estoque()

    # Mover, Clicar -> Requisição Estoque
    # def mover_clicar_requisiçãoEstoque():
    p.moveTo(295, 62)
    p.click(295,62)
    # mover_clicar_requisiçãoEstoque()

    # Mover e Clicar -> Manutenção
    # def mover_clicar_manutenção():
    p.moveTo(566, 56)
    p.click(566,56)
    # mover_clicar_manutenção()

    # Abrindo banco de dados para armazenar len(cod)
    arq = pd.read_excel('dados para req estoque.xlsx')
    cod = arq['CODIGO']
    y = 0
    while y < len(cod):  # Vai repetir até chegar ao final do len(cod)

        # Apertar 'i' para -> Incluir
        hotkeys('i')

        # Escrever solicitante
        # def escrever_solicitante():
        p.typewrite("Rafael Matos")
        # escrever_solicitante()

        # Apertar Tab para ir de Solicitante para -> Departamento
        hotkeys('tab')

        # Escrever Departamento
        p.typewrite("3,0110")


        ################################# ACESSANDO BANCO DE DADOS #################################
        # Pegar dado da planilha e passar para as variaveis codigos e quantidade
        codigos = []  # Codigos a serem requisitados
        quantidade = []  # Quantidade a requisitar

        arq = pd.read_excel('dados para req estoque.xlsx')
        cod = arq['CODIGO']
        for cell in cod:
            codigos.append(str(cell))

        arq = pd.read_excel('dados para req estoque.xlsx')
        qnt = arq['QUANTIDADE']
        for cell in qnt:
            quantidade.append(str(cell))
        ###############################################################################################


        #################################### CASO EU QUERIA MOSTRAR OS DADOS RECOLHIDOS ####################################
        # data = {
        # "codigos": codigos,
        # "quantidade": quantidade
        # }

        # #load data into a DataFrame object:
        # df = pd.DataFrame(data)
        # print(df)

        ####################################################################################################################


        #################################### PASSANDO DADOS RECOLHIDOS DAS VARIAVEIS PARA A HPRO ####################################
        # Tab para sair de Departamento -> Código do produto a requisitar
        hotkeys('tab')
        p.typewrite(codigos[y], interval=0.2)

        # Mover, Clicar e Escrever -> Quantidade
        # def mover_clicar_escrever_quantidade():
        p.moveTo(842, 473, duration=0.2)
        p.doubleClick(872,473,duration=0.2)
        p.typewrite(quantidade[y], interval=0.2)
        # mover_clicar_escrever_quantidade()

        # Mover, Clicar e Escrever -> Observção (Se tiver)
        # def mover_clicar_escrever_observação():
        p.moveTo(912, 593, duration=0.2)
        p.click(912,593,duration=0.2)
        p.typewrite('ITEM ESTA NO KANBAN', interval=0.2)
        # mover_clicar_escrever_observação()

        hotkeys('alt', 'g')
        sleep(0.5)

        y += 1    


def requisitar_item_compra():
    p.moveTo(303, 33, duration=0.1)  # Mover para suplimentos
    p.click(303, 33, duration=0.1)  # Clicar em suplimentos

    p.moveTo(353, 154, duration=0.1)  # Mover para requisição
    p.click(353, 154, duration=0.1)  # Clicar em requisição 

    p.moveTo(608, 151, duration=0.1)  # Mover para digitação
    p.click(608, 151, duration=0.1)  # Clicar em digitação

    hotkeys('i')


    arq = pd.read_excel('dados para req compras.xlsx')
    cod = arq['CODIGO']
    x = 0
    while x < len(cod):

        hotkeys('i')

        codigos = []  # Codigos a serem requisitados
        quantidade = []  # Quantidade a requisitar
        data_necessidade = []
        projetos = [] 
        observação = []

        arq = pd.read_excel('dados para req compras.xlsx')
        data = arq['DATA NECESSIDADE']
        for cell in data:
            data_necessidade.append(str(cell))

        arq = pd.read_excel('dados para req compras.xlsx')
        cod = arq['CODIGO']
        for cell in cod:
            codigos.append(str(cell))

        arq = pd.read_excel('dados para req compras.xlsx')
        qnt = arq['QUANTIDADE']
        for cell in qnt:
            quantidade.append(str(cell))

        arq = pd.read_excel('dados para req compras.xlsx')
        proj = arq['PROJETOS']
        for cell in proj:
            projetos.append(str(cell))

        arq = pd.read_excel('dados para req compras.xlsx')
        obs = arq['OBSERVACAO']
        for cell in obs:
           observação.append(str(cell))


        p.typewrite(data_necessidade[x]) # Data de Necessidade

        hotkeys('tab')

        p.typewrite('Rafael Matos') # Solicitante

        hotkeys('tab')

        p.typewrite('3,0110') # Departamento

        hotkeys('tab')

        p.typewrite(codigos[x]) # Código

        hotkeys('tab')

        p.typewrite(quantidade[x]) # Qauntidade
 
        hotkeys('alt', 'i') # Para ir para Incluir Projeto
        p.typewrite(projetos[x])
        hotkeys('alt', 'g') # Apertar Gravar o Projeto
        
        p.moveTo(1065, 721) # Mover para Campo de Observção
        p.click(1065, 721) # Clicar no Campo de Observação
        p.typewrite(observação[x]) # Escrever no Campo de Observação

        hotkeys('alt', 'g')

        x += 1

        
def alterar_localização_estoque():
    hotkeys('alt', 'c')
    hotkeys('m')
    hotkeys('enter')

    p.moveTo(43, 290)
    p.click(43, 290)
    
    arq = pd.read_excel('alteração de localização.xlsx')
    cod = arq['CODIGO']

    

    




