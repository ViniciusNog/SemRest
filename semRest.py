#!/usr/bin/env python3
from datetime import date, datetime
import time
from time import sleep # usar 'sleep(value)' ao invés de 'time.sleep(value)'
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
import sys, os
import pymssql
import pyautogui
import importlib
from db import *

nome_consultor = 'samuel'

# funcao de acessar frames
def acessarframe(string,frame):
    driver.switch_to.default_content()
    for x in range(0,frame):
        driver.switch_to.frame(string)

def ElementoExiste(xpath, elementToSearch):
    try:
        elementToSearch.find_element_by_xpath(xpath)
        return True
    except:
        return False

# caminho no banco para acessar o pdf e a concatenação com o nome do arquivo
caminho_arquivo = ('C:/WEBSITES/MDS/bmss_zurich/SolicitacaoEndosso/PDFs/' + nome_arq)

def fechaAlert():
    driver.switch_to.default_content()
    alerta_registro = ''
    
    try:
        xpath = 'body > div.swal2-container.swal2-center.swal2-fade.swal2-shown'
        alerta_registro = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, xpath)))
    except:
        pass

    if(alerta_registro != ''):
        #print('apareceu')
        xpath = '/html/body/div[3]/div/div[2]/div[1]/div/div/button[2]'
        btn_fechar = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
   
# dados para preencher os campos obrigatórioss
dados = {
    "nu_rs": "1234",
    "num_proposta": "%7",
    "nome_cliente" : 'PERBRAS EMP BRAS DE PERFURACOES LTDA',
    "fim_vigencia" : '31/05/2027',
    "seguradora": 'AUSTRAL SEGURADORA',
    "tipo_endosso": 1,
    "sub_tipo_endosso": 1.1,
    "inicio_vigencia": '01/06/2021',
    "proposta_cia": '6431228/19',
    "premios": '000000',
    "valor_estorno" : '150,00', # 150, 00? não sei quanto deve ser o valor necessário para preencher
    "premio_liquido": '200,00',
    "descricao_arq": 'Proposta de endosso'
}

# fazer o login na plataforma quiver sem o pyautogui:

driver = webdriver.Chrome(executable_path=r'chromedriver.exe')

wait = WebDriverWait(driver, 20, poll_frequency=5)

driver.maximize_window()

#busca = 'http://svc.itpowerhom:og\hEVuckt@quiver.hom.mdsinsure.com' # acesso para homologação
busca = 'http://svc.itendosso:WTR8nbf!NF@quiver.hom.mdsinsure.com' # acessa o site desejado(quiver) já fazendo o login automático

driver.get(busca)
sleep(2)

# abertura da Nav bar

# driver.execute_script('openNav()')

# driver.execute_script("SelecionaModuloJQuery('FrmPortal.aspx?pagina=Operacional','OPERACIONAL','Professional|AcompanhamentoVendas|MultiNovo','OPERACIONAL','Operacional');")

driver.execute_script("SelecionaModuloJQuery('ConsultaEmissaoERecusa;Fast/FrmCadastroNovo.aspx?pagina=Documento','EMISSOESRECUSAS','Professional','EMISSOESRECUSAS','Propostas/Apólices'); ")

driver.switch_to.frame('ZonaInterna')

# xpath = '//*[@id="DIVTipoConsulta2"]/div/span/span[1]/span/span[2]'
# clica nome apólice
xpath = '/html/body/form/div[5]/div[1]/div/div/div[2]/div[1]/div[2]/div/div/span/span[1]/span/span[1]'
select_campo = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
sleep(2)
xpath = '/html/body/span[2]/span/span[2]/ul/li[2]'
select_apolice = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
sleep(2)

xpath = '/html/body/form/div[5]/div[1]/div/div/div[2]/div[1]/div[4]/div[1]/div[1]/div/input'
num_proposta = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
num_proposta.send_keys(nuapolice)
sleep(2)

xpath = '//*[@id="SpanToolBarRightB"]/button'
btn_pesquisar = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
sleep(2)

driver.switch_to.default_content()
alerta_registro = ''

try:
    xpath = 'body > div.swal2-container.swal2-center.swal2-fade.swal2-shown'
    alerta_registro = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, xpath)))
except:
    pass

# caso apareca um alerta de registro ele fecha automaticamente
if(alerta_registro != ''):
    print('apareceu')
    xpath = '/html/body/div[3]/div/div[3]/button[1]'
    btn_fechar = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()

driver.switch_to.frame('ZonaInterna')
sleep(2)

xpath = '/html/body/form/div[5]/div[2]/div[2]/div/div[3]/div[3]/div/table/tbody/tr[2]/td[11]'
numeroapolice = driver.find_element_by_xpath(xpath).get_attribute("innerHTML")
print(f'numeroapolice: {numeroapolice}')
sleep(2)

xpath = '/html/body/form/div[5]/div[2]/div[2]/div/div[3]/div[3]/div/table/tbody/tr[2]/td[10]'
fim_apolice = driver.find_element_by_xpath(xpath).get_attribute("innerHTML")
print(f'fim apolice: {fim_apolice}')
sleep(2)

xpath = '/html/body/form/div[5]/div[2]/div[2]/div/div[3]/div[3]/div/table/tbody/tr[2]/td[4]'
nome_seguradora = driver.find_element_by_xpath(xpath).get_attribute("innerHTML")
print(f'num seguradora: {nome_seguradora}')
sleep(2)

xpath = '/html/body/form/div[5]/div[2]/div[2]/div/div[3]/div[3]/div/table/tbody/tr[2]/td[15]'
checar_emitida = driver.find_element_by_xpath(xpath).get_attribute("innerHTML")
print(f'checar: {checar_emitida}')
sleep(2)

checar_emitida1 = 'Emitida'

if(numeroapolice != '' ) and (fim_apolice == fimVigencia) and (nome_seguradora == seguradora) and(checar_emitida == checar_emitida1): # and checar_emitida = ''
    #print('tudo ok')
    xpath = '//*[@id="BtEdiReg"]'
    #xpath = '//*[@id="2882968,0,1"]/td[1]'
    btn_editar = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
    sleep(2)
    
    driver.switch_to.frame('ZonaInterna')
    driver.execute_script('AbreEndosso()')

    driver.switch_to.default_content()
    alerta_registro = ''

    try:
        xpath = 'body > div.swal2-container.swal2-center.swal2-fade.swal2-shown'
        alerta_registro = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, xpath)))
    except:
        pass
    
    sleep(2)
    if(alerta_registro != ''):
        #print('apareceu')
        # fechar o alerta de registro da apólice
        xpath = '/html/body/div[3]/div/div[3]/button[1]'
        #xpath = '/html/body/div[4]/div/div[3]/button[1]'
        btn_fechar = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()

    driver.switch_to.frame('ZonaInterna')
    driver.switch_to.frame('ZonaInterna')
    sleep(2)

    if(dados['tipo_endosso'] == 2) and (dados['sub_tipo_endosso'] == 1.1):
            #print("SOLICITACAO DO SEGURADO")
            #print('entrou no novo indosso')
            acessarframe('ZonaInterna',2)
            driver.switch_to.frame('Documento')
            #print("SOLICITACAO DO SEGURADO")
    
            # tipo do documento
            StatusRPA = '1 - Preenchimento dos dados'
            xpath = '/html/body/form/div[7]/div/div[2]/div[1]/div/div[1]/div[2]/div[3]/div[1]/div/span/span[1]/span'
            #xpath = '//*[@id="select2-Documento_TipoDocumento-container"]'
            select_campo_tipo = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
            sleep(2)
    
            xpath = '/html/body/span[2]/span/span[2]/ul/li[2]'
            select_campo_tipo_doc = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
            sleep(2)
            
            #xpath = '//*[@id="DIVDocumento_SubTipo"]/div/span/span[1]/span'
            
            # subtipo do documento
            xpath = '//*[@id="select2-Documento_SubTipo-container"]'
            select_campo_sub = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
            sleep(2)

            xpath = '/html/body/span[2]/span/span[2]/ul/li[2]'
            select_solicitacao = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
            sleep(2)
            
            # dados de inicio de vigencia
            xpath = '//*[@id="Documento_InicioVigencia"]'
            inicio_vigencia = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
            inicio_vigencia.clear()
            inicio_vigencia.send_keys(inicioVigencia)
            sleep(2)

            # inserindo dados da companhia e em seguida já fazendo o try/except caso apareca um alerta aleatorio
            xpath = '/html/body/form/div[7]/div/div[2]/div[1]/div/div[1]/div[2]/div[7]/div[4]/div/input'
            proposta_cia = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
            proposta_cia.send_keys(nupropostacia)
            sleep(2)

            # abre aba de dados complementares
            xpath = '//*[@id="TitDadosComplementares"]'
            dados_complementares = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
            sleep(2)

            # campo de observação adicionando o valor do nu_rs como esta de acordo no diagrama
            #xpath = '//*[@id="Documento_Texto1"]' só a xpath, sem o full. caso necessario:
            xpath = '/html/body/form/div[7]/div/div[2]/div[2]/div/div[2]/div[2]/div/div/textarea'
            campo_observacao = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
            campo_observacao.clear()
            campo_observacao.send_keys(rs)

            # xpath = '/html/body/form/div[5]/div/button[3]'
            xpath = '//*[@id="BtGravar"]'
            btn_gravar = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
            sleep(2)
            btn_gravar.send_keys(Keys.ENTER)

            xpath = '//*[@id="swalbtn1"]'
            alerta_registro = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
            
            try:
                #acessarframe('ZonaInterna',2)
                xpath = '//*[@id="swalbtn1"]'
                alerta_registro = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, xpath))).click()
            except:
                pass

            # abre a aba premios
            StatusRPA = '2 - Premios'
            xpath = '//*[@id="TitPremios"]' 
            premios = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
       
            # xpath = '//*[@id="select2-Documento_TipoDocumento-result-ykk4-5"]'
            xpath = '//*[@id="Documento_PremioLiquido"]'
            # premios = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
            premios = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
            premios.send_keys('0,00')
            sleep(2)


            try:
                # xpath = '/html/body/form/div[5]/div/button[3]'
                xpath = '//*[@id="BtGravar"]'
                btn_gravar = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
                sleep(2)
                btn_gravar.send_keys(Keys.ENTER)
            except:
                driver.execute_script("eventoAjax('Gravar');")

            xpath = '//*[@id="BtGravar"]'
            btn_gravar = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
            sleep(2)
            btn_gravar.send_keys(Keys.ENTER)


            # acessar aba arquivos -> em descrição: proposta de endosso -> em tipo de imagens: endosso -> carergar o pdf -> gravar
            StatusRPA = '2.1 - Arquivos'
            xpath = '//*[@id="BtAnexar"]/div[2]/div[2]/a/span'  # full xpath: /html/body/form/div[7]/div/div[1]/ul/li[2]/div[2]/div[2]/a/span
            aba_arquivo = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()

            # sleep(2)

            # descrição: proposta de endosso
            xpath = '//*[@id="ScanImagem_Descricao"]'
            descricao_arquivo = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
            descricao_arquivo.clear()
            descricao_arquivo.send_keys('Proposta de endosso')
            
            sleep(2)

            # tipo de imagem: endosso
            xpath = '//*[@id="select2-ScanImagem_TipoImagem-container"]'
            select_imagem = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
            sleep(2)
            xpath = '//*[@id="select2-ScanImagem_TipoImagem-result-k0id-16"]'
            select_tipo_imagem = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
            sleep(2)
            
            # carregar pdf 
            # xpath do botão carregar arquivo caso necessário:
            # /html/body/form/div[3]/div/div/span/div/div/div[2]/div/div[1]/div[3]/span/div/div/table/tbody/tr/td[2]/div/input[2]
            filepath = driver.find_element(by = By.ID, value = 'file')
            filepath.send_keys(caminho_arquivo)

            xpath = '/html/body/div[5]/div/div[3]/button[1]'
            click_ok = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
            sleep(2)

            try:
                # xpath = '/html/body/form/div[5]/div/button[3]'
                xpath = '/html/body/div[5]/div/div[3]/button[1]'
                btn_fechar = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
                sleep(2)
                btn_fechar.send_keys(Keys.ENTER)
            except:
                pass

            # clica em gravar e em seguida já clica no botão voltar
            xpath = '/html/body/form/div[5]/div/button[3]' 
            btn_gravar = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
            btn_gravar.send_keys(Keys.ENTER)

            sleep(2) 

            xpath = '//*[@id="BtVoltar"]' 
            btn_voltar = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
            btn_voltar.send_keys(Keys.ENTER)
            sleep(2)

            # abre aba de produtores
            xpath = '//*[@id="BtProdutores"]/div[2]/div[2]/a/span'
            abre_produtores = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
            StatusRPA = '3 - Produtores'
            driver.execute_script('AbreProdutores();')
            driver.switch_to.default_content()
            driver.switch_to.frame('DocumentoRepasse')
            
            ultima_pag = driver.find_element_by_xpath('//*[@id="sp_1_BAR_GridCadastro2"]').get_attribute("innerHTML")
            ultima_pag = int(ultima_pag)
            pag_atual = 1
            limite = 4
            sleep(2)
            while (pag_atual <= ultima_pag):
                #print(f'Pagina atual -> {pag_atual}')
                for x in range(2, 6):
                    try:
                        td_consultor = driver.find_element_by_xpath(f'/html/body/form/div[3]/div/span[2]/span/div/div[3]/div[3]/div/table/tbody/tr[{x}]/td[1]')
                        td_consultor = td_consultor.get_attribute("innerHTML")
                        #print(td_consultor)
            
                        if(td_consultor == 'CONSULTOR'):
                            #print('Achou o consultor, apagando...')
                            xpath = f'/html/body/form/div[3]/div/span[2]/span/div/div[3]/div[3]/div/table/tbody/tr[{x}]/td[4]/a'
                            btn_excluir = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
                            driver.switch_to.default_content()
                            try:
                                xpath = 'body > div.swal2-container.swal2-center.swal2-fade.swal2-shown'
                                alerta_registro = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, xpath)))
                            except:
                                pass
            
                            if(alerta_registro != ''):
                                #print('apareceu')
                                xpath = '/html/body/div[5]/div/div[3]/button[1]'
                                btn_fechar = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
                            
                            
                            driver.switch_to_frame('DocumentoRepasse')
                            break
                    except:
                        #print('Não tem consultor')
                        pass
            
                if(pag_atual != ultima_pag):
                    xpath = '//*[@id="next_BAR_GridCadastro2"]'
                    btn_prox =  wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
                    pag_atual += 1
                else:
                    break
        
            # # Incluir Consultor
            StatusRPA = '3.1 - Consultor'  
            driver.switch_to.default_content()
            driver.switch_to.frame('DocumentoRepasse')
            driver.execute_script("eventoAjax('Incluir');")
            driver.execute_script('acionaJanelaNivel1();')
            driver.switch_to.frame('SearchNivelHierarq')
            xpath = '//*[@id="Descricao"]'
            campo_buscar = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
            campo_buscar.send_keys('CONSULTOR')
            sleep(2)
            driver.execute_script("eventoAjax('EXECUTA');")
            driver.execute_script('RowDblClick("11");')
            sleep(2)
            driver.switch_to.default_content()
            driver.switch_to.frame('DocumentoRepasse')
            sleep(2)
            # Segunda parte
            btnPesquisaLupa = driver.find_element_by_xpath('//*[@id="Divisao1_Bt"]')
            driver.execute_script("arguments[0].click();", btnPesquisaLupa)
            sleep(2)
            driver.switch_to.frame('SearchDivisao')
            xpath = '//*[@id="Nome"]'
            campo_buscar_nome = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
            campo_buscar_nome.send_keys(nome_consultor)
            xpath = '//*[@id="SpanToolBarRightB"]/button'
            wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
            sleep(2)
            # Click no nome do consultor - Lucas S
            xpath = '/html/body/form/div[5]/div[2]/div[2]/div/div[3]/div[3]/div/table/tbody/tr[2]/td[1]/a'
            btn_inclui =  wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
            sleep(2)
            driver.switch_to.default_content()
            driver.switch_to.frame('DocumentoRepasse')
            xpath = '//*[@id="BtGravar"]'
            btn_gravar = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
            sleep(2)
            StatusRPA = '4 - Finalizado'
            status = "Cadastrado"
            
            xpath = '//*[@id="BtAnexar"]/div[2]/div[2]/a/span'
            aba_arquivo = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
            # seguir o diagrama 3.0 linha de cima daqui em diante:
    
else:
    print('RPA NÂO EXISTE')
    driver.quit()

sleep(30)
# start_quiver()