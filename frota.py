from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import pandas as pd

class Frota():
    def __init__(self):
        edge_options = webdriver.EdgeOptions()

        ser = Service("C:\\Users\\victory\\OneDrive - HDI SEGUROS SA\\Área de Trabalho\\msedgedriver.exe")    

        self.navegador = webdriver.Edge(service=ser)

    def ManipulaDf(self):
        #Criação do dataframe
        df = pd.read_excel(r"C:\Users\victory\OneDrive - HDI SEGUROS SA\Área de Trabalho\Programas\Robo Frota Novo\novaPlanilhaHDI - 01.xlsx")

		#Trocando o nome da coluna para conseguir realizar a leitura df.itertuples()
        df.rename(columns={'* Item': 'item'}, inplace = True) 
        df.rename(columns={'* Placa': 'placa'}, inplace = True) 
        df.rename(columns={'* Chassi': 'chassi'}, inplace = True) 
        df.rename(columns={'Tipo de Negócio': 'tpneg'}, inplace = True) 
        df.rename(columns={'Congênere': 'cong'}, inplace = True) 
        df.rename(columns={'Apólice Anterior': 'apolant'}, inplace = True) 
        df.rename(columns={'CI': 'ci'}, inplace = True) 
        df.rename(columns={'Fim Vigência Apólice': 'fimvigant'}, inplace = True) 
        df.rename(columns={'Bônus': 'bonus'}, inplace = True) 
        df.rename(columns={'Quantidade de Sinistro': 'qtdsinistros'}, inplace = True) 
        df.rename(columns={'Cobertura': 'cobertura'}, inplace = True) 
        df.rename(columns={'CEP': 'cep'}, inplace = True) 
        df.rename(columns={'Franquia': 'franquia'}, inplace = True) 
        df.rename(columns={'% FIPE': 'percentfipe'}, inplace = True) 
        df.rename(columns={" Marca": 'marca'}, inplace = True) 
        df.rename(columns={"Modelo": 'modelo'}, inplace = True) 
        df.rename(columns={"Ano Fab.": 'anofab'}, inplace = True) 
        df.rename(columns={"Ano Mod.": 'anomodelo'}, inplace = True) 
        df.rename(columns={"UF": 'uf'}, inplace = True) 
        df.rename(columns={"Renavam": 'renovam'}, inplace = True) 
        df.rename(columns={"Veículo Zero": 'zerokm'}, inplace = True) 
        df.rename(columns={"Número Nota Fiscal": 'notafiscal'}, inplace = True) 
        df.rename(columns={"Data Nota Fiscal": 'dtnotafiscal'}, inplace = True) 
        df.rename(columns={"Código FIPE": 'codfipe'}, inplace = True) 
        df.rename(columns={"Cod. Modelo HDI": 'codmodelohdi'}, inplace = True) 
        df.rename(columns={"Dispositivos de Proteção": 'disprotecao'}, inplace = True) 
        df.rename(columns={"Danos Materiais": 'dm'}, inplace = True) 
        df.rename(columns={"Danos Corporais": 'dc'}, inplace = True) 
        df.rename(columns={"APP Morte": 'app'}, inplace = True) 
        df.rename(columns={"Carroceria ": 'carroceria'}, inplace = True) 
        df.rename(columns={"Limite de Guincho (Km)": 'guincho'}, inplace = True) 
        df.rename(columns={"Carro Reserva": 'carroreserva'}, inplace = True) 
        df.rename(columns={"Vidros": 'vidros'}, inplace = True) 
        df.rename(columns={"Coberturas Adicionais": 'cobadicional'}, inplace = True) 
        df.rename(columns={"Cobertura Complementar": 'cobcomplementar'}, inplace = True) 
        df.rename(columns={"Acessorio/Carroceria": 'acessorio'}, inplace = True) 
        df.rename(columns={"Acessorio/Carroceria Marca": 'acessoriomarca'}, inplace = True) 
        df.rename(columns={"Acessorio/Carroceria Valor": 'acessoriovalor'}, inplace = True) 

        #Coletando a quantidade de itens 
        df['Ultimo_Item'] = df['item'].iloc[-1]

        return df

    def InfoAdicionais(self):
        #Criação do dataframe
        dfadd = pd.read_excel(r"C:\Users\victory\OneDrive - HDI SEGUROS SA\Área de Trabalho\Programas\Robo Frota Novo\massatestesfrotahml.xlsx")

		#Trocando o nome da coluna para conseguir realizar a leitura df.itertuples()
        dfadd.rename(columns={'Corretora': 'corretora'}, inplace = True)
        dfadd.rename(columns={'Usuario': 'usuario'}, inplace = True)  
        dfadd.rename(columns={'CPF/CNPJ do Segurado': 'segcpfcnpj'}, inplace = True) 
        dfadd.rename(columns={'Nome Segurado': 'nome'}, inplace = True) 
        dfadd.rename(columns={'Email Segurado': 'email'}, inplace = True) 
        dfadd.rename(columns={'Produto': 'produto'}, inplace = True) 
        dfadd.rename(columns={'Alugados e Representantes Legais': 'replegais'}, inplace = True) 
        dfadd.rename(columns={'PF (ascendente, descendente)': 'pf'}, inplace = True) 
        dfadd.rename(columns={'Atividade Empresa': 'atvemp'}, inplace = True) 
        dfadd.rename(columns={'Utilização Veiculo': 'utilveic'}, inplace = True) 
        dfadd.rename(columns={'Carga Transportada': 'carga'}, inplace = True) 
        dfadd.rename(columns={'Prevenção e Ger.Risco': 'prevrisco'}, inplace = True) 
        dfadd.rename(columns={'Cobertura RJ e/ou SP': 'cobsprj'}, inplace = True) 
        dfadd.rename(columns={'Motorista Determinado': 'motdet'}, inplace = True) 
        dfadd.rename(columns={'CPF Condutor': 'cpfcondutor'}, inplace = True) 
        dfadd.rename(columns={"Nome Condutor": 'nomecondutor'}, inplace = True) 
        dfadd.rename(columns={"Data Nascimento": 'dtnascimento'}, inplace = True) 
        dfadd.rename(columns={"Sexo": 'sexo'}, inplace = True) 
        dfadd.rename(columns={"Estado Civil": 'estadocivil'}, inplace = True) 
        dfadd.rename(columns={"Propriedade Veiculo": 'propveiculo'}, inplace = True) 

        return dfadd
    
    def EntraIntranet(self,dict_infoadd):
        self.navegador.get("https://www.hdi.com.br/hdi_intranet/login_form")

        time.sleep(10) #Tempo para login intranet

        try:
            WebDriverWait(self.navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[1]/div[1]/form/input')))
            
            time.sleep(1)

            original_window = self.navegador.current_window_handle

            self.navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[1]/form/input').clear()

            self.navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[1]/form/input').send_keys('extranet')

            self.navegador.find_element(By.XPATH, '//*[@id="hdi_search_results"]/a').click()

            time.sleep(4)

            self.navegador.switch_to.window(self.navegador.window_handles[1])

            time.sleep(5)

            nmcorret = str(dict_infoadd['corretora'].iloc[-1])

            usuario = str(dict_infoadd['usuario'].iloc[-1])

            self.navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr/td/form/table[1]/tbody/tr[9]/td[2]/input').send_keys(nmcorret)

            time.sleep(1)

            self.navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr/td/form/table[2]/tbody/tr[1]/td/input[2]').click()

            WebDriverWait(self.navegador, 120).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="id_header_principal"]/tbody/tr/td/table[2]/tbody/tr[2]/td/p')))

            tdnum = 1

            while tdnum < 100:
                tdnum = str(tdnum)
                td = '/html/body/table/tbody/tr/td/table/tbody/tr/td/form/table[1]/tbody/tr[2]/td[3]/table/tbody/tr['+ tdnum + ']/td[1]'
                td1 = '/html/body/table/tbody/tr/td/table/tbody/tr/td/form/table[1]/tbody/tr[2]/td[3]/table/tbody/tr['+ tdnum + ']'
                tdnum = int(tdnum)
                tdnum = tdnum + 1
                xpat = self.navegador.find_element(By.XPATH,td).text
                if xpat == usuario:
                    self.navegador.find_element(By.XPATH, td1).click()
                else:
                    pass

            time.sleep(1)

        except:
            pass

        time.sleep(2)
        pag = ''
        while pag != 'HDI Digital':
            try:
                self.navegador.switch_to.window(self.navegador.window_handles[-1])
                pag = self.navegador.title
            except:
                pass

    def CotacaoFrota(self,dataframe,dict_infoadd):
        try:
            WebDriverWait(self.navegador, 3).until(EC.alert_is_present())
            self.navegador.switch_to.alert.accept()
            time.sleep(2)
        except:
            pass

        try:
            WebDriverWait(self.navegador, 1).until(EC.presence_of_element_located(('xpath','/html/body/div[2]/div/div[20]/span')))
            self.navegador.find_element('xpath','/html/body/div[2]/div/div[20]/span').click()
        except:
            pass

        try:
            WebDriverWait(self.navegador, 1).until(EC.presence_of_element_located(('xpath','/html/body/div[2]/div/div[14]/span')))
            self.navegador.find_element('xpath','/html/body/div[2]/div/div[14]/span').click()
        except:
            pass
        
        WebDriverWait(self.navegador, 15).until(EC.presence_of_element_located(('xpath','/html/body/div[2]/div/div[14]/span')))

        #CPF/CNPJ
        cpfcnpj = int(dict_infoadd['segcpfcnpj'])
        cpfcnpj = str(cpfcnpj)
        self.navegador.find_element('xpath','//*[@id="cpf_cnpj_nova_cotacao"]').send_keys(cpfcnpj)
        self.navegador.find_element('xpath','//*[@id="btnBuscaCotacao"]').click()

        WebDriverWait(self.navegador, 10).until(EC.presence_of_element_located(('xpath','//*[@id="nomeCliente"]')))
        time.sleep(2)

        x = 0    
        for row in dataframe.itertuples():
            # try:
            #Numero do item
            item = int(row.item)
            item = str(item)

            if item == '1':
                #Nome
                nome  = dict_infoadd['nome'].iloc[-1]
                try:		
                    self.navegador.find_element('xpath','//*[@id="nomeCliente"]').clear()
                    self.navegador.find_element('xpath','//*[@id="nomeCliente"]').send_keys(nome)
                except:
                    print('Campo nome já está preenchido.')
                    pass

                time.sleep(1)

                #E-mail
                email = dict_infoadd['email'].iloc[-1]
                if pd.isna(email) is False:
                    try:	
                        self.navegador.find_element('xpath','//*[@id="emailCliente"]').clear()
                        self.navegador.find_element('xpath','//*[@id="emailCliente"]').send_keys(email)
                    except:
                        pass
                else:
                    pass

                time.sleep(1)

                #Produto
                produto = int(dict_infoadd['produto'].iloc[-1])
                produto = str(produto)
                select = Select(self.navegador.find_element(By.XPATH, "/html/body/div[2]/div[2]/div[1]/div/div[1]/div[3]/form/div[2]/div/div[3]/div/fieldset/div[1]/div/select"))
                if item == '1':
                    try:
                        if produto == '131':
                            select.select_by_value('131')
                        elif produto == '139':
                            select.select_by_value('139')
                        else:
                            print('Produto incorreto')
                    except:
                        pass
                else:
                    pass
                
                try:
                    #Clicar no botão de confirmar o produto selecionado
                    self.navegador.find_element('xpath','/html/body/div[2]/div[2]/div[7]/div/div/div/div/div/div[2]/div[3]/button[1]').click()
                except:
                    pass

                time.sleep(4)

                #Numero Itens Total 
                qtitenstotal = int(row.Ultimo_Item)
                qtitenstotal = str(qtitenstotal)
                if item == '1':
                    if produto == '131':
                        self.navegador.find_element('xpath','/html/body/div[2]/div[2]/div[2]/div/div/div/div[2]/span/div/div[2]/span/input').clear()
                        self.navegador.find_element('xpath','/html/body/div[2]/div[2]/div[2]/div/div/div/div[2]/span/div/div[2]/span/input').send_keys(qtitenstotal)
                        time.sleep(1)
                        self.navegador.find_element('xpath','/html/body/div[2]/div[2]/div[2]/div/div/div/div[2]/span/div/div[3]/button').click() 
                    elif produto == '139':
                        self.navegador.find_element('xpath','/html/body/div[2]/div[2]/div[1]/div/div[1]/div[2]/input').clear()
                        self.navegador.find_element('xpath','/html/body/div[2]/div[2]/div[1]/div/div[1]/div[2]/input').send_keys(qtitenstotal)
                    else:
                        pass
                else:
                    pass

                #Avaliação de Risco
                time.sleep(2)

                #Alugados e Representantes Legais
                arl = int(dict_infoadd['replegais'].iloc[-1])
                arl = str(arl)
                if produto == '131':		
                    self.navegador.find_element('xpath','//*[@id="text_pergunta_1"]').send_keys(arl)
                else:
                    pass
                
                time.sleep(1)

                #PF (ascendentes, descendentes e cônjuges)
                pf = int(dict_infoadd['pf'].iloc[-1])
                pf = str(pf)
                if produto == '131':		
                    self.navegador.find_element('xpath','//*[@id="text_pergunta_2"]').send_keys(pf)
                else:
                    pass
                
                time.sleep(1)

                #Atividade Empresa
                atvemp = int(dict_infoadd['atvemp'].iloc[-1])
                atvemp = str(atvemp)
                select = Select(self.navegador.find_element(By.XPATH,'//*[@id="comboRespostasFrota0"]'))
                try:
                    select.select_by_value(atvemp)
                except:
                    print('Seleção de atividade da empresa já selecionada ou não encontrada.')
                
                time.sleep(1)

                #Utilização Veiculo
                utilveic = int(dict_infoadd['utilveic'].iloc[-1])
                utilveic = str(utilveic)
                if produto == '131':
                    select = Select(self.navegador.find_element(By.XPATH,'//*[@id="comboRespostasFrota1"]'))
                    select.select_by_value(utilveic)
                else:
                    pass
                
                time.sleep(1)

                #Carga
                carga = int(dict_infoadd['carga'].iloc[-1])
                carga = str(carga)
                if produto == '131':
                    select = Select(self.navegador.find_element(By.XPATH,'//*[@id="comboRespostasFrota2"]'))
                    try:
                        select.select_by_value(carga)
                    except:
                        print('Seleção de carga já selecionada ou não encontrada.')
                else:
                    pass

                time.sleep(1)

                #Frota com prevenção e gerenciamento de risco
                prevrisco = int(dict_infoadd['prevrisco'].iloc[-1])
                prevrisco = str(prevrisco)
                if  prevrisco == '1': 
                    self.navegador.find_element('xpath','//*[@id="avaliacaoFrotaOpcao0_0"]').click()
                else:   
                    self.navegador.find_element('xpath','//*[@id="avaliacaoFrotaOpcao0_1"]').click()
                
                time.sleep(1)

                #Coertura SP e/ou RJ
                cobsprj = int(dict_infoadd['cobsprj'].iloc[-1])
                cobsprj = str(cobsprj)
                if cobsprj == '1':
                    self.navegador.find_element('xpath','//*[@id="avaliacaoFrotaOpcao1_0"]').click()
                else:
                    self.navegador.find_element('xpath','//*[@id="avaliacaoFrotaOpcao1_1"]').click()
                
                time.sleep(1)
            else:
                pass

            #Placa
            placa = row.placa
            chassi = row.chassi
            anofab = row.anofab
            modelo = row.modelo
            marca = row.marca
            modelo = modelo.split(' ')
            modelo = modelo[0] + ' ' + modelo[1]
            if pd.isna(anofab) is False:
                anofab = int(row.anofab)
                anofab = str(anofab)
            try:
                self.navegador.find_element('xpath','//*[@id="placa"]').clear()
                time.sleep(1)
                self.navegador.find_element('xpath','//*[@id="placa"]').send_keys(placa)
                time.sleep(1)
                self.navegador.find_element('xpath','//*[@id="btnProcurarVeiculoPlaca"]').click()
                time.sleep(2)
                try:
                    WebDriverWait(self.navegador, 8).until(EC.presence_of_element_located(('xpath','//*[@id="modalConteudo"]/table')))
                    time.sleep(2)
                    self.navegador.find_element('xpath','//*[@id="modalConteudo"]/table').click()
                except:
                    pass
                time.sleep(3)
                msg = self.navegador.find_element(By.ID,'modalMensagensConteudo').text
                print(msg)
                if msg == 'O ano retornado não se encontra cadastrado no sistema.':
                    self.navegador.find_element('xpath','//*[@id="modalMensagens"]/div/div/div[1]/span').click()
                    time.sleep(1) 
                    self.navegador.find_element('xpath','/html/body/div[2]/div[2]/div[1]/div/div[1]/div[3]/form/div[7]/div/div[3]/div/fieldset/div[6]/div/select[1]/option[2]').click()
                elif msg == 'Não foi possível encontrar um modelo equivalente a esta placa. Por favor informe o Chassi para uma nova busca':
                    self.navegador.find_element('xpath','//*[@id="modalMensagens"]/div/div/div[1]/span').click()
                    time.sleep(1)
                    self.navegador.find_element('xpath','//*[@id="marcaVeiculoBusca"]').clear()
                    time.sleep(1)
                    self.navegador.find_element('xpath','//*[@id="marcaVeiculoBusca"]').send_keys(f'{marca} {modelo}')
                    WebDriverWait(self.navegador, 10).until(EC.presence_of_element_located(('xpath','//*[@id="btnProcurarVeiculoMarcaModelo"]')))
                    time.sleep(1)
                    self.navegador.find_element('xpath','//*[@id="btnProcurarVeiculoMarcaModelo"]').click()
                    time.sleep(2)
                    try:
                        WebDriverWait(self.navegador, 8).until(EC.presence_of_element_located(('xpath','//*[@id="modalConteudo"]/table')))
                        time.sleep(2)
                        self.navegador.find_element('xpath','//*[@id="modalConteudo"]/table').click()
                    except:
                        pass
                    time.sleep(3)
                    select = Select(self.navegador.find_element(By.XPATH,'//*[@id="anoFabricacao"]'))
                    select.select_by_value(anofab) #Seleciona o ano de fabricação
                else:
                    print('Chassi correto')
            except:
                print(f'Veiculo não encontrado pela placa e chassi no item {item}, por favor rever.')	

            time.sleep(5)

            #TpNegocio
            tpneg = row.tpneg
            cong = row.cong
            bonus = row.bonus
            qtdsinistros = row.qtdsinistros
            apolant = row.apolant
            ci = row.ci
            fimvigant = str(row.fimvigant).replace('00:00:00','').split('-')
            fimvigant = f'{fimvigant[2]}/{fimvigant[1]}/{fimvigant[0]}'
            if pd.isna(bonus) is False:
                bonus = int(row.bonus)
                bonus = str(bonus)
                if len(bonus) == 1:
                    bonus = '0' + bonus
            if pd.isna(qtdsinistros) is False:
                qtdsinistros = int(row.qtdsinistros)
                qtdsinistros = str(qtdsinistros)
            if pd.isna(apolant) is False:
                apolant = int(row.apolant)
                apolant = str(apolant)
            if pd.isna(ci) is False:
                ci = int(row.ci)
                ci = str(ci)
                
            if tpneg == 'Novo Negócio':
                self.navegador.find_element('xpath','//*[@id="tipoNegociacaoNovoNegocio"]').click() #Seleciona SN
            elif tpneg == 'Renovação HDI':
                self.navegador.find_element('xpath','//*[@id="tipoRenovacaoHDI"]').click() #Seleciona RHDI
                time.sleep(2)
                self.navegador.find_element('xpath','//*[@id="btnBuscaCongener"]').click() #Aperta o botão para buscar as congeneres
                WebDriverWait(self.navegador, 10).until(EC.presence_of_element_located(('xpath','//*[@id="modalConteudo"]/table/tbody/tr[3]')))
                time.sleep(1)
                self.navegador.find_element('xpath','//*[@id="modalConteudo"]/table/tbody/tr[3]').click() #Atualmente está selecionada a opção 6572, caso necessario mudar
                time.sleep(3)
                select = Select(self.navegador.find_element(By.XPATH,'//*[@id="selectListaClasseBonus"]'))
                select.select_by_value(bonus) #Seleciona a classe bônus
                self.navegador.find_element('xpath','//*[@id="quantidadeSinistros"]').send_keys(qtdsinistros) #Quantidade de sinistros
                self.navegador.find_element('xpath','//*[@id="apoliceNumero"]').send_keys(apolant) #Verificar quais campos são preenchidos na apolice anterior
                self.navegador.find_element('xpath','//*[@id="dataFinalVigenciaApolice"]').send_keys(Keys.CONTROL, 'a')
                self.navegador.find_element('xpath','//*[@id="dataFinalVigenciaApolice"]').send_keys(fimvigant) #Data do fim de vigência anterior 
                self.navegador.find_element('xpath','//*[@id="seqCiante"]').send_keys(ci) #CI anterior
            elif tpneg == 'Renovação Congênere':
                self.navegador.find_element('xpath','//*[@id="tipoNegociacaoRenovacaoCongenere"]').click() #Seleciona RC
                time.sleep(2)
                self.navegador.find_element('xpath','//*[@id="nomeCia"]').send_keys(cong) #Seleciona RC
                time.sleep(1)
                self.navegador.find_element('xpath','//*[@id="btnBuscaCongener"]').click() #Aperta o botão para buscar as congeneres
                WebDriverWait(self.navegador, 10).until(EC.presence_of_element_located(('xpath','//*[@id="modalConteudo"]/table/tbody/tr')))
                time.sleep(1)
                self.navegador.find_element('xpath','//*[@id="modalConteudo"]/table/tbody/tr').click() #Seleciona a congênere buscada
                time.sleep(3)
                select = Select(self.navegador.find_element(By.XPATH,'//*[@id="selectListaClasseBonus"]'))
                select.select_by_value(bonus) #Seleciona a classe bônus
                self.navegador.find_element('xpath','//*[@id="quantidadeSinistros"]').send_keys(qtdsinistros) #Quantidade de sinistros
                self.navegador.find_element('xpath','//*[@id="dataFinalVigenciaApolice"]').send_keys(Keys.CONTROL, 'a')
                self.navegador.find_element('xpath','//*[@id="dataFinalVigenciaApolice"]').send_keys(fimvigant) #Data do fim de vigência anterior 
            else:
                print('Tipo de negócio não identificado, por favor corrigir na planilha.')

            time.sleep(2)

            #CEP
            cep = str(row.cep)
            cep = cep.replace('-','').replace('.0','')
            if pd.isna(cep) is False:
                while len(cep) != 8:
                    cep = '0' + cep

                cep = cep[0:5] + '-' + cep[5:]

                self.navegador.find_element('xpath','//*[@id="cepPernoite"]').clear()
                self.navegador.find_element('xpath','//*[@id="cepPernoite"]').send_keys(Keys.CONTROL, 'a')
                self.navegador.find_element('xpath','//*[@id="cepPernoite"]').send_keys(cep)
            else:
                pass

            #Avaliação de Risco

            time.sleep(1)

            #Motorista Determinado - Produto 139
            motdet = str(dict_infoadd['motdet'].iloc[-1])
            if produto == '139':
                select_cond = Select(self.navegador.find_element(By.XPATH,'///*[@id="comboRespostasCondutorDeterminado_204"]'))
                if motdet == 'Sim' or motdet == 'sim':
                    cpfcondutor = int(dict_infoadd['cpfcondutor'].iloc[-1])
                    cpfcondutor = str(cpfcondutor)
                    nomecondutor = str(dict_infoadd['nomecondutor'].iloc[-1])
                    dtnascimento = str(dict_infoadd['dtnascimento'].iloc[-1])
                    sexo = int(dict_infoadd['sexo'].iloc[-1])
                    sexo = str(sexo)
                    estadocivil = str(dict_infoadd['estadocivil'].iloc[-1])
                    estadocivil = '003,0' + estadocivil
                    propveiculo = int(dict_infoadd['propveiculo'].iloc[-1])
                    propveiculo = str(propveiculo)
                    utilveic = int(dict_infoadd['utilveic'].iloc[-1])
                    utilveic = str(utilveic)
                    select_cond.select_by_value('1121')
                    time.sleep(1)
                    self.navegador.find_element('xpath','//*[@id="cpfCondutor0"]').send_keys(cpfcondutor) #CPF Condutor
                    time.sleep(1)
                    self.navegador.find_element('xpath','//*[@id="nomeCondutor0"]').send_keys(nomecondutor) #Nome Condutor
                    time.sleep(1)
                    self.navegador.find_element('xpath','//*[@id="dataNascimentoCondutor0"]').send_keys(dtnascimento) #Data Nascimento Condutor
                    if sexo == '1':
                        self.navegador.find_element('xpath','//*[@id="sexoCondutorMasculino0"]').click() #Escolha do sexo masculino
                    elif sexo == '2':
                        self.navegador.find_element('xpath','//*[@id="sexoCondutorFeminino0"]').click() #Escolha do sexo feminino
                    time.sleep(1)
                    select_estdcivil = Select(self.navegador.find_element(By.XPATH,'//*[@id="comboRespostas0_0"]'))
                    select_estdcivil.select_by_value(estadocivil) #Escolha do Estado Civil
                else:
                    select_cond.select_by_value('1122')
                select_propveic = Select(self.navegador.find_element(By.XPATH,'//*[@id="comboRespostasGeral_205"]'))
                select_propveic.select_by_value(propveiculo) #Seleção da propriedade do veiculo
                select_utilzveic = Select(self.navegador.find_element(By.XPATH,'//*[@id="comboRespostasGeral_206"]'))
                select_utilzveic.select_by_value(utilveic) #Seleção da utilização do veiculo
            else:
                pass

            time.sleep(1)

            #Cobertura e Franquia
            cobertura = row.cobertura	
            franquia = row.franquia
            if pd.isna(cobertura) is False and pd.isna(franquia) is False:
                if cobertura == 'RCF':
                    self.navegador.find_element('xpath','//*[@id="tipoCoberturaExclusiva"]').click()
                    time.sleep(2)
                    select_franquia = Select(self.navegador.find_element(By.XPATH,'//*[@id="franquia"]'))
                    try:
                        select_franquia.select_by_visible_text(franquia) #Seleção da franquia
                    except:
                        print('Franquia não selecionada.')
                elif cobertura == 'Compreensiva':
                    self.navegador.find_element('xpath','//*[@id="tipoCoberturaCompreensiva"]').click()
                else:
                    print('Cobertura  Inválida!')
            else:
                pass

            time.sleep(1)

            #FIPE %
            fipe = row.percentfipe
            if pd.isna(fipe) is False:
                fipe = int(row.percentfipe)
                fipe = str(fipe)
                try:
                    self.navegador.find_element('xpath','//*[@id="valorMercadoPercentual"]').clear()
                    self.navegador.find_element('xpath','//*[@id="valorMercadoPercentual"]').send_keys(fipe)
                except:
                    pass
            else:
                pass
            
            time.sleep(1)

            #Valor Determinado
            carroceria = row.carroceria
            if pd.isna(carroceria) is False:
                carroceria = float(row.carroceria)
                carroceria = str(carroceria)
                self.navegador.find_element('xpath','//*[@id="valorDeterminado"]').clear()
                self.navegador.find_element('xpath','//*[@id="valorDeterminado"]').send_keys(carroceria)
            else:
                pass

            time.sleep(1)

            #RCFV Danos Materiais
            dm = row.dm
            if pd.isna(dm) is False:
                dm = int(row.dm)
                dm = str(dm)
                select = Select(self.navegador.find_element(By.XPATH,'//*[@id="danosMateriais"]'))
                select.select_by_value(dm)
            else:
                pass

            time.sleep(1)

            #RCFV Danos Corporais
            dc = row.dc
            if pd.isna(dc) is False:
                dc = int(row.dc)
                dc = str(dc)
                select = Select(self.navegador.find_element(By.XPATH,'//*[@id="danosCorporais"]'))
                select.select_by_value(dc)
            else:
                pass

            time.sleep(1)

            #APP Morte/Invalidez
            app = row.app
            if pd.isna(app) is False:
                app = int(row.app)
                app = str(app) + '00'
                self.navegador.find_element('xpath','//*[@id="morteInvalidez"]').clear()
                self.navegador.find_element('xpath','//*[@id="morteInvalidez"]').send_keys(app)
            else:
                pass

            time.sleep(2)
            
            #Necessidades do cliente
            #Guincho
            guincho = row.guincho
            if pd.isna(guincho) is False:
                if guincho == '' or guincho == 'Sem Guincho': 
                    self.navegador.find_element('xpath','//*[@id="8"]').click() #Sem Beneficios
                elif guincho == '100 KM' or guincho == '300 KM':
                    try:
                        self.navegador.find_element('xpath','//*[@id="21"]').click() #Básico
                    except:
                        self.navegador.find_element('xpath','//*[@id="23"]').click() #Básico
                    if guincho == '100 KM':
                        self.navegador.find_element('xpath','//*[@id="coberturaBen0381"]').click() #Guincho 100 KM
                    else:
                        self.navegador.find_element('xpath','//*[@id="coberturaBen0382"]').click() #Guincho 300 KM
                elif guincho == '500 KM':
                    try:
                        self.navegador.find_element('xpath','//*[@id="22"]').click() #Completo
                    except:
                        self.navegador.find_element('xpath','//*[@id="24"]').click() #Completo
                elif guincho == '800 KM' or guincho == 'Sem Limite Km' or guincho == 'Sem Limite KM':
                    try:
                        self.navegador.find_element('xpath','//*[@id="3"]').click() #VIP
                    except:
                        self.navegador.find_element('xpath','//*[@id="6"]').click() #VIP
                else:
                    print('Necessidades do cliente desconhecida, por favor confira as disponibilidades da planilha!')
            else:
                pass

            time.sleep(1)

            #Carro Reserva
            carroreserva = row.carroreserva
            if pd.isna(carroreserva) is False:
                if carroreserva == '' or carroreserva == 'Sem Carro Reserva':
                    self.navegador.find_element('xpath','//*[@id="coberturaBen0106"]').click()
                elif carroreserva == '07 Dias CR Manual':
                    self.navegador.find_element('xpath','//*[@id="coberturaBen0107"]').click()
                elif carroreserva == '15 Dias CR Manual':
                    self.navegador.find_element('xpath','//*[@id="coberturaBen0108"]').click()
                elif carroreserva == '30 Dias CR Manual':
                    self.navegador.find_element('xpath','//*[@id="coberturaBen0109"]').click()
                elif carroreserva == '45 Dias CR Manual':
                    self.navegador.find_element('xpath','//*[@id="coberturaBen0110"]').click()
                elif carroreserva == '07 Dias CR Automático':
                    self.navegador.find_element('xpath','//*[@id="coberturaBen0241"]').click()
                elif carroreserva == '15 Dias CR Automático':
                    self.navegador.find_element('xpath','//*[@id="coberturaBen0242"]').click()
                elif carroreserva == '30 Dias CR Automático':
                    self.navegador.find_element('xpath','//*[@id="coberturaBen0243"]').click()
                elif carroreserva == '45 Dias CR Automático':
                    self.navegador.find_element('xpath','//*[@id="coberturaBen0244"]').click()
                else:
                    print('Valor de carro reserva desconhecido, por favor confira as disponibilidades da planilha!')
            else:
                pass

            time.sleep(1)

            #Vidros
            vid = row.vidros
            if pd.isna(vid) is False:
                vid = vid.split(';')
                for vidros in vid:
                    if vidros == '':
                        pass
                    elif vidros == 'HDI Auto Vidros' or vidros == 'Vidros':
                        self.navegador.find_element('xpath','//*[@id="coberturaBen0017"]').click()
                    elif vidros == 'Faróis, Lanternas, Retrov, Auxiliar':
                        self.navegador.find_element('xpath','//*[@id="coberturaBen0019"]').click()
                    elif vidros == 'Hdi Auto Vidros Blindados':
                        self.navegador.find_element('xpath','//*[@id="coberturaBen0176"]').click()
                    elif vidros == 'Lant, Retrov, Farois Blindados':
                        self.navegador.find_element('xpath','//*[@id="coberturaBen0177"]').click()
                    elif vidros == 'Logomarca':
                        self.navegador.find_element('xpath','//*[@id="coberturaBen0169"]').click()
                    else:
                        print('Cobertura de vidros desconhecida, por favor confira as disponibilidades da planilha!')
            else:
                pass

            time.sleep(1)

            #Cobertura Complementar
            cobcomp = row.cobcomplementar
            if pd.isna(cobcomp) is False:
                cobcomp = cobcomp.split(';')
                for cobcomplementar in cobcomp:
                    if cobcomplementar == '':
                        pass
                    elif cobcomplementar == 'Higienização':
                        self.navegador.find_element('xpath','//*[@id="coberturaBen0169"]').click()
                    elif cobcomplementar == 'Teto Solar':
                        self.navegador.find_element('xpath','//*[@id="coberturaBen0169"]').click()
                    else:
                        print('Cobertura complementar desconhecida, por favor confira as disponibilidades da planilha!')
            else:
                pass

            time.sleep(1)

            #Cobertura Adicional
            cobadd = row.cobadicional
            if pd.isna(cobadd) is False:
                cobadd = cobadd.split(';')
                for cobadicional in cobadd:
                    if cobadicional == '':
                        pass
                    elif cobadicional == 'Despesas Extras':
                        self.navegador.find_element('xpath','//*[@id="coberturaBen0015"]').click()
                    elif cobadicional == 'Extensao Perim Urug,Arg E Paraguai':
                        self.navegador.find_element('xpath','//*[@id="coberturaBen0043"]').click()
                    elif cobadicional == 'Extensao Soc Dir':
                        self.navegador.find_element('xpath','//*[@id="coberturaBen0042"]').click()
                    elif cobadicional == 'Danos Morais':
                        self.navegador.find_element('xpath','//*[@id="coberturaBen0041"]').click()
                    elif cobadicional == 'Desp. Medicas Hospitalares':
                        self.navegador.find_element('xpath','//*[@id="coberturaBen0082"]').click()
                    elif cobadicional == 'Basculamento':
                        self.navegador.find_element('xpath','//*[@id="coberturaBen0175"]').click()
                    else:
                        print('Cobertura complementar desconhecida, por favor confira as disponibilidades da planilha!')
            else:
                pass
            
            #Acessorios
            acess = row.acessorio
            acessmarca = row.acessoriomarca
            acessvalor = row.acessoriovalor
            if pd.isna(acess) is False:
                acess = acess.split(';')
                acessmarca = acessmarca.split(';')
                acessvalor = acessvalor.split(';')
                id = 0
                if acessorio == '':
                    pass
                else:
                    for acessorio in acess:
                        for acessoriomarca in acessmarca:
                            for acessoriovalor in acessvalor:
                                self.navegador.find_element('xpath','//*[@id="checkboxComAcessorios"]').click()
                                WebDriverWait(self.navegador, 5).until(EC.presence_of_element_located(('xpath','//*[@id="selectListaAcessorio"]')))
                                time.sleep(1)
                                selectacess = Select(self.navegador.find_element(By.XPATH,'//*[@id="selectListaAcessorio"]'))
                                try:
                                    select.select_by_visible_text(acessorio)
                                    self.navegador.find_element('xpath','//*[@id="marcaAcessorio"]').send_keys(acessoriomarca)
                                    self.navegador.find_element('xpath','//*[@id="valorAcessorio"]').send_keys(acessoriovalor)
                                    id += 1
                                    if id != len(acess):
                                        self.navegador.find_element('xpath','//*[@id="acessorios"]/tfoot/tr/td[3]/span').click()
                                    else:
                                        pass
                                except:
                                    print('Acessório não existente para a categoria tarifaria ou com erros no Excel. Por favor, verificar.')
            else:
                pass

            x = x + 1
            print(x)

            #Decisão entre adicionar novo item ou calcular
            if x != len(dataframe.axes[0]):
                WebDriverWait(self.navegador, 15).until(EC.element_to_be_clickable(('xpath','//*[@id="btnAdicionarItemCotacaoPainel"]')))
                time.sleep(2)
                #Adicionar novo item
                self.navegador.find_element('xpath','//*[@id="btnAdicionarItemCotacaoPainel"]').click()
                time.sleep(6)
            else:
                #Calcular Cotação
                self.navegador.find_element('xpath','//*[@id="botaoCalcularPC"]').click()
                time.sleep(6)

                #Fehar Janelas adicionais
                try:
                    WebDriverWait(self.navegador, 10).until(EC.presence_of_element_located(('xpath','/html/body/div[2]/div[2]/div[2]/div/div/div/div[1]/span')))
                    self.navegador.find_element('xpath','/html/body/div[2]/div[2]/div[2]/div/div/div/div[1]/span').click()
                except:
                    pass

                #Pegar numero da cotação
                nc = ''
                fr = ''
                pl = ''
                while nc == '' or pl == '':
                    nc = self.navegador.find_element(By.ID,'valorNumeroCotacao').text
                    fr = self.navegador.find_element(By.ID,'franquiaCotacao').get_attribute('value')
                    pl = self.navegador.find_element(By.ID,'premioLiquidoCotacao').get_attribute('value')
                valid = {"Numero Cotação": nc, "Franquia" : fr ,"Prêmio Liquido" : pl}
                valida_tabelas = pd.DataFrame([valid])

                #Enviar informações da cotação para o arquivo csv
                valida_tabelas.to_csv(r'C:\Users\victory\OneDrive - HDI SEGUROS SA\Área de Trabalho\Programas\Robo Frota Novo\NroCotFrota.csv', index=False,mode = 'a',sep=';',header=True)

                print(nc)

            # except:
            # 	c = int(c)
            # 	x = x - (c - 1)
            # 	self.navegador.refresh()
            # 	time.sleep(5)
            # 	try:
            # 		WebDriverWait(self.navegador, 20).until(EC.alert_is_present())
            # 		self.navegador.switch_to.alert.accept()
            # 		time.sleep(5)
            # 	except:
            # 		pass
            # 	try:
            # 		WebDriverWait(self.navegador, 20).until(EC.presence_of_element_located(('xpath','/html/body/div[2]/div/div[14]/span')))
            # 		self.navegador.find_element('xpath','/html/body/div[2]/div/div[14]/span').click()
            # 	except:
            # 		pass
            # 	time.sleep(5)

if  __name__ == "__main__":
    iniciaprograma = Frota()
    df = iniciaprograma.ManipulaDf()
    df_add = iniciaprograma.InfoAdicionais()
    iniciaprograma.EntraIntranet(df_add)
    iniciaprograma.CotacaoFrota(df,df_add)