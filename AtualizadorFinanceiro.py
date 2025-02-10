import pyautogui
from flask import Flask, request, jsonify
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time

app = Flask(__name__)

#Acessa a página fornecida pela URL e espera que o elemento com ID 'txt1' esteja presente.
def acessar_pagina(driver, url):
    try:
        driver.get(url)
        forcarCancelamentodaPagina()
        EsperarElementoByID(driver, 'txt1', 30)
    except TimeoutException:
        return "Erro: O campo 'txt1' não foi encontrado após 30 segundos."
    except Exception as e:
        return f"Erro ao acessar a página: {e}"

#Espera até que o texto esteja presente no valor do campo identificado pelo ID fornecido
def AguardaCampoSerApresentadoByID(driver, nomeElemento, valor):
    WebDriverWait(driver, 10).until(
        EC.text_to_be_present_in_element_value((By.ID, str(nomeElemento)), str(valor))
    )

#Espera até que o valor esteja presente no campo identificado pelo ID fornecido
def EsperarElementoByID(driver, nomeElemento, tempo):
    WebDriverWait(driver, tempo).until(
        EC.presence_of_element_located((By.ID, nomeElemento))
    )

#Remove as barras da data fornecida e usa a função SelecionarDataFunc para selecionar o dia, mês e ano
def removerBarra_e_Selecionar(driver, SelecionarDataFunc, data):
    partes = data.split('/')
    dia = partes[0]
    mes = partes[1]
    ano = partes[2]
    SelecionarDataFunc(driver, dia, mes, ano)

#Localiza um elemento na página pelo seu nome
def localizaElementoByNAME(driver, elemento):
    return driver.find_element(By.NAME, elemento)

#Localiza um elemento na página pelo seu nome de classe
def localizaElementoByCLASSNAME(driver, elemento):
    return driver.find_element(By.CLASS_NAME, elemento)

#Localiza um elemento na página pelo seu ID.
def localizaElementoByID(driver, elemento):
    return driver.find_element(By.ID, elemento)

#Localiza e clica em um elemento na página pelo seu nome
def clicarElementoByNAME(driver, nomeElemento):
    elemento = localizaElementoByNAME(driver, nomeElemento)
    elemento.click()

#Seleciona a opção dropdown pelo valor usando o ID do campo
def selecionaCompoPorValorByID(driver, elemento, valor):
    Select(localizaElementoByID(driver, elemento)).select_by_value(valor)

#Seleciona a data (dia, mês, ano) p/ os elementos com final comboData(Dia,Mes,Ano)2
def selecionarData_1(driver, valor_dia, valor_mes, valor_ano):
    selecionaCompoPorValorByID(driver, 'comboDataDia2', valor_dia)
    selecionaCompoPorValorByID(driver, 'comboDataMes2', valor_mes)
    selecionaCompoPorValorByID(driver, 'comboDataAno2', valor_ano)

#Seleciona a data (dia, mês, ano) p/ os elementos com final comboData(Dia,Mes,Ano)3
def selecionarData_2(driver, valor_dia, valor_mes, valor_ano):
    selecionaCompoPorValorByID(driver, 'comboDataDia3', valor_dia)
    selecionaCompoPorValorByID(driver, 'comboDataMes3', valor_mes)
    selecionaCompoPorValorByID(driver, 'comboDataAno3', valor_ano)

#Formata uma string numérica para float, removendo pontos, vírgulas e sinais de porcentagem
def formatar_numero(valor, casas_decimais):
    numero = float(valor.replace(".", "").replace(",", ".").replace("%", "").strip())
    return round(numero, casas_decimais)

#Adaptacao Tecnica para forçar o cancelamento do carregamento da pagina
def forcarCancelamentodaPagina():
    time.sleep(4)
    for i in range(2):
        pyautogui.press('esc')

@app.route('/calculoexato', methods=['POST'])
def calculoexato():
    #Recebe os dados enviados na requisição da aplicação JAVA
    data = request.json
    valor = data['valor']
    data_inicio = data['dataInicio']
    data_fim = data['dataFim']

    #configura e inicializa o WebDriver do Chrome instalando automaticamente o ChromeDriver necessário
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)

    try:
        #acessa a URL diretamente na opção - Atualização de um valor por um índice financeiro
        acessar_pagina(driver, "https://calculoexato.com.br/parprima.aspx?codMenu=FinanAtualizaIndice")
        driver.maximize_window()  # Maximiza o navegador

        #Aguardar o valor ser apresentado e clica 2 vezes no campo para inserir o próximo valor
        AguardaCampoSerApresentadoByID(driver, 'txt1', '0,00')
        clicaNoCampo = ActionChains(driver)
        campoValor = localizaElementoByID(driver, 'txt1')
        clicaNoCampo.double_click(campoValor).perform()
        campoValor.send_keys(str(valor))

        removerBarra_e_Selecionar(driver, selecionarData_1, data_inicio)
        removerBarra_e_Selecionar(driver, selecionarData_2, data_fim)

        selecionaCompoPorValorByID(driver, 'comboIndice4', 'igpm')
        clicarElementoByNAME(driver, 'btnContinuar')
        forcarCancelamentodaPagina()

        #Grava o elemento dos resultados apresentados em tela
        resultadosValoresIndice = localizaElementoByCLASSNAME(driver, "mldi")
        texto = resultadosValoresIndice.text

        #Pega o texto completo e enumera as linhas
        linhas = texto.strip().split('\n')
        linhas_numeradas = list(enumerate(linhas, 1))

        #Inicia as variaveis sem valores para atribuir os valores dentro do loop for
        valor_atualizado = None
        percentual = None
        fator = None
        meses_atualizados = []

        for numero, linha in linhas_numeradas:
            if "Valor atualizado:" in linha:
                valor_atualizado = formatar_numero(linha.split("R$")[1], 2)
            elif "Em percentual:" in linha:
                percentual = formatar_numero(linha.split(":")[1], 4)
            elif "Em fator de multiplicação:" in linha:
                fator = formatar_numero(linha.split(":")[1], 6)
            elif "Os valores do índice utilizados neste cálculo foram:" in linha:
                for linha_mes in linhas[numero:]:
                    if "=" in linha_mes:
                        meses_atualizados.append(linha_mes.strip())
                    else:
                        break

    except NoSuchElementException:
        return jsonify({"error": "Elemento não encontrado"})
    except Exception as e:
        return jsonify({"error": f"Ocorreu um erro: {e}"})
    finally:
        driver.quit()  # fecha o navegador

    # retorna resposta no formato JSON
    # exibe os dados conforme as informações de retorno que são obrigatórios
    return jsonify({
        "sucesso": True,  # Campo booleano
        "valor_original": float(valor),  # Valor original (2 casas decimais)
        "data_inicio": data_inicio,  # Data de início
        "data_fim": data_fim,  # Data de fim
        "valor_atualizado": round(valor_atualizado, 2),  # Valor atualizado (2 casas decimais)
        "percentual": round(percentual, 4),  # Percentual (4 casas decimais)
        "fator": round(fator, 6),  # Fator (6 casas decimais)
        "meses_atualizados": meses_atualizados,  # Lista de variações mensais
        "mensagem": "Consulta realizada com sucesso!"  # Mensagem de sucesso
    })

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000)