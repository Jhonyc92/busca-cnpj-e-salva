# Importa o módulo http.client para realizar solicitações HTTP e HTTPS.
import http.client

# Importa o módulo json para manipulação de dados no formato JSON.
import json

# Importa o módulo pandas, um poderoso pacote de análise de
# dados, usado aqui para manipular dados e exportá-los para Excel.
import pandas as pd

def obter_dados_empresa_por_cnpj(cnpj):
    
    """
    Esta função realiza uma consulta à API ReceitaWS para obter 
            informações detalhadas sobre uma empresa dado seu CNPJ.
    
    Parâmetros:
    cnpj (str): CNPJ da empresa a ser consultada.
    
    Retorna:
    dict: Um dicionário com dados da empresa ou uma mensagem de 
            erro se algo der errado.
    """
    
    # Cria uma conexão HTTPS com o domínio da API da ReceitaWS.
    conexao = http.client.HTTPSConnection("www.receitaws.com.br")
    
    # Envia uma requisição GET para a API incluindo o CNPJ na
    # URL para buscar informações específicas.
    conexao.request("GET", f"/v1/cnpj/{cnpj}")
    
    # Obtém a resposta do servidor à requisição enviada.
    resposta = conexao.getresponse()

    # Imprime o status HTTP da resposta para fins de depuração.
    print(f"Status da Resposta HTTP: {resposta.status}")
    
    # Verifica se o status da resposta é diferente
    # de 200 (OK), indicando um erro.
    if resposta.status != 200:
        
        # Retorna um dicionário com status de erro e a mensagem correspondente.
        return {"status": "ERROR", "message": f"Resposta HTTP com status {resposta.status}"}

    # Lê o conteúdo da resposta HTTP, que está em bytes.
    dados = resposta.read()
    
    # Fecha a conexão HTTPS.
    conexao.close()

    # Tenta decodificar o JSON recebido para um dicionário Python.
    try:
        
        # Decodifica os dados recebidos do tipo bytes para string usando codificação UTF-8.
        # Esta etapa é necessária porque a resposta da API vem em bytes e precisamos convertê-la para
        # uma string antes de tentar transformá-la em um dicionário com json.loads().
        empresa = json.loads(dados.decode("utf-8"))
    
        # Imprime os dados da empresa decodificada para fins de depuração.
        # Esta impressão é útil para verificar se os dados estão sendo corretamente interpretados e
        # convertidos. Mostra o conteúdo do dicionário que representa a empresa, ajudando a identificar
        # se todos os campos necessários estão presentes e corretos.
        print(f"Empresa decodificada: {empresa}")
        
        # Retorna o dicionário contendo as informações da empresa.
        # Se a decodificação foi bem-sucedida e não entrou no bloco 'except', retorna-se o dicionário
        # que pode então ser utilizado para outras finalidades no código, como salvar em um arquivo Excel.
        return empresa
    
    # Captura erros de decodificação JSON, se houver.
    except json.JSONDecodeError as e:
        
        # Se ocorrer um erro durante a decodificação do JSON, ele será capturado aqui.
        # Este bloco 'except' é específico para erros de decodificação JSON, o que significa que se algo
        # der errado durante json.loads(), este bloco será executado.
    
        # Imprime o erro de decodificação para fins de depuração.
        # A impressão do erro ajuda a diagnosticar o problema, mostrando a mensagem de erro
        # associada à exceção. Isso pode indicar, por exemplo, que a resposta da API não estava no
        # formato JSON esperado, o que pode ser causado por um erro no servidor ou uma mudança na API.
        print(f"Erro na decodificação do JSON: {str(e)}")
    
        # Retorna um dicionário com status de erro e uma mensagem personalizada.
        # A mensagem personalizada indica que houve um erro na decodificação do JSON, o que pode
        # ajudar na resolução de problemas e no tratamento de erros no código que chama essa função.
        return {"status": "ERROR", "message": "Erro na decodificação do JSON."}


def salvar_dados_empresa_excel(dados_empresa, nome_arquivo="dados_empresa.xlsx"):
    
    """
    Esta função salva os dados de uma empresa em um arquivo Excel, após 
                verificar se não contêm erros e processar quaisquer dados 
                aninhados para simplificação.
    
    Parâmetros:
                dados_empresa (dict): Dicionário contendo as informações da empresa.
                nome_arquivo (str): Nome do arquivo onde os dados serão salvos, 
                        com valor padrão 'dados_empresa.xlsx'.
    """

    # Verifica se o dicionário de dados da empresa não está vazio e se não contém um status de erro.
    # A condição dados_empresa.get('status') != 'ERROR' assegura que somente dados válidos e sem erros
    # serão processados e salvos. Se o status for 'ERROR', os dados não
    # são salvos e uma mensagem de erro é exibida.
    if dados_empresa and dados_empresa.get('status') != 'ERROR':
        
        # Processa dados aninhados para um formato mais simples antes de salvar.
        # Muitas vezes, os dados da API podem vir em estruturas complexas como listas de dicionários.
        # A função tratar_dados_aninhados é chamada para transformar esses dados aninhados em strings
        # simplificadas ou outros formatos mais convenientes para visualização em um arquivo Excel.
        dados_empresa = tratar_dados_aninhados(dados_empresa)

        # Converte os dados processados da empresa em um DataFrame do pandas.
        # Pandas é uma biblioteca que fornece estruturas de dados poderosas e flexíveis, como o DataFrame,
        # que facilitam a manipulação de dados. Aqui, um DataFrame é criado a partir de uma lista que contém
        # o dicionário dados_empresa, transformando cada par chave-valor do
        # dicionário em colunas e valores no DataFrame.
        df = pd.DataFrame([dados_empresa])

        # Salva o DataFrame em um arquivo Excel.
        # O método to_excel do DataFrame permite a exportação direta
        # dos dados para um arquivo Excel.
        # O parâmetro index=False significa que o índice do DataFrame não será escrito no arquivo,
        # deixando o arquivo mais limpo e focado apenas nos dados.
        df.to_excel(nome_arquivo, index=False)

        # Imprime uma confirmação de que os dados foram salvos com sucesso.
        # Isso fornece um feedback visual no console sobre a conclusão
        # bem-sucedida da operação de salvamento.
        print(f"Dados da empresa salvos com sucesso no arquivo {nome_arquivo}")
        
    else:
        
        # Imprime uma mensagem de erro se não houver dados válidos para salvar.
        # Isso ocorre se o dicionário de dados da empresa for nulo ou contiver um status de 'ERROR'.
        # A mensagem de erro específica é obtida do dicionário dados_empresa e exibida.
        print(f"Não há dados válidos para salvar. Mensagem de erro: {dados_empresa.get('message')}")


def tratar_dados_aninhados(dados):
    
    """
    Esta função processa um dicionário de dados para simplificar a 
                estrutura de campos que contêm listas ou dicionários aninhados, 
                facilitando a posterior visualização e manipulação desses dados.
    
    Parâmetros:
                dados (dict): Dicionário contendo dados complexos, 
                com listas ou dicionários aninhados.
    
    Retorna:
    dict: Retorna o dicionário com os dados aninhados simplificados.
    """

    # Verifica e processa o campo 'atividade_principal', que geralmente contém uma lista de dicionários.
    # Cada dicionário representa uma atividade principal e possui um campo 'text' com a descrição da atividade.
    # O método 'join' é usado para concatenar todas as descrições com um ponto e vírgula entre elas,
    # transformando a lista de descrições em uma única string.
    if "atividade_principal" in dados:
        
        dados['atividade_principal'] = "; ".join([ativ['text'] for ativ in dados['atividade_principal']])

    # Verifica e processa o campo 'atividades_secundarias', semelhante ao campo 'atividade_principal'.
    # Concatena todas as descrições das atividades secundárias em uma única string.
    if "atividades_secundarias" in dados:
        
        dados['atividades_secundarias'] = "; ".join([ativ['text'] for ativ in dados['atividades_secundarias']])

    # Verifica e processa o campo 'qsa', que geralmente contém uma lista de dicionários representando sócios.
    # Cada dicionário tem campos como 'nome' e 'qual' (qualificação do sócio).
    # A string final para cada sócio inclui seu nome e qualificação, separados por parênteses,
    # e todos os sócios são concatenados em uma única string.
    if "qsa" in dados:
        
        dados['qsa'] = "; ".join([f"{q['nome']} ({q.get('qual', '')})" for q in dados['qsa']])

    # Verifica se existe um campo 'billing', que pode ser um dicionário ou um valor específico.
    # Converte o valor ou dicionário completo para string para uniformidade e simplicidade.
    if "billing" in dados:
        
        dados['billing'] = str(dados['billing'])

    # Verifica se existe um campo 'extra', que pode conter informações adicionais em forma de dicionário ou valor.
    # Similar ao campo 'billing', converte todo o conteúdo para string.
    if "extra" in dados:
        
        dados['extra'] = str(dados['extra'])

    # Retorna o dicionário modificado com campos simplificados, facilitando o
    # uso futuro desses dados,
    # especialmente útil para exportação para formatos como CSV ou Excel.
    return dados


# Define um CNPJ para consulta.
cnpj_exemplo = "06947283000160"

# Chama a função para obter dados da empresa usando o CNPJ exemplo.
dados_empresa = obter_dados_empresa_por_cnpj(cnpj_exemplo)

# Chama a função para salvar os dados obtidos em um arquivo Excel.
salvar_dados_empresa_excel(dados_empresa)
