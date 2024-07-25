from botcity.web import WebBot, Browser, By
from botcity.core import DesktopBot
from prj_T2C_GoogleViagens.classes_t2c.utils.T2CMaestro import T2CMaestro, LogLevel, ErrorType
from prj_T2C_GoogleViagens.classes_t2c.utils.T2CExceptions import BusinessRuleException
from prj_T2C_GoogleViagens.classes_t2c.sqlite.T2CSqliteQueue import T2CSqliteQueue
from prj_T2C_GoogleViagens.classes_t2c.FuncoesDadosMundiais import FuncoesDadosMundiais
from prj_T2C_GoogleViagens.classes_t2c.FuncoesGoogleViagens import FuncoesGoogleViagens
# testar plugin da botcity para extração de tabelas
from botcity.web.parsers import table_to_dict
from openpyxl import Workbook
import os

class T2CInitAllApplications:
    """
    Classe feita para ser invocada principalmente no começo de um processo, para iniciar os processos necessários para a automação.
    """

    #Iniciando a classe, pedindo um dicionário config e o bot que vai ser usado e enviando uma exceção caso nenhum for informado
    def __init__(self, arg_dictConfig:dict, arg_clssMaestro:T2CMaestro, arg_botWebbot:WebBot=None, arg_botDesktopbot:DesktopBot=None, arg_clssSqliteQueue:T2CSqliteQueue=None):
        """
        Inicializa a classe T2CInitAllApplications.

        Parâmetros:
        - arg_dictConfig (dict): dicionário de configuração.
        - arg_clssMaestro (T2CMaestro): instância de T2CMaestro.
        - arg_botWebbot (WebBot): instância de WebBot (opcional, default=None).
        - arg_botDesktopbot (DesktopBot): instância de DesktopBot (opcional, default=None).
        - arg_clssSqliteQueue (T2CSqliteQueue): instância de T2CSqliteQueue (opcional, default=None).

        Retorna:
        """
        
        if(arg_botWebbot is None and arg_botDesktopbot is None): raise Exception("Não foi possível inicializar a classe, forneça pelo menos um bot")
        else:
            self.var_botWebbot = arg_botWebbot
            self.var_botDesktopbot = arg_botDesktopbot
            self.var_dictConfig = arg_dictConfig
            self.var_clssMaestro = arg_clssMaestro
            self.var_clssSqliteQueue = arg_clssSqliteQueue

    def add_to_queue(self):
        """
        Adiciona itens à fila no início do processo, se necessário.

        Observação:
        - Código placeholder.
        - Se o seu projeto precisa de mais do que um método simples para subir a sua fila, considere fazer um projeto dispatcher.

        Parâmetros:
        """
       

    def execute(self, arg_boolFirstRun=False, arg_clssSqliteQueue=None):
        """
        Executa a inicialização dos aplicativos necessários.

        Parâmetros:
        - arg_boolFirstRun (bool): indica se é a primeira execução (default=False).
        - arg_clssSqliteQueue (T2CSqliteQueue): instância da classe T2CSqliteQueue (opcional, default=None).
        
        Observação:
        - Edite o valor da variável `var_intMaxTentativas` no arquivo Config.xlsx.
        
        Retorna:

        Raises:
        - BusinessRuleException: em caso de erro de regra de negócio.
        - Exception: em caso de erro geral.
        """

         #inicialização das funções
        var_clssFuncoesDadosMundiais = FuncoesDadosMundiais(arg_dictConfig=self.var_dictConfig, arg_botDesktopbot=self.var_botDesktopbot,arg_botWebbot=self.var_botWebbot,arg_clssMaestro=self.var_clssMaestro)
        var_clssFuncoesGoogleViagens = FuncoesGoogleViagens(arg_dictConfig=self.var_dictConfig, arg_botDesktopbot=self.var_botDesktopbot,arg_botWebbot=self.var_botWebbot,arg_clssMaestro=self.var_clssMaestro)
        

        #inicialização das variáveis
        var_strURLDadosMundiais = self.var_dictConfig["strURLDadosMundiais"]
        var_strURLGoogleViagens = self.var_dictConfig["strURLGoogleViagens"]


        #Chama o método para subir a fila, apenas se for a primeira vez
        if(arg_boolFirstRun):
            self.add_to_queue()
            #criando o arquivo excel para o salvamento dos resultados extraídos

            var_strCaminhoArquivoPassagens = self.var_dictConfig["strCaminhoArquivoPassagens"]
            

            # Verifica se a pasta do arquivo passagens está vazia
            #lista os arquivos da pasta
            arquivos = os.listdir(var_strCaminhoArquivoPassagens)

            
            if len(arquivos) == 0:
                print(f"A pasta '{var_strCaminhoArquivoPassagens}' está vazia.")
            else:
                print(f"A pasta '{var_strCaminhoArquivoPassagens}' não está vazia. Excluindo arquivos...")

                # Itera sobre os arquivos e os exclui
                for arquivo in arquivos:
                    caminho_arquivo = os.path.join(var_strCaminhoArquivoPassagens, arquivo)
                    try:
                        if os.path.isfile(caminho_arquivo):
                            os.remove(caminho_arquivo)
                            print(f"Arquivo '{arquivo}' excluído com sucesso.")
                        else:
                            print(f"O caminho '{caminho_arquivo}' não é um arquivo.")
                    except Exception as e:
                        print(f"Erro ao excluir o arquivo '{arquivo}': {e}")

                print("Todos os arquivos foram excluídos.")


    #        Criando um objeto Workbook (livro)
                workbook = Workbook()

                # Acessando a primeira aba (padrão)
                aba_todos = workbook.active
                aba_todos.title = "Todos"

                # Adicionando cabeçalhos (primeira linha)
                aba_todos.append(["País", "Cidade", "Preço"])                  

                nova_aba = workbook.create_sheet(title='Baratos')

                # Adicionando cabeçalhos (primeira linha) na nova aba
                nova_aba.append(["País", "Cidade", "Preço"])  

                                # Salvando o arquivo Excel
                var_strNomeArquivo = os.path.join(var_strCaminhoArquivoPassagens,"Passagens.xlsx")

                workbook.save(filename=var_strNomeArquivo)

                print(f"Arquivo '{var_strNomeArquivo}' criado com sucesso.")


           
            # Abrir o site dados mundiais
            self.var_clssMaestro.write_log("Iniciando a extração dos países...")
            var_clssFuncoesDadosMundiais.abrirDadosMundiais(var_strURLDadosMundiais)

            # Extrai a tabela de países 
            var_strTabelaPaises = self.var_botWebbot.find_element('//table[@class="std100 hover tabsort sticky"]', By.XPATH)

            #transforma a tabela de países webelement em uma lista
            var_dictTabelaPaises = table_to_dict(table=var_strTabelaPaises)

            # Limpando fila
            self.var_clssMaestro.write_log("Limpando a fila de execução...")
            self.var_clssSqliteQueue.abandon_queue()
            
            #Laço para o append da lista de 30 países para a lista de execução
            self.var_clssMaestro.write_log("Inserindo os países na fila de execução...")
            var_listPaises = []
            for i in var_dictTabelaPaises[:1]:
                var_listPaises.append(i['paísregião'])  # Adicionar o nome do país à lista

                self.var_clssSqliteQueue.insert_new_queue_item(arg_strReferencia=i['paísregião'])

            self.var_clssMaestro.write_log("Fila de execução criada com sucesso")           


        #Edite o valor dessa variável a no arquivo Config.xlsx
        var_intMaxTentativas = self.var_dictConfig["MaxRetryNumber"]
        
        for var_intTentativa in range(var_intMaxTentativas):
            try:
                self.var_clssMaestro.write_log("Iniciando aplicativos, tentativa " + (var_intTentativa+1).__str__())
                #Insira aqui seu código para iniciar os aplicativos

                var_clssFuncoesGoogleViagens.abrirGoogleViagens(var_strURLGoogleViagens)


            except BusinessRuleException as exception:
                self.var_clssMaestro.write_log(arg_strMensagemLog="Erro de negócio: " + str(exception), arg_enumLogLevel=LogLevel.ERROR, arg_enumErrorType=ErrorType.BUSINESS_ERROR)

                raise
            except Exception as exception:
                self.var_clssMaestro.write_log(arg_strMensagemLog="Erro, tentativa " + (var_intTentativa+1).__str__() + ": " + str(exception), arg_enumLogLevel=LogLevel.ERROR, arg_enumErrorType=ErrorType.APP_ERROR)

                if(var_intTentativa+1 == var_intMaxTentativas): raise
                else: 
                    #inclua aqui seu código para tentar novamente
                    var_clssFuncoesGoogleViagens.abrirGoogleViagens(var_strURLGoogleViagens)
                    
                    continue
            else:
                self.var_clssMaestro.write_log("Aplicativos iniciados, continuando processamento...")
                break
            
