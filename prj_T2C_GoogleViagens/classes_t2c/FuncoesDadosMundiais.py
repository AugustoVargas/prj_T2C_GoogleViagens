from botcity.web import WebBot, Browser,By
from botcity.core import DesktopBot
from prj_T2C_GoogleViagens.classes_t2c.utils.T2CMaestro import T2CMaestro, LogLevel, ErrorType
from prj_T2C_GoogleViagens.classes_t2c.utils.T2CExceptions import BusinessRuleException
from prj_T2C_GoogleViagens.classes_t2c.sqlite.T2CSqliteQueue import T2CSqliteQueue


class FuncoesDadosMundiais:
    def __init__(self, arg_dictConfig:dict, arg_clssMaestro:T2CMaestro, arg_botWebbot:WebBot=None, arg_botDesktopbot:DesktopBot=None,arg_clssSqliteQueue:T2CSqliteQueue=None):
        '''
        Classe com funções que serão utilizadas para navegação e realização de operações dentro do portal Dados Mundiais    

        Parâmetros:
        - arg_dictConfig (dict): dicionário de configuração.
        - arg_clssMaestro (T2CMaestro): instância de T2CMaestro.
        - arg_botWebbot (WebBot): instância de WebBot (opcional, default=None).
        - arg_botDesktopbot (DesktopBot): instância de DesktopBot (opcional, default=None).
        - arg_clssSqliteQueue (T2CSqliteQueue): instância de T2CSqliteQueue (opcional, default=None).
        
        '''
        if(arg_botWebbot is None and arg_botDesktopbot is None): raise Exception("Não foi possível inicializar a classe, forneça pelo menos um bot")
        else:
            self.var_botWebbot = arg_botWebbot
            self.var_botDesktopbot = arg_botDesktopbot
            self.var_dictConfig = arg_dictConfig
            self.var_clssMaestro = arg_clssMaestro
            self.var_clssSqliteQueue = arg_clssSqliteQueue

    def abrirDadosMundiais(self, arg_strURLDadosMundiais):
        """
        Função que realiza o acesso ao site Dados Mundiais
        
        Parâmetros:
            - arg_strURLDadosMundiais: Url de acesso ao site Dados Mundiais        
        Retorna:
        """
        try:                    
            #realiza o acesso ao site Dados Mundiais
            self.var_clssMaestro.write_log("Abrindo o portal Dados Mundiais")
           
            #abre o site do Dados Mundiais
            self.var_botWebbot.browse(arg_strURLDadosMundiais)
            self.var_botWebbot.maximize_window()

            var_intCont = 0
            var_xpathPaginaInicial = self.var_botWebbot.find_element('.//img[@src="//cdn.dadosmundiais.com/pics/logo-pt750.png"]', By.XPATH)            

            #aguardando carregamento do site
            while not var_xpathPaginaInicial and var_intCont<10:
                    self.var_clssMaestro.write_log("carregando o site Dados Mundiais, aguarde.....")
                    self.var_botWebbot.wait(10000)    
                    var_intCont = var_intCont + 1

                    #verifica se o contador de tentativas é igual a 10
                    if var_intCont >= 10:
                        raise Exception("falha ao carregar o site Dados Mundiais")
                    
            self.var_clssMaestro.write_log("Dados Mundiais, carregado com sucesso!")       

          

        except Exception as exception:              
            raise