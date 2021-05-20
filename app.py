from seleniumrequests import Chrome
#para controlar o install e permissão do ChromeDriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import xlrd
import time

class app:
    def __init__(self):
        self.driver = Chrome(ChromeDriverManager().install())
        self.url = 'http://93.188.161.218/'
        self.user = ''
        self.password = ''
        self.arr_dados_planilha = []

        #False apenas para gerar o .exe
        if(False):
            self.caminho_planilha_exe = 'planilha_insert.xls'
        else:
            self.caminho_planilha_exe = '../../planilha_insert.xls'

        self.ler_planilha()

    def iniciar(self):
        print(' \n ------------------------- Iniciando aplicação... ------------------------- \n')
        self.input_html()
        print('\n -------------------------  Finalizando aplicação... -------------------------')

    def input_html(self):
        self.login()
        for dados_planilha in self.arr_dados_planilha:
            self.input_dados(dados_planilha)
            time.sleep(2)

    def ler_planilha(self):
        planilha = xlrd.open_workbook(self.caminho_planilha_exe)
        dados = planilha.sheet_by_index(0)
        row_skip_line = 0
        for coluna in range(dados.nrows):
            if row_skip_line == 0:
                row_skip_line = 1
            else:
                self.arr_dados_planilha.append(dados.row(coluna))
        return self.arr_dados_planilha

    def login(self):
        self.driver.get(self.url)
        time.sleep(3)
        if self.user and self.password:
            self.driver.find_element_by_id('login').send_keys(self.user)
            self.driver.find_element_by_id('password').send_keys(self.password)
            self.driver.find_element_by_name('avancar').send_keys(Keys.ENTER)
            time.sleep(1)
            self.driver.get(f'{self.url}gerenciador/lancarDespesa.php')
            time.sleep(1)
        else:
            print('Usuário ou Senha não informado!')
            exit()

    def input_dados(self, arr_dados):
        planilha = xlrd.open_workbook(self.caminho_planilha_exe)
        dt_venc = xlrd.xldate.xldate_as_datetime(arr_dados[4].value, planilha.datemode)
        dt_venc = datetime.strftime(dt_venc, '%d/%m/%Y')

        self.driver.find_element_by_id('nome').send_keys(str(arr_dados[0].value))
        self.driver.find_element_by_id('cpfOuCnpj').send_keys(str(arr_dados[1].value))
        self.driver.find_element_by_id('val').send_keys(str(arr_dados[2].value))
        self.driver.find_element_by_id('obs').send_keys(str(arr_dados[3].value))
        self.driver.find_element_by_id('venc').send_keys(str(dt_venc))
        self.driver.find_element_by_id('con').send_keys(str(arr_dados[5].value))
        self.driver.find_element_by_id('ndoc').send_keys(str(arr_dados[6].value))
        self.driver.find_element_by_name('add').send_keys(Keys.ENTER)
        time.sleep(3)

app = app()
app.iniciar()

#cxfreeze .\app.py