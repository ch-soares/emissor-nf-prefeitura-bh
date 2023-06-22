from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
from time import sleep
from google_sheets.sheets import emite_nf as nf, salvar_nfs_emitidas, atualizar_planilha
import credenciais
import os


class EmissorNf:
    def __init__(self):
        opcoes = Options()
        opcoes.add_argument(f'--user-data-dir={os.environ["HOME"]}/.config/google-chrome')
        self.navegador = webdriver.Chrome(executable_path=f'{os.getcwd()}/chromedriver/chromedriver.exe', options=opcoes)
        self.espera = WebDriverWait(self.navegador, 10)
        self.url = credenciais.URL_NF
        self.navegador.get(self.url)
        self.dados = nf()

    def login(self):
        """
        Tela de autenticação
        """
        pyautogui.press('esc')
        autenticacao = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_AUTENTICACAO))
        autenticacao.click()

        try:
            """
            Tentativa em caso do login não estar no cache
            """
            login = self.navegador.find_element(By.XPATH, credenciais.ELEMENTO_LOGIN)
            login.send_keys(credenciais.LOGIN_NF)
            senha = self.navegador.find_element(By.XPATH, credenciais.ELEMENTO_SENHA)
            senha.send_keys(credenciais.SENHA_NF)
            entrar = self.navegador.find_element(By.XPATH, credenciais.ELEMENTO_ENTRAR)
            entrar.click()
        except:
            pass

    def esperar_elemento(self, localizador):
        return self.espera.until(EC.visibility_of_element_located(localizador))

    def carregar_certificado_digital(self):
        return sleep(5)

    def gerar_nf(self):
        self.login()
        self.nfs_emitidas = []
        for dado in self.dados:
            try:
                # Tela geração da nf
                geracao_nf = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_INCLUIR_NF))
                geracao_nf.click()
                data_emissao = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_DATA_EMISSAO))
                data_emissao.send_keys(str(dado['DATA']))
                confirmar_data = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_CONFIRMAR_DATA_EMISSAO))
                confirmar_data.click()

                # Tela dados do cliente
                self.carregar_certificado_digital()
                cpf_click = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_CPF))
                cpf_click.click()
                cpf_input = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_CPF))
                cpf_input.send_keys(dado['CPF'])
                cliente = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_CLIENTE))
                cliente.send_keys(dado['CLIENTE'])
                cep_click = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_CEP))
                cep_click.click()
                cep_input = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_CEP))
                cep_input.send_keys(dado['CEP'])
                endereco = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_ENDERECO))
                endereco.send_keys(dado['ENDERECO'])
                numero = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_NUMERO))
                numero.send_keys(dado['NUMERO'])
                bairro = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_BAIRRO))
                bairro.send_keys(dado['BAIRRO'])

                if dado['MUNICIPIO'] != "Belo Horizonte":
                    pesquisa_municipio = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_PESQUISA_MUNICIPIO))
                    pesquisa_municipio.click()
                    municipio = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_MUNICIPIO))
                    municipio.send_keys(dado['MUNICIPIO'])
                    pyautogui.press('enter')
                    escolhe_municipio = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_ESCOLHE_MUNICIPIO))
                    escolhe_municipio.click()

                # Tela identificação serviço
                tela_identificacao_servico = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_TELA_SERVICO))
                tela_identificacao_servico.click()

                servico = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_DESCRICAO_SERVICO))
                servico.send_keys(credenciais.DISCRIMINACAO_SERVICO)
                regime_tributacao = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_REGIME_TRIBUTACAO))
                regime_tributacao.click()
                pyautogui.write(credenciais.REGIME_TRIBUTACAO)
                sleep(1)

                # Tela de valores
                pyautogui.hotkey('ctrl', 'f')
                pyautogui.write('Valores')
                pyautogui.press('esc')
                pyautogui.press('enter')
                valor = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_VALOR))
                valor.send_keys(dado['VALOR'])
                gerar_nf = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_GERAR_NF))
                gerar_nf.click()

                # Tela leitura dos valores e confirmação da emissão
                assinatura = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_ASSINATURA_NF))
                assinatura.click()
                numero_nf = self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_NUMERO_NF)).text
                numero_nf_tratado = numero_nf[3:]
                nf = [dado['ID'], dado['DATA'], dado['CPF'], dado['CLIENTE'], dado['CEP'],
                    dado['ENDERECO'], dado['NUMERO'], dado['BAIRRO'], dado['MUNICIPIO'],
                    dado['UF'], dado['VALOR'], dado['EMAIL'], numero_nf_tratado]
                self.nfs_emitidas.append(nf)
                self.esperar_elemento((By.XPATH, credenciais.ELEMENTO_PAGINA_PRINCIPAL)).click()
                self.carregar_certificado_digital()

            except Exception:
                nf_nao_emitida = ' '.join([dado['DATA'], dado['CPF'], dado['CLIENTE']])
                print(f'NF não emitida: {nf_nao_emitida}')
                salvar_nfs_emitidas(self.nfs_emitidas)
                atualizar_planilha()

        salvar_nfs_emitidas(self.nfs_emitidas)
        atualizar_planilha()
        self.navegador.quit()


if __name__ == '__main__':
    emissor = EmissorNf()
    emissor.gerar_nf()
