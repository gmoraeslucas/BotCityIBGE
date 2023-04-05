import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

import openpyxl
from botcity.core import DesktopBot
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

servico = Service(ChromeDriverManager().install())
class Bot(DesktopBot):
    def action(self, execution=None):
        navegador = webdriver.Chrome(service=servico)

        estados = ["Acre", "Alagoas", "Amapá", "Amazonas", "Bahia", "Ceará", "Distrito Federal", "Espírito Santo",
                   "Goiás", "Maranhão", "Mato Grosso", "Mato Grosso do Sul",
                   "Minas Gerais", "Paraná", "Paraíba", "Pará", "Pernambuco", "Piauí", "Rio Grande do Norte",
                   "Rio Grande do Sul", "Rio de Janeiro",
                   "Rondônia", "Roraima", "Santa Catarina", "Sergipe", "São Paulo", "Tocantins"]

        estadosNav = ["acre", "alagoas", "amapa", "amazonas", "bahia", "ceara", "distrito", "espirito",
                   "goias", "maranhao", "mato", "matogrosso",
                   "minas", "parana", "paraiba", "para", "pernambuco", "piaui", "rionorte",
                   "riosul", "rio",
                   "rondonia", "roraima", "santa", "sergipe", "saop", "tocantins"]

        gentilico = []
        capital = []
        governador = []
        populacaoEstimada = []
        idh = []

        navegador.get("https://cidades.ibge.gov.br/")

        book = openpyxl.Workbook()
        book.create_sheet('Estados')
        estados_page = book.active
        estados_page.title = 'Estados'
        estados_page.append(['Estado', 'Gentilico', 'Capital', 'Governador', 'População Estimada', 'IDH'])

        self.wait(1000)
        if not self.find( "comece", matching=0.97, waiting_time=10000):
            self.not_found("comece")
        self.click()
        
        if not self.find( "estados", matching=0.97, waiting_time=10000):
            self.not_found("estados")
        self.click()

        gentilico = [""] * 27
        capital = [""] * 27
        governador = [""] * 27
        populacaoEstimada = [""] * 27
        idh = [""] * 27
        for i in range(27):
            if i > 19:
                self.scrollDown(1000)
                self.wait(1000)
            if not self.find( estadosNav[i], matching=0.97, waiting_time=10000):
                self.not_found(estadosNav[i])
            self.click()
            self.wait(2000)
            if i > 0:
                if not self.find( "populacaoclick", matching=0.97, waiting_time=10000):
                    self.not_found("populacaoclick")
                self.click()
            self.wait(5000)
            p1 = navegador.find_element('xpath', '//*[@id="dados"]/panorama-resumo/div/div[1]/div[2]/div/p')
            p2 = navegador.find_element('xpath', '//*[@id="dados"]/panorama-resumo/div/div[1]/div[3]/div/p')
            p3 = navegador.find_element('xpath', '//*[@id="dados"]/panorama-resumo/div/div[1]/div[4]/div/p')
            p4 = navegador.find_element('xpath', '//*[@id="dados"]/panorama-resumo/div/table/tr[2]/td[3]/valor-indicador/div/span[1]')
            gentilico[i] = p1.text
            capital[i] = p2.text
            governador[i] = p3.text
            populacaoEstimada[i] = p4.text
            if not self.find( "economia", matching=0.97, waiting_time=10000):
                self.not_found("economia")
            self.click()
            p5 = navegador.find_element('xpath', '//*[@id="dados"]/panorama-resumo/div/table/tr[41]/td[3]/valor-indicador/div/span[1]')
            idh[i] = p5.text
            estados_page.append([estados[i], gentilico[i], capital[i], governador[i], populacaoEstimada[i], idh[i]])

            if not self.find( "selecionarlocal", matching=0.97, waiting_time=10000):
                self.not_found("selecionarlocal")
            self.click()
            
            if not self.find("estados", matching=0.97, waiting_time=10000):
                self.not_found("estados")
            self.click()

            def not_found(self, label):
                print(f"Element not found: {label}")

        book.save('ListaEstados.xlsx')
        enviar_email()

    def not_found(self, label):
        print(f"Element not found: {label}")


def enviar_email():

    df = pd.read_excel('')

    de_email = ''
    para_email = ''
    senha = ''
    assunto = 'Arquivo Excel'
    mensagem = 'Segue em anexo o arquivo Excel'

    msg = MIMEMultipart()
    msg['From'] = de_email
    msg['To'] = para_email
    msg['Subject'] = assunto

    msg.attach(MIMEText(mensagem, 'plain'))

    arquivo = ''
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(open(arquivo, 'rb').read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{arquivo}"')
    msg.attach(part)

    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(de_email, senha)
    texto_email = msg.as_string()
    server.sendmail(de_email, para_email, texto_email)
    server.quit()

    print('Email enviado com sucesso!')

if __name__ == '__main__':
    Bot.main()














