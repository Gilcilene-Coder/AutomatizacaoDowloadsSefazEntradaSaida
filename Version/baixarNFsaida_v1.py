from tkinter import ttk
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
from threading import Thread
import re
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
import time
from selenium import webdriver
import requests
import shutil
import zipfile
import os
import tkinter as tk
from tkinter import messagebox as mbox
import sys


def download_chromedriver(url):
    # Fazendo o download do chromedriver
    response = requests.get(url, stream=True)
    file_name = url.split("/")[-1]
    with open(file_name, 'wb') as out_file:
        shutil.copyfileobj(response.raw, out_file)
    del response

def extract_chromedriver(file_name, directory):
    # Extraindo o arquivo .zip
    with zipfile.ZipFile(file_name, 'r') as zip_ref:
        zip_ref.extractall(directory)
    # Removendo o arquivo .zip após a extração
    os.remove(file_name)

chromedriver_url = "https://chromedriver.storage.googleapis.com/109.0.5414.74/chromedriver_win32.zip"
download_chromedriver(chromedriver_url)
extract_chromedriver("chromedriver_win32.zip", "./") # or any directory you want

class UserInputApplication(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)

        self.title("Coleta de Informações")
        self.geometry("350x490")  # Dimensões da janela
        self.nfe_option = ""
        self.nf_model_option = ""
        self.username_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.start_date_var = tk.StringVar()
        self.end_date_var = tk.StringVar()
        self.state_registration_number_var = tk.StringVar()

        # Estilo
        style = ttk.Style()
        style.configure("TLabel", foreground="black", font=("Vani", 10))
        style.configure("TEntry", foreground="black", background="white", font=("Arial", 10))
        style.configure("TButton", foreground="black", background="white", font=("Arial", 10, "bold"))

        frame = ttk.Frame(self)
        frame.pack(pady=10)  # Centraliza o frame verticalmente

        # Campos e labels
        ttk.Label(frame, text="Usuário").pack()
        self.username_entry = ttk.Entry(frame, textvariable=self.username_var)
        self.username_entry.pack(pady=6)

        ttk.Label(frame, text="Senha").pack()
        self.password_entry = ttk.Entry(frame, show="*", textvariable=self.password_var)
        self.password_entry.pack(pady=6)

        ttk.Label(frame, text="Data Inicial").pack()
        self.start_date_entry = ttk.Entry(frame, textvariable=self.start_date_var)
        self.start_date_entry.pack(pady=6)

        ttk.Label(frame, text="Data Final").pack()
        self.end_date_entry = ttk.Entry(frame, textvariable=self.end_date_var)
        self.end_date_entry.pack(pady=6)

        ttk.Label(frame, text="Inscrição Estadual").pack()
        self.state_registration_number_entry = ttk.Entry(frame, textvariable=self.state_registration_number_var)
        self.state_registration_number_entry.pack(pady=6)

        self.entry_fields = [
            self.username_entry,
            self.password_entry,
            self.start_date_entry,
            self.end_date_entry,
            self.state_registration_number_entry
        ]

        ttk.Label(frame, text="Downloads de NF-e").pack()
        self.download_option = ttk.Combobox(frame,
                                            values=[
                                                "Entrada",
                                                "Saída"],
                                            state="readonly")
        self.download_option.pack(pady=6)

        ttk.Label(frame, text="Modelo de NF").pack()
        self.nf_model_option = ttk.Combobox(frame,
                                            values=[
                                                "Todos",
                                                "Modelo 55",
                                                "Modelo 65"],
                                            state="readonly")
        self.nf_model_option.pack(pady=6)

        submit_button = ttk.Button(frame, text="Enviar", command=self.submit)
        submit_button.pack(pady=10)

    def format_date(self, date_str):
        if len(date_str) != 8:
            return date_str
        else:
            formatted_date = date_str[:2] + "/" + date_str[2:4] + "/" + date_str[4:]
            return formatted_date

    def validate_form(self, values):
        errors = []

        # Verifica se todos os campos foram preenchidos
        if not all(value is not None and value != '' for value in values):
            errors.extend([0, 1, 2, 3, 4,5,6])  # Índices dos campos com erro
            mbox.showerror("Erro", "Todos os campos devem ser preenchidos.")
            return errors
        # Verifica se o campo "inscrição" contém apenas números
        if not re.match("^\d+$", values[4]):
            errors.append(4)  # Índice do campo "Inscrição Estadual"

        if not self.download_option.get():
            errors.append(5)  # Índice do campo "Downloads de NFe"
        if not self.nf_model_option.get():
            errors.append(6)  # Índice do campo "Modelo de NF"

        if errors:
            error_message = "Os seguintes campos apresentam erros:\n"
            for idx in errors:
                if idx == 0:
                    error_message += "- Usuário\n"
                elif idx == 1:
                    error_message += "- Senha\n"
                elif idx == 4:
                    error_message += "- Inscrição Estadual\n"
                elif idx == 5:
                    error_message += "- Downloads de NFe\n"
                elif idx == 6:
                    error_message += "- Modelo de NF\n"

            mbox.showerror("Erro", error_message)
            return errors

        return errors  # Lista vazia indica que não há erros

    def download_notes(self):
        for _ in range(2):
            baixarxmlnfe2 = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="content"]/div[2]/div/button')))
            baixarxmlnfe2.click()
            time.sleep(15)

            try:
                try:
                    baixar1 = self.driver.find_element(By.ID, 'dnwld-all-btn-ok')
                    baixar1.click()
                    time.sleep(15)

                    # Verifica se o elemento okxml está disponível dentro de 5 segundos
                    WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located(
                            (By.XPATH, '//button[contains(@class, "dwnld-loading-window-cancel")]'))
                    )
                except:
                    cancelar = self.driver.find_element(By.XPATH, '/html/body/div[6]/div/div/div[1]/button/span[1]')
                    self.driver.execute_script("arguments[0].click();", cancelar)
                    time.sleep(15)
                    continue

                okxml = self.driver.find_element(By.XPATH, '//button[contains(@class, "dwnld-loading-window-cancel")]')
                okxml.click()
                time.sleep(15)
            except Exception as e:
                print(f"Ocorreu um erro: {e}")
                continue
            break

        self.exibir_mensagem_conclusao()

    def exibir_mensagem_conclusao(self):
        self.driver.quit()
        mbox.showinfo("Processo Concluído", "A coleta de informações foi concluída com sucesso!")
        self.perguntar_nova_consulta()

    def perguntar_nova_consulta(self):
        resposta = mbox.askyesno("Nova Consulta", "Deseja realizar uma nova consulta?")
        if resposta:
                self.username_entry.delete(0, tk.END)
                self.password_entry.delete(0, tk.END)
                self.start_date_entry.delete(0, tk.END)
                self.end_date_entry.delete(0, tk.END)
                self.state_registration_number_entry.delete(0, tk.END)
                self.download_option.set("")
                self.nf_model_option.set("")
        else:
            sys.exit()

    def submit(self):
        self.values = [
            self.username_entry.get(),
            self.password_entry.get(),
            self.format_date(self.start_date_entry.get()),
            self.format_date(self.end_date_entry.get()),
            self.state_registration_number_entry.get()
        ]

        # Limpa a formatação anterior de campos com erro
        for entry in self.entry_fields:
            entry.configure(style="TEntry")

        errors = self.validate_form(self.values)

        if errors:
            for idx in errors:
                self.highlight_error_widget(idx)
        else:
            self.start_process()

    def send_esc_after_delay(self, delay):
        time.sleep(delay)
        pyautogui.press('esc')

    def common_operations(self):
        esc_thread = Thread(target=self.send_esc_after_delay, args=(10,))
        esc_thread.start()
        self.driver.get("https://portal.sefaz.go.gov.br/portalsefaz-apps/auth/login-form")

    def start_process(self):
        # substitua o caminho para o chromedriver se não estiver na mesma pasta do script
        chrome_driver_path = os.path.join(os.getcwd(), 'chromedriver.exe')  # define chrome_driver_path

        # Inicializa o WebDriver.
        self.driver = webdriver.Chrome(executable_path=chrome_driver_path)
        self.wait = WebDriverWait(self.driver, 5)

        # Define os valores a serem usados.
        values = self.values
        nfe_option = self.download_option.get()  # Valor do combobox
        nf_model_option = self.nf_model_option.get()  # Valor do combobox

        try:
            if nfe_option == "Saída":
                # Executa a operação somente para saída.
                self.execute_operation(values, select_saida=True, nf_model_option=nf_model_option)
            else:
                # Executa a operação somente para entrada.
                self.execute_operation(values, select_saida=False, nf_model_option=nf_model_option)
        except Exception as e:
            print(e)
        finally:
            self.driver.quit()

    def execute_operation(self, values, select_saida, nf_model_option):
        self.common_operations()
        nfe_option = self.download_option.get()  # Obtenha o valor selecionado no combobox

        login = self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="username"]')))
        login.send_keys(values[0])
        time.sleep(5)

        senha = self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="password"]')))
        senha.send_keys(values[1])
        time.sleep(5)

        entrar = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnAuthenticateText"]')))
        entrar.click()
        time.sleep(5)

        self.driver.get('https://www.sefaz.go.gov.br/NETACCESS/default.asp')
        time.sleep(5)

        iframe = self.driver.find_element(By.XPATH, '//*[@id="iNetaccess"]')
        self.driver.switch_to.frame(iframe)
        time.sleep(5)

        baixarxmlnfe = self.driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/ul/li[2]/a')
        baixarxmlnfe.click()
        time.sleep(5)

        login2 = self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="NetAccess.Login"]')))
        login2.send_keys(values[0])
        time.sleep(5)

        senha2 = self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="NetAccess.Password"]')))
        senha2.send_keys(values[1])
        time.sleep(5)

        autenticador2 = self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="btnAuthenticate"]')))
        autenticador2.click()
        time.sleep(5)

        iframe = self.driver.find_element(By.XPATH, '//*[@id="iNetaccess"]')
        self.driver.switch_to.frame(iframe)
        time.sleep(5)

        baixarxmlnfe2 = self.driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/ul/li[2]/a')
        baixarxmlnfe2.click()
        time.sleep(5)

        datainicial = self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="cmpDataInicial"]')))
        self.driver.execute_script("arguments[0].value = '{}';".format(values[2]), datainicial)
        time.sleep(5)

        datafinal = self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="cmpDataFinal"]')))
        self.driver.execute_script("arguments[0].value = '{}';".format(values[3]), datafinal)
        time.sleep(5)

        inscricaoestadual = self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="cmpNumIeDest"]')))
        self.driver.execute_script("arguments[0].value = '{}';".format(values[4]), inscricaoestadual)
        time.sleep(5)

        if select_saida and nfe_option == 'Saída':
            saida1 = self.driver.find_element(By.XPATH,
                                              '/html/body/div[4]/div/div[1]/div[2]/form/div[3]/div/div[2]/input')
            saida1.click()
            time.sleep(5)

        modelodeNf = self.driver.find_element(By.XPATH,
                                              '/html/body/div[4]/div/div[1]/div[2]/form/div[4]/div/select')
        modelodeNf.click()
        time.sleep(5)

        if nf_model_option == "Todos":
            modelodeNftodos = self.driver.find_element(By.XPATH,
                                                       '/html/body/div[4]/div/div[1]/div[2]/form/div[4]/div/select/option[1]')
            modelodeNftodos.click()
            time.sleep(5)
        elif nf_model_option == "Modelo 55":
            modelodeNf55 = self.driver.find_element(By.XPATH,
                                                    '/html/body/div[4]/div/div[1]/div[2]/form/div[4]/div/select/option[2]')
            modelodeNf55.click()
            time.sleep(5)
        else:
            modelodeNf65 = self.driver.find_element(By.XPATH,
                                                    '/html/body/div[4]/div/div[1]/div[2]/form/div[4]/div/select/option[3]')
            modelodeNf65.click()
            time.sleep(5)

        notasCanceladas = self.driver.find_element(By.XPATH,
                                                   '/html/body/div[4]/div/div[1]/div[2]/form/div[5]/div/div/input')
        notasCanceladas.click()
        time.sleep(15)

        try:
            pesquisarInscricao = self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="btnPesquisar"]')))
            pesquisarInscricao.click()
            time.sleep(10)

            # Adicionar a verificação aqui
            try:
                erroInscricao = self.wait.until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="panel"]/div[3]/div/button')))
                mbox.showerror("Erro",
                               "A inscrição não tem notas fiscais ou não está habilitado para consulta. Por favor, verifique e tente novamente.")
                return  # Retorna do método para evitar executar o resto do código
            except TimeoutException:
                pass  # Nenhum erro encontrado, continuar com o processo normal

            # Tente localizar o botão de download. Se não conseguir, suponha que nenhuma nota foi encontrada.
            baixarxmlnfe2 = self.wait.until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="content"]/div[2]/div/button')))
            baixarxmlnfe2.click()

        except Exception as e:
            print(e)
            mbox.showerror("Erro",
                           "Não foi possível realizar a consulta. Verifique as seguintes possíveis causas: \n- Período de Consulta Inválido \n- Inscrição Estadual Incorreta.")

        self.download_notes()
def run_application():
    while True:
        app = UserInputApplication()
        app.mainloop()
        if not app.winfo_exists():
            break

if __name__ == "__main__":
    run_application()