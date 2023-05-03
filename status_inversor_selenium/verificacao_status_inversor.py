import win32com.client as win32

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time
from datetime import datetime

nome = '********'
login = '********'
senha = '********'

nome1 = '********'
login1 = '********'
senha1 = '********'

nome2 = '********'
login2 = '********'
senha2 = '********'

nome3 = '********'
login3 = '********'
senha3 = '********'

nome4 = '********'
login4 = '********'
senha4 = '********'

nome5 = '********'
login5 = '********'
senha5 = '********'

nome6 = '********'
login6 = '********'
senha6 = '********'

nome7 = '********'

nome8 = '********'

#modo = "--headless" ou "--handless"
modo1 = '--headless'
modo2 = '--handless'


def growatt_residencial(nome, login, senha, modo):

    # Iniciar Webdriver
    options = Options()
    #options.add_argument("--headless")
    #options.add_argument("--handless")
    options.add_argument(modo)
    driver = webdriver.Chrome(chrome_options=options)

    # Acessar endereço
    driver.get("https://server.growatt.com/login")

    # Esperar até a caixa de busca aparecer
    searchField = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "[id=val_loginAccount]"))
    )

    searchField.send_keys(login)

    time.sleep(1)

    searchField = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "[id=val_loginPwd]"))
    )

    searchField.send_keys(senha)

    searchField.send_keys(Keys.ENTER)

    time.sleep(10)

    elemento1 = searchField = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, '#tb_device_con > div > table > tbody > tr:nth-child(2) > td:nth-child(3) > span > span'))
    ).text

    time.sleep(3)

    driver.quit()

    time.sleep(5)

    # ENVIANDO E-MAIL AUTOMATICAMENTE CONFORME A CONDICIONAL
    if elemento1 != 'Normal':
        # criar a integração com o outlook
        outlook1 = win32.Dispatch('outlook.application')
        # criar um email
        email = outlook1.CreateItem(0)
        # configurar as informações do seu e-mail
        email.To = "douglas.silva@motormac.com.br;"  # engenharia.solar@motormac.com.br
        email.Subject = "IMPORTANTE!"
        email.HTMLBody = f'''
                         <p>Status do inversor {nome} está anormal. Favor verificar na plataforma de monitoramento</p>
                         '''
        email.Send()

    else:
        hora = datetime.now()
        print(f'Inversor {nome} operando normalmente - {hora}')

    time.sleep(5)


def growatt_residencial2(nome, login, senha, modo):

    # Iniciar Webdriver
    options = Options()
    options.add_argument(modo)
    #options.add_argument("--handless")
    driver = webdriver.Chrome(chrome_options=options)

    # Acessar endereço
    driver.get("https://server.growatt.com/login")

    # Esperar até a caixa de busca aparecer
    searchField = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "[id=val_loginAccount]"))
    )

    searchField.send_keys(login)

    time.sleep(1)

    searchField = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "[id=val_loginPwd]"))
    )

    searchField.send_keys(senha)

    searchField.send_keys(Keys.ENTER)

    time.sleep(10)

    elemento4 = searchField = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, '#tb_device_con > div > table > tbody > tr:nth-child(2) > td:nth-child(3) > span > span'))
    ).text

    time.sleep(2)

    driver.quit()

    time.sleep(5)

    # ENVIANDO E-MAIL AUTOMATICAMENTE CONFORME A CONDICIONAL
    if elemento4 != 'Connection':
        # criar a integração com o outlook
        outlook1 = win32.Dispatch('outlook.application')
        # criar um email
        email = outlook1.CreateItem(0)
        # configurar as informações do seu e-mail
        email.To = "douglas.silva@motormac.com.br;"  # engenharia.solar@motormac.com.br #douglas.silva@motormac.com.br;mateus.dapper@motormac.com.br
        email.Subject = "IMPORTANTE!"
        email.HTMLBody = f'''
                         <p>Status do inversor {nome} está anormal. Favor verificar na plataforma de monitoramento</p>
                         '''
        email.Send()
    else:
        hora = datetime.now()
        print(f'Inversor {nome} operando normalmente - {hora}')

    time.sleep(5)




def growatt_residencial3(nome, login, senha, modo):

    # Iiciar Webdriver
    options = Options()
    options.add_argument(modo)
    #options.add_argument("--handless")
    driver = webdriver.Chrome(chrome_options=options)

    # Acessar endereço
    driver.get("https://server.growatt.com/login")

    # Esperar até a caixa de busca aparecer
    searchField = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "[id=val_loginAccount]"))
    )

    searchField.send_keys(login)

    time.sleep(1)

    searchField = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "[id=val_loginPwd]"))
    )

    searchField.send_keys(senha)

    searchField.send_keys(Keys.ENTER)

    time.sleep(7)

    driver.get("https://server.growatt.com/index")

    time.sleep(10)

    elemento7 = searchField = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, '#tb_device_con > div > table > tbody > tr:nth-child(2) > td:nth-child(3) > span > span'))
    ).text

    time.sleep(2)

    driver.quit()

    time.sleep(5)

    # ENVIANDO E-MAIL AUTOMATICAMENTE CONFORME A CONDICIONAL
    if elemento7 != 'Connection':
        # criar a integração com o outlook
        outlook1 = win32.Dispatch('outlook.application')
        # criar um email
        email = outlook1.CreateItem(0)
        # configurar as informações do seu e-mail
        email.To = "douglas.silva@motormac.com.br;"  # engenharia.solar@motormac.com.br
        email.Subject = "IMPORTANTE!"
        email.HTMLBody = f'''
                         <p>Status do inversor {nome} está anormal. Favor verificar na plataforma de monitoramento</p>
                         '''
        email.Send()
    else:
        hora = datetime.now()
        print(f'Inversor {nome} operando normalmente - {hora}')

    time.sleep(5)




def growatt_comercial(nome, login, senha, modo):

    # Iniciar Webdriver
    options = Options()
    options.add_argument(modo)
    #options.add_argument("--handless")
    driver = webdriver.Chrome(chrome_options=options)

    # Acessar endereço
    driver.get("https://server.growatt.com/login")

    # Esperar até a caixa de busca aparecer
    searchField = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "[id=val_loginAccount]"))
    )

    searchField.send_keys(login)

    time.sleep(1)

    searchField = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "[id=val_loginPwd]"))
    )
    searchField.send_keys(senha)

    searchField.send_keys(Keys.ENTER)

    time.sleep(15)

    elemento = searchField = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, '#invStatusTxt > ul > li:nth-child(1) > em'))
    ).text

    time.sleep(2)

    driver.quit()

    time.sleep(5)

    # ENVIANDO E-MAIL AUTOMATICAMENTE CONFORME A CONDICIONAL
    if elemento != '1/1':
        # criar a integração com o outlook
        outlook1 = win32.Dispatch('outlook.application')
        # criar um email
        email = outlook1.CreateItem(0)
        # configurar as informações do seu e-mail
        email.To = "douglas.silva@motormac.com.br;"  # engenharia.solar@motormac.com.br
        email.Subject = "IMPORTANTE!"
        email.HTMLBody = f'''
                         <p>Status do inversor {nome} está anormal. Favor verificar na plataforma de monitoramento</p>
                         '''
        email.Send()
    else:
        hora = datetime.now()
        print(f'Inversor {nome} operando normalmente - {hora}')

    time.sleep(5)




def growatt_comercial2(nome, login, senha, modo):

    # niciar Webdriver
    options = Options()
    options.add_argument(modo)
    #options.add_argument("--handless")
    driver = webdriver.Chrome(chrome_options=options)

    # Acessar endereço
    driver.get("https://server.growatt.com/login")

    # Esperar até a caixa de busca aparecer
    searchField = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "[id=val_loginAccount]"))
    )

    searchField.send_keys(login)

    time.sleep(1)

    searchField = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "[id=val_loginPwd]"))
    )

    searchField.send_keys(senha)

    searchField.send_keys(Keys.ENTER)

    time.sleep(7)

    driver.get("http://server.growatt.com/indexbC/getInvPage")

    time.sleep(5)

    searchField = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, '#sectionBox > div.sectionBox_right > div.sectionBox_right_top.flex_space-between > div.droplistBox.mt_15.ml_15 > span > span'))
    ).click()

    time.sleep(3)

    searchField = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "#sectionBox > div.sectionBox_right > div.sectionBox_right_top.flex_space-between > div.droplistBox.mt_15.ml_15 > div > ul > li:nth-child(1) > span"))
    ).click()

    time.sleep(8)

    elemento6 = searchField = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, '#invStatusTxt > ul > li:nth-child(1) > em'))
    ).text

    time.sleep(2)

    driver.quit()

    time.sleep(5)

    # ENVIANDO E-MAIL AUTOMATICAMENTE CONFORME A CONDICIONAL
    if elemento6 != '1/1':
        # criar a integração com o outlook
        outlook1 = win32.Dispatch('outlook.application')
        # criar um email
        email = outlook1.CreateItem(0)
        # configurar as informações do seu e-mail
        email.To = "douglas.silva@motormac.com.br;"  # engenharia.solar@motormac.com.br
        email.Subject = "IMPORTANTE!"
        email.HTMLBody = f'''
                         <p>Status do inversor {nome} está anormal. Favor verificar na plataforma de monitoramento</p>
                         '''
        email.Send()
    else:
        hora = datetime.now()
        print(f'Inversor {nome} operando normalmente - {hora}')


    time.sleep(5)




def growatt_comercial3(nome, login, senha, modo):

    # Iniciar Webdriver
    options = Options()
    options.add_argument(modo)
    #options.add_argument("--handless")
    driver = webdriver.Chrome(chrome_options=options)

    # Acessar endereço
    driver.get("https://server.growatt.com/login")

    # Esperar até a caixa de busca aparecer
    searchField = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "[id=val_loginAccount]"))
    )

    searchField.send_keys(login)

    time.sleep(1)

    searchField = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "[id=val_loginPwd]"))
    )

    searchField.send_keys(senha)
    searchField.send_keys(Keys.ENTER)

    time.sleep(7)

    driver.get("http://server.growatt.com/indexbC/getInvPage")

    time.sleep(5)

    searchField = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, '#sectionBox > div.sectionBox_right > div.sectionBox_right_top.flex_space-between > div.droplistBox.mt_15.ml_15 > span > span'))
    ).click()

    time.sleep(3)

    searchField = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "#sectionBox > div.sectionBox_right > div.sectionBox_right_top.flex_space-between > div.droplistBox.mt_15.ml_15 > div > ul > li:nth-child(2) > span"))
    ).click()

    time.sleep(8)

    elemento8 = searchField = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, '#invStatusTxt > ul > li:nth-child(1) > em'))
    ).text

    time.sleep(2)

    driver.quit()

    time.sleep(5)

    # ENVIANDO E-MAIL AUTOMATICAMENTE CONFORME A CONDICIONAL
    if elemento8 != '1/1':
        # criar a integração com o outlook
        outlook1 = win32.Dispatch('outlook.application')
        # criar um email
        email = outlook1.CreateItem(0)
        # configurar as informações do seu e-mail
        email.To = "douglas.silva@motormac.com.br;"  # engenharia.solar@motormac.com.br
        email.Subject = "IMPORTANTE!"
        email.HTMLBody = f'''
                         <p>Status do inversor {nome} está anormal. Favor verificar na plataforma de monitoramento</p>
                         '''
        email.Send()
    else:
        hora = datetime.now()
        print(f'Inversor {nome} operando normalmente - {hora}')

    time.sleep(5)

def fronius(nome, login, senha, modo):

    # Iniciar Webdriver
    options = Options()
    options.add_argument(modo)
    #options.add_argument("--handless")
    driver = webdriver.Chrome(chrome_options=options)

    # Acessar endereço
    #driver.get("https://login.fronius.com/authenticationendpoint/login.do?client_id=mf_o9iTAyKemNLQTa6Sp6HYonCIa&commonAuthCallerPath=%2Foauth2%2Fauthorize&forceAuth=false&nonce=638086257871638703.ZjMyMWRlMTUtYjczZi00Yzk3LWE5MDMtZjA1NzQ2MTE4ZWFiYzhlMjllZDktMjA5Zi00OTQ5LWFjMTgtNWM3Mzc0NDc5ODgz&passiveAuth=false&redirect_uri=https%3A%2F%2Fwww.solarweb.com%2FAccount%2FExternalLoginCallback&response_mode=form_post&response_type=code+id_token&scope=openid+profile+solarweb+solweb_browserid_93c65d8e866a6421e43a4e3d7b75f200&state=OpenIdConnect.AuthenticationProperties%3DSsAjCQwhtLM63G0X-Cxoiy_-_ow2a7qxO9yctSrggonlxzJuhUS_de7JBb9zlMsAf4LIMAlGtwAQQMWmJEamucYmOXNg8dosJDRAMK9VvKjX7-r46CFMo3Q_kTjT7BrBHFC4dHaEMRmzxG5-5C_gxMh5_-8-A-Yb8hSTmtbzIeh8PTfraQSHRdQaSSTeu8zwrD3H3iB3_E7-t_4X7b_t5pJR4woHFPoDpXuF81PAQZTAbOLiXiPpVUmO4TYK6HZ_ZqzANA&tenantDomain=carbon.super&x-client-SKU=ID_NET461&x-client-ver=6.9.0.0&sessionDataKey=35339208-7e07-4eab-ab28-4529dc8f999c&relyingParty=mf_o9iTAyKemNLQTa6Sp6HYonCIa&type=oidc&sp=Solar.web+-+Portals&isSaaSApp=false&authenticators=SAMLSSOAuthenticator:Fronius%20Login;CustomAuthenticatorLocalMain:LOCAL:LOCAL")
    #driver.get('https://login.fronius.com/authenticationendpoint/login.do?client_id=mf_o9iTAyKemNLQTa6Sp6HYonCIa&commonAuthCallerPath=%2Foauth2%2Fauthorize&forceAuth=false&nonce=638086251709867249.NzgzNTcxNDMtMTM3Zi00YTNlLWI1YTAtZmFlMTU1YjQzOTQ5ZjNjZGQyZDgtZWZhNS00ZTAyLWIyMjQtMzI5NWYzMDNjZmZk&passiveAuth=false&redirect_uri=https%3A%2F%2Fwww.solarweb.com%2FAccount%2FExternalLoginCallback&response_mode=form_post&response_type=code+id_token&scope=openid+profile+solarweb+solweb_browserid_93c65d8e866a6421e43a4e3d7b75f200&state=OpenIdConnect.AuthenticationProperties%3DScsnYzhvOIcVVBnLAKDTdMTMHlrY3YgGd61BY4So6U4__SX9UU0kZxcuvq-UhVWae8NsvX41059QeyF2m1TF_2_Yxw4deeTuilEmGq2BeIum57rC1snyKZsdaTRsnXwdqaTuumF3N_cxG4gsGezwfXtBFIMCOrn2d8tYG1UlJhqQFI_RGeXlQn3pKozeO69aJeISfMuiOKNueo9ayw9o_OrmWBA7TgDFhZ5uSA8PA30mJQIyN1YCgxn0dcuKjXPyOcQE0Q&tenantDomain=carbon.super&x-client-SKU=ID_NET461&x-client-ver=6.9.0.0&sessionDataKey=d744ee0f-04d2-4803-ba4a-0ce3dab7608d&relyingParty=mf_o9iTAyKemNLQTa6Sp6HYonCIa&type=oidc&sp=Solar.web+-+Portals&isSaaSApp=false&authenticators=SAMLSSOAuthenticator:Fronius%20Login;CustomAuthenticatorLocalMain:LOCAL:LOCAL')
    driver.get("https://login.fronius.com/authenticationendpoint/login.do?client_id=mf_o9iTAyKemNLQTa6Sp6HYonCIa&commonAuthCallerPath=%2Foauth2%2Fauthorize&forceAuth=false&nonce=638103471616323518.NWRjMGQ2MzMtZjhhNC00Y2M2LTg1NmQtZWQ4MmU1ZjdkZmI5MTlhNGRmN2UtNTUxZC00ZTdjLWE3NmYtYWY3ZDdmYjg5Mzcy&passiveAuth=false&redirect_uri=https%3A%2F%2Fwww.solarweb.com%2FAccount%2FExternalLoginCallback&response_mode=form_post&response_type=code+id_token&scope=openid+profile+solarweb+solweb_browserid_ed0eded3b89dc44c0342be52df85d30f&state=OpenIdConnect.AuthenticationProperties%3D03qcIWl3LEO3zX__EfUVeN9W0FM0UGj1285a5Xnc-JS7uahEpl0tglmmy-rpmvFp6P90oVOrcg9NcQmgXZ1rxfkzq66TdDYd5sMsFl-_8mYmrT37A4Dm-pKYYQzGVssQqKw4AFWrbAfKYWIR2LPdO9H8mYVEVIKdRjgkQUpNdw1ekFkKePLY4gc87ZaC_om5AHnk73pOYVrg3DaBxfx05wg725qNQYbTUCgLV1BrbZhmeNXXGL__smaeJETM1AvWSXAgZQ&tenantDomain=carbon.super&x-client-SKU=ID_NET461&x-client-ver=6.9.0.0&sessionDataKey=6fdb9919-b795-42bd-8f75-0392c6b10168&relyingParty=mf_o9iTAyKemNLQTa6Sp6HYonCIa&type=oidc&sp=Solar.web+-+Portals&isSaaSApp=false&authenticators=SAMLSSOAuthenticator:Fronius%20Login;CustomAuthenticatorLocalMain:LOCAL:LOCAL")

    time.sleep(10)

    searchField = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, '#CybotCookiebotDialogBodyButtonDecline'))
    ).click()

    time.sleep(3)

    searchField = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, '#navigation > ul.nav.navbar-nav.nav-pills.pull-right.landing-page-menu-right > li:nth-child(1) > a > span'))
    ).click()

    time.sleep(5)

    # Esperar até a caixa de busca aparecer
    searchField = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "#username"))
    )

    searchField.send_keys(login)

    time.sleep(1)

    searchField = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "#password"))
    )

    searchField.send_keys(senha)

    searchField.send_keys(Keys.ENTER)

    time.sleep(10)

    elemento2 = searchField = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, '#powerWidget > div > div.savings-widget-body.js-live-data-widget-content > div:nth-child(2) > div > div > span:nth-child(3)'))
    ).text

    time.sleep(3)

    driver.quit()

    time.sleep(5)

    # ENVIANDO E-MAIL AUTOMATICAMENTE CONFORME A CONDICIONAL
    if elemento2 != 'Offline':
        hora = datetime.now()
        print(f'Inversor {nome} operando normalmente - {hora}')
    else:
        # criar a integração com o outlook
        outlook1 = win32.Dispatch('outlook.application')
        # criar um email
        email = outlook1.CreateItem(0)
        # configurar as informações do seu e-mail
        email.To = "douglas.silva@motormac.com.br;"  # engenharia.solar@motormac.com.br
        email.Subject = "IMPORTANTE!"
        email.HTMLBody = f'''
                         <p>Status do inversor {nome} está anormal. Favor verificar na plataforma de monitoramento</p>
                         '''
        email.Send()

    time.sleep(5)

while True:

    growatt_residencial(nome1, login1, senha1, modo1)

    growatt_comercial(nome, login, senha, modo1)

    fronius(nome2, login2, senha2, modo1)

    growatt_residencial(nome3, login3, senha3, modo1)

    growatt_residencial2(nome4, login4, senha4, modo1)

    growatt_comercial(nome5, login5, senha5, modo1)

    growatt_comercial2(nome6, login6, senha6, modo1)

    growatt_residencial3(nome7, login6, senha6, modo1)

    growatt_comercial3(nome8, login6, senha6, modo1)

    time.sleep(3595)