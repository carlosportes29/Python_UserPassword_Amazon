from selenium import webdriver
import time
import logging

browser = webdriver.Chrome("C:/Petro_py/chromedriver.exe")  # Path to where I installed the web driver

browser.get('https://www.amazon.com.br/ap/signin?openid.pape.max_auth_age=0&openid.return_to=https%3A%2F%2Fwww.amazon.com.br%2F%3Fref_%3Dnav_ya_signin&openid.identity=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.assoc_handle=brflex&openid.mode=checkid_setup&openid.claimed_id=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.ns=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0&');
browser.maximize_window()#maximizar a janela do chrome
time.sleep(1)
##box email
USER_MAIL = browser.find_element("name", "email")
USER_MAIL.send_keys('carlosmst1@gmail.com')
##botão continuar
BTN_CONTINUAR = browser.find_element("id", "continue");
BTN_CONTINUAR.submit();
time.sleep(1)
##box senha
USER_PASSWORD = browser.find_element("name", "password")
USER_PASSWORD.send_keys('redigir_senha')

##fazer Login
BTN_LOGAR = browser.find_element("id", "signInSubmit");
BTN_LOGAR.submit();
##validar se o usuário está logado
# valida_usuario_logado = browser.find_element("/html/body/div[1]/header/div/div[1]/div[3]/div/a[1]/div/span").is_displayed(True);
time.sleep(3); # Let the user actually see something!


browser.quit();
