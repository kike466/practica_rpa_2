import time
from openpyxl import load_workbook
from selenium.webdriver import ActionChains, Keys
from selenium.webdriver.common.by import By

from usuarios import usuario
from selenium import webdriver

#Se abre el libro del excel y se extraen el DNI y la matriula
wb = load_workbook(filename='suma.xlsx')
sheet_ranges = wb['Suma']

usrs = []
for c1, c2 in sheet_ranges['A2':'B13']:
    DNI = c1.value
    Matricula = c2.value

    #Se crea un usuario
    usr = usuario(DNI, Matricula)
    #Se añade a una lista
    usrs.append(usr)
#cerramos el libro
wb.close()

#iniciamos el driver de chromedriver
driver = webdriver.Chrome(r'C:\Users\Usuario\Downloads\chromedriver_win32\chromedriver.exe')
#maximizamos la ventana
driver.maximize_window()
#accedemos a una url
driver.get("https://www.suma.es/")
time.sleep(1)

#action chains nos permite interactuar con los elementos
achain = ActionChains(driver)
#hacemos el hover del contribuyentes
elemento = driver.find_element(By.XPATH, '//*[@id="navbarNavDropdown"]/ul/li[1]')
achain.move_to_element(elemento).click().perform()

time.sleep(1)
# click en obtener recibo
driver.find_element(By.LINK_TEXT, 'Obtener un recibo').click()
time.sleep(1)
# click en Con el numero fijo
driver.find_element(By.LINK_TEXT, 'Con el número fijo (periodo voluntario)').click()
#se ban buscando resultados para cada usuario
resultado=[]
for usuario in usrs:
    #hacemos el clear de los input
    driver.find_element(By.ID, 'pantalla:nif').clear()
    driver.find_element(By.ID, 'pantalla:cprv_numerofijo').clear()
    #añadimos datos a los inputs
    elemento_DNI = driver.find_element(By.ID, 'pantalla:nif')
    elemento_Matricula = driver.find_element(By.ID, 'pantalla:cprv_numerofijo')

    elemento_DNI.send_keys(usuario.get_DNI())
    elemento_Matricula.send_keys(usuario.get_Matricula())

    time.sleep(1)
    #click en el enviar
    driver.find_element(By.ID, 'pantalla:obtenerRecibos').click()
    time.sleep(1)

    msg=''
#error de no se encuentra
    try:
        if driver.find_element(By.XPATH, '//*[@id="pantalla"]/div/div[4]/div[2]/div/span').is_displayed():
            msg='Error '+ driver.find_element(By.XPATH, '//*[@id="pantalla"]/div/div[4]/div[2]/div/span').text
            resultado.append(msg)
            print(msg)
    except:
        print("An exception occurred 1")
#error de que los datos introducidos no cumplen los requisitos
    try:

        if driver.find_element(By.XPATH, '//*[@id="pantalla"]/div/div[4]/div/div[1]/div[2]/span').is_displayed():
            msg = 'Error ' + driver.find_element(By.XPATH, '//*[@id="pantalla"]/div/div[4]/div/div[1]/div[2]/span').text
            resultado.append(msg)
            print(msg)

    except:
        print("An exception occurred 2")

#en caso de que todo este bien
    try:
       if msg=='':
            msg = 'OK'
            print(msg)
            resultado.append(msg)
            time.sleep(1)
            #descargar pdf
            driver.find_element(By.XPATH, '//*[@id="pantalla:listaVallistaValores"]/tbody/tr/td[9]/div/a').click()

            driver.find_element(By.XPATH, '//*[@id="panelbotones"]/div[3]/img').click()

            driver.find_element(By.XPATH, '//*[@id="pantalla:tabladocumentosGroup"]/table[1]/tbody/tr/td[3]/a/i').click()
    except:
        print("An exception occurred 3")


#Se abre el libro para poner los mensajes de resultado
wb = load_workbook(filename='suma.xlsx')
sheet_ranges = wb['Suma']
#sheet_ranges['C'+str(3)]='hola'
#sheet_ranges['C4']='adios'

count=2
print("--------------------------------")
for items in resultado:
    #se insertan los valores de los resultados
    sheet_ranges['C'+str(count)]=items
    count+=1
    print(count)
#se guardan los cambios
wb.save('suma.xlsx')
#se cierra el libro
wb.close()
driver.close()



