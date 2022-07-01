from operator import indexOf
import os
import shutil
import eel
import time
from smtplib import SMTPAuthenticationError
import concurrent.futures
from comtypes.client import CreateObject
from dotenv import load_dotenv
import logging
from time import asctime

# Загружаем секретные переменные
load_dotenv()

path = os.getcwd()

eel.init(path + "\\Web")

@eel.expose
def create_email_info(e, p):
    # Создаем или перезаписываем файл имеющейся информацией
    with open(".env", 'w') as env:
        env.write(f"MY_ADRESS = {e}\n")
        env.write(f"MY_PASSWORD = {p}\n")
    load_dotenv()

def init_powerpoint():
    powerpoint = CreateObject('PowerPoint.Application')
    powerpoint.UserControl = 0
    powerpoint.Visible = 1
    return powerpoint

@eel.expose
def start(input_file_name, output_file_name, send):    
    
    logger = logging.getLogger("Certificater")
    logger.setLevel(logging.INFO)
    
    # create the logging file handler
    fh = logging.FileHandler("loggig_work.log")
 
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    fh.setFormatter(formatter)
    
    # add handler to logger object
    logger.addHandler(fh)
    
    logger.info("Program started")
    
    


    time1 = time.perf_counter()

    from PPTX_to_PDF import pptx_to_pdf
    from sending import login, send_email
    from all_names import all_names
    from PPTX_GENERATOR import PPTX_GENERATOR

    if send:
        try:
            os.environ["MY_ADRESS"]
        except KeyError:
            eel.raise_error("Почтовые данные не найдены, но сертификаты созданы")
            logger.warning("Почтовые данные не найдены, но сертификаты созданы!")
            send = False

    data = all_names(input_file_name, output_file_name)
    # Если Excel файл не найден
    if data == 'Excel':
        eel.raise_error('Excel файл не найден или не указан')
        logger.warning("Excel файл не найден или не указан")
        return
    elif data == 'Template':
        eel.raise_error('Шаблон не найден или не указан')
        logger.warning("Шаблон не найден или не указан")
        return

    try:
        os.makedirs(f"GENERATED_PPTX/{data[0]['date']}", exist_ok=True)
        os.makedirs(f"GENERATED_PDF/{data[0]['date']}", exist_ok=True)

        shutil.rmtree(f"GENERATED_PPTX/{data[0]['date']}")
        shutil.rmtree(f"GENERATED_PDF/{data[0]['date']}")

        os.makedirs(f"GENERATED_PPTX/{data[0]['date']}", exist_ok=True)
        os.makedirs(f"GENERATED_PDF/{data[0]['date']}", exist_ok=True)

        logger.info("Создана папка: ", f"GENERATED_PPTX/{data[0]['date']}")
        logger.info("Создана папка: ", f"GENERATED_PDF/{data[0]['date']}")
    except Exception as e:
        logger.error(e)
        exit(-20)


    # Запускаем асинхронное редактирование pptx
    #with concurrent.futures.ThreadPoolExecutor() as executor:
        #file_name = list(executor.map(PPTX_GENERATOR, data))
    
    file_name = list(map(PPTX_GENERATOR, data))
    # Устанавлиавем соединение
    if send:
        try:
            smtps = login()
            logger.info("Подключено к SMTPS успещно")
        except Exception as e:
            logger.error(e)

    # Перебираем каждый элемент в массиве
    for loc in data:
        powerpoint = init_powerpoint()
        pptx_to_pdf(file_name[indexOf(data, loc)], loc['date'], powerpoint)
        print(file_name[indexOf(data, loc)], loc['email'], ' - готов' )
        if send:
            try:
                send_email(loc['email'], smtps, loc['date'], file_name[indexOf(data, loc)])
                print(file_name[indexOf(data, loc)], loc['email'], ' - отправлен' )
            except SMTPAuthenticationError:
                eel.raise_error("Не верно указана почта или пароль от почты\nЕсли все указано верно, то ваша почта не поддержиаватся") # Сделать ссылку
                send = False
            except KeyError:
                eel.raise_error("В Excel документе нет поля email, сертификаты не были отправлены")
                send = False
            except Exception as e:
                print("\n\n")
                print("ИСКЛЮЧЕНИЕ!!!!!!!!!!!!!!!!!!!!", e)
                print("\n\n")
                STOP = True
                while STOP:
                    STOP = False
                    powerpoint.Quit()
                    smtps.quit()
                    time.sleep(5)
                    try:
                        powerpoint = init_powerpoint()
                        smtps = login()
                        send_email(loc['email'], smtps, loc['date'], file_name[indexOf(data, loc)])
                        print(file_name[indexOf(data, loc)], loc['email'], ' - отправлен' )
                    except Exception as e:
                        print("\n\n")
                        print("ИСКЛЮЧЕНИЕ!!!!!!!!!!!!!!!!!!!!", e)
                        print("\n\n")
                        STOP = True
                    if STOP == False:
                        break
    smtps.quit()
    eel.raise_error("Процесс успешно завершен")
    powerpoint.Quit()
    time2 = time.perf_counter()
    print(f"Finished in {time2-time1} second(s)")
    
    logger.info("Done!", "\n", f"Finished in {time2-time1} second(s)")

if __name__ == "__main__":
    eel.start("HomePage.html", geometry={"size": (600, 400), "position": (400, 600)}, port=8002)
