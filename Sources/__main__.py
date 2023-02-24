import agent_initializetion
from agent_initializetion import *


def active_directory_change(user: dict):
    add_group_ad(user)


def doc_change(user: dict):
    kill_browser()
    sleep(4)
    chrome = None
    success = False
    c = 0
    while True:
        try:
            chrome = ChromeOtbasy()
            chrome.doc_change_(data=user)
            success = True
        except Exception as error:
            if isinstance(error, RemoteDisconnected) or isinstance(error, ProtocolError):
                print("Remote end closed connection without response, trying to restart the WebDriver...")
            else:
                print("Error occurred:", error)
                break
        finally:
            sleep(3)
            # chrome.quit()
            if success:
                break


def bpm_change(user: dict):
    kill_browser()
    sleep(4)
    driver = None
    success = False
    while True:
        try:
            driver = ExplorerOtbasy()
            driver.bpm_change_(user)
            success = True
        except Exception as error:
            if isinstance(error, RemoteDisconnected) or isinstance(error, ProtocolError):
                print("Remote end closed connection without response, trying to restart the WebDriver...")
            else:
                print("Error occurred:", error)
                break
        finally:
            # driver.quit()
            if success:
                break


def colvir_change(user: dict):
    colvir = Colvir(user)
    colvir.kill()
    if 'colvir' not in user.keys():
        colvir.colvir_is_card_exist()
        colvir.kill()
        sleep(2)
        colvir.colvir_create_user()
        colvir.colvir_add_permissions()
    else:
        colvir.change_role()
    colvir.kill()


def main():
    kill_browser()
    sleep(4)
    chrome = None
    success = False
    data = None
    while True:
        try:
            chrome = ChromeOtbasy()
            data = chrome.bpm_parse()
            success = True
        except Exception as error:
            if isinstance(error, RemoteDisconnected) or isinstance(error, ProtocolError):
                print("Remote end closed connection without response, trying to restart the WebDriver...")
            else:
                print("Error occurred:", error)
                break
        finally:
            # chrome.quit()
            if success:
                break

    if data is None:
        print('Данные не собраны заявок')
        return

    for user in data:
        try:
            if 'roles' not in user.keys():
                continue
            if len(user['roles'].keys()) == 1 and 'СЭО Nomad' in user['roles'].keys():
                print('Only NOMAD')
                continue

            # ! Active Directory
            active_directory_change(user)

            # ! BPM
            bpm_change(user)

            # ! СЭД "Documentolog"
            if 'СЭД «Documentolog»' in user['roles'].keys():
                try:
                    doc_change(user)
                except Exception as ex:
                    print(f'Ошибка в создании Documentolog. Ошибка: {ex}')

            # ! Colvir
            if 'АБИС "Colvir"' in user['roles'].keys():
                colvir_change(user)

        except Exception as ex:
            print(f'Ошибка при создании пользователя {user["username"]}. Ошибка: {ex}')


if __name__ == '__main__':
    main()
