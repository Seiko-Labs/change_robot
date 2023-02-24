import datetime
import io
from time import sleep
from typing import Optional
import psutil
from selenium.webdriver.support.select import Select
import selenium.webdriver.ie.options
from pyPythonRPA.Robot import pythonRPA
import pyautogui
from selenium.webdriver.common import keys
from selenium.webdriver.common.by import By
from transliterate import translit
from webdriver_manager.microsoft import IEDriverManager
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver import Ie, Chrome, ActionChains, Keys
from selenium.webdriver.chrome.service import Service as CService
from selenium.webdriver.ie.service import Service
from selenium.common.exceptions import NoSuchElementException, SessionNotCreatedException
from active_directory import *
from urllib3.exceptions import ProtocolError


def kill_browser():
    processes = ('iexplore.exe', 'IEDriverServer.exe', 'EXCEL.EXE')
    try:
        for proc in psutil.process_iter():
            if proc.name() in processes:
                proc.kill()
    except Exception as e:
        print('Error on closing Internet Explorer: ' + str(e))


class ChromeOtbasy(Chrome):
    def __init__(self, options=None):
        if options:
            self.options = options
        else:
            self.options = selenium.webdriver.ie.options.Options()
            self.options.add_argument("--enable-smooth-scrolling")
            self.options.add_argument("--start-maximized")
            self.options.add_argument("--no-sandbox")
            self.options.add_argument("--disable-dev-shm-usage")
            self.options.add_argument("--disable-extensions")
            self.options.add_argument("--disable-notifications")
            self.options.add_argument("--ignore-ssl-errors=yes")
            self.options.add_argument("--ignore-certificate-errors")
            self.options.add_argument('--allow-insecure-localhost')
        super().__init__(service=CService(ChromeDriverManager().install()), options=options)
        self.implicitly_wait(15)
        self.maximize_window()

    def doc_change_(self, data: dict):
        # * Authorize
        self.get('https://doc.hcsbk.kz/structure/index')
        self.find_element(By.XPATH, "//input[@id='login']").send_keys(login)
        self.find_element(By.XPATH, "//input[@id='password']").send_keys(password)
        self.find_element(By.XPATH, "//input[@id='submit']").click()
        sleep(1)
        pythonRPA.keyboard.press('ENTER')
        if len(self.find_elements(By.XPATH, "//input[@value='Все равно войти']")) == 1:
            sleep(1)
            pythonRPA.keyboard.press('ENTER')

        # data = replace_filial(data, doc_map)

        if 'ДДО' in data['department'] and 'ОПЕРАТОР' in str(data['department']).upper():
            data['department'] = 'Департамент дистанционного обслуживания'

        # * Choose department
        if data['filial_type'] == 'Пользователи ЦА':
            self.find_element(By.XPATH, '//input[@id="structure_search_input"]').send_keys(data['branch'])
            els = self.find_elements(By.XPATH, '//li[@class="ui-menu-item"]')
            els[-1].click()
            sleep(1)
            # selector = self.find_element(By.XPATH, f'//span[contains(.,"{data["branch"]}")]/parent::*/following::*')
            selector = self.find_element(By.XPATH, f'//p[@class="selected"]/span[@class="buttons"]/span')
            selector.click()
            sleep(3)
            selector1 = self.find_element(By.XPATH,
                                          '//span[@class="drop_down_list"][@style=""]/a[contains(.,"сотрудника")]')
            selector1.click()
            sleep(3)
            menu_items = self.find_elements(By.XPATH, "//ul[@id='tabtitles']/li")
        elif data['department'] in department_exceptions:
            sleep(1)
            self.find_element(By.XPATH, '//input[@id="structure_search_input"]').send_keys(data['department'])
            els = self.find_elements(By.XPATH, f'//li[@class="ui-menu-item"]/a[text()=\'{data["department"]}\']')
            els[-1].click()
            sleep(1)
            selector = self.find_element(By.XPATH, f'//p[@class="selected"]/span[@class="buttons"]/span')
            selector.click()
            sleep(3)
            selector1 = self.find_element(By.XPATH,
                                          '//span[@class="drop_down_list"][@style=""]/a[contains(.,"сотрудника")]')
            selector1.click()
            sleep(3)
            menu_items = self.find_elements(By.XPATH, "//ul[@id='tabtitles']/li")
        elif 'ТРУДОВЫЕ' in str(data['department']).upper() or 'ТРУДОВОЕ' in str(data['department']).upper():
            sleep(1)
            self.find_element(By.XPATH, '//input[@id="structure_search_input"]').send_keys(data['filial_doc'])
            els = self.find_elements(By.XPATH, f'//li[@class="ui-menu-item"]/a[contains(.,\'{data["filial_doc"]}\')]')
            els[-1].click()
            sleep(1)
            selector = self.find_element(By.XPATH, f'//p[@class="selected"]/span[@class="buttons"]/span')
            selector.click()
            sleep(3)
            selector1 = selector.find_element(By.XPATH,
                                              '//span[@class="drop_down_list"][@style=""]/a[contains(.,"сотрудника")]')
            selector1.click()
            sleep(3)
            menu_items = self.find_elements(By.XPATH, "//ul[@id='tabtitles']/li")
        elif data['filial_type'] == 'Пользователи филиалов':
            # data = replace_filial(data, doc_map)
            self.find_element(By.XPATH, '//input[@id="structure_search_input"]').send_keys(data['department'])
            els = self.find_elements(By.XPATH, f'//li[@class="ui-menu-item"][contains(.,"{data["filial_doc"]}")]')
            els[-1].click()
            sleep(1)
            selector = self.find_element(By.XPATH, f'//p[@class="selected"]/span[@class="buttons"]/span')
            selector.click()
            sleep(3)
            selector = self.find_element(By.XPATH,
                                         '//span[@class="drop_down_list"][@style=""]/a[contains(.,"сотрудника")]')
            selector.click()
            sleep(3)
            menu_items = self.find_elements(By.XPATH, "//ul[@id='tabtitles']/li")

        elif data['department'] in department_exceptions:
            sleep(1)
            self.find_element(By.XPATH, '//input[@id="structure_search_input"]').send_keys(data['department'])
            els = self.find_elements(By.XPATH, f'//li[@class="ui-menu-item"]/a[text()=\'{data["filial_map"]}\']')
            els[-1].click()
            sleep(1)
            selector = self.find_element(By.XPATH, f'//p[@class="selected"]/span[@class="buttons"]/span')
            selector.click()
            sleep(3)
            selector1 = self.find_element(By.XPATH,
                                          '//span[@class="drop_down_list"][@style=""]/a[contains(.,"сотрудника")]')
            selector1.click()
            sleep(3)
            menu_items = self.find_elements(By.XPATH, "//ul[@id='tabtitles']/li")
        else:
            raise ValueError(f'"{data["filial_type"]}" не существует')

        # ! ~~~~~~~~~~~~~~~~~~~~~~~ Fill employee card ~~~~~~~~~~~~~~~~~~~~~~~~~~

        # * Menu "Должность"
        self.execute_script(menu_items[0].get_attribute('onclick'))
        sleep(1)
        self.find_element(By.ID, 'edit_field_display_name').send_keys(data['title'])
        self.find_element(By.ID, 'edit_field_work_phone').send_keys('-')
        self.find_element(By.ID, 'edit_field_f_email').send_keys(data['mail'])

        # * Menu "Сотрудник"
        self.execute_script(menu_items[1].get_attribute('onclick'))
        sleep(1)
        self.find_element(By.ID, 'edit_field_f_surname').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_f_name').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_f_middle_name').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_employee_name').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_employee_name_kz').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_f_display_short_name').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_f_display_short_name_kz').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_dative_name').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_dative_name_kz').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_instrumental_display_name').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_instrumental_display_name_kz').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_genetive_name').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_genetive_name_kz').send_keys(data['username'])
        sleep(1)

        # * Меню "Пользователь"
        self.execute_script(menu_items[2].get_attribute('onclick'))
        sleep(1)
        self.find_element(By.ID, 'edit_field_login').send_keys(data['login'])
        self.find_element(By.ID, 'edit_field_email').send_keys(data['mail'])
        sleep(1)
        self.find_element(By.ID, 'is_active_true').click()

        # * Save
        self.find_element(By.XPATH, '/html/body/div[4]/div[3]/div/button[1]').click()
        sleep(3)
        sleep(2)

    def bpm_parse(self) -> list:
        self.get('http://bpm/site/Matrix/UserMessages.aspx')
        sleep(4)
        pythonRPA.keyboard.write('robot.ad', timing_after=1)
        pythonRPA.keyboard.press('TAB')
        pythonRPA.keyboard.write('Asd12345678', timing_after=1)
        pythonRPA.keyboard.press('TAB', timing_after=1)
        pythonRPA.keyboard.press('ENTER', timing_after=2)
        main_window = self.current_window_handle
        sleep(5)

        # * Первичный сбор заявок
        users = self.find_elements(By.XPATH, '//table[@id="GV_UserMessages"]/tbody/tr[not(@align)]')
        data = []
        for i in users:
            user = i.find_elements(By.TAG_NAME, 'td')
            data.append({'username': user[5].text, 'department': user[4].text,
                         # 'link': user[0].find_element(By.TAG_NAME, 'a')
                         })

        # * Удаление заявок на создание учётной записи
        to_remove = []
        for user in data:
            if ad_find_user(connect_to_ad(), user['username']) is None:
                to_remove.append(user)

        for user in to_remove:
            data.remove(user)

        # * Сбор данных в словарь
        for i in data:
            sleep(5)
            try:
                self.find_element(By.XPATH, f'//a[text()="{i["username"]}"]/parent::td/parent::tr/td/a').send_keys(
                    keys.Keys.ENTER)
            except NoSuchElementException:
                print(f"{i['username']} - не собран")
                continue
            sleep(1)
            self.switch_to.window(self.window_handles[-1])
            i['title'] = self.find_element(By.ID, 'lbl_emp_pos').text  # Должность
            i['ID'] = self.find_element(By.ID, 'LBL_TABID').text  # Табельный номер
            i['manager'] = self.find_element(By.XPATH,
                                             '//label[text()="Руководитель:"]/following::label').text  # Руководитель
            i['description'] = self.find_element(By.ID, 'lbl_employee_name').text  # Регистрационный номер
            i['filial'] = self.find_element(By.XPATH, '//label[text()="Банк:"]/following::label').text  # Филиал
            i['department'] = self.find_element(By.ID, 'lbl_emp_departamentn').text  # Подразделение

            # Сбор ролей в словарь
            self.implicitly_wait(2)
            role_table = self.find_elements(By.XPATH, '//tr[contains(@id, "maintr")][not(@style="display: none;")]')
            self.implicitly_wait(15)
            if len(role_table) < 1:
                print('Нет прав для предоставления')
                self.close()
                sleep(1)
                for window in self.window_handles:
                    self.switch_to.window(window)
                    if 'Система' in self.title:
                        break
                continue
            roles = {}
            for role in role_table:
                self.implicitly_wait(1)
                current_line = role.find_elements(By.TAG_NAME, 'td')
                if len(current_line[5].find_elements(By.XPATH, 'div/input[not(@checked)]')) == 1:
                    continue
                self.implicitly_wait(15)
                if current_line[0].text in roles.keys():
                    roles[current_line[0].text].append(current_line[1].text)
                else:
                    roles[current_line[0].text] = [current_line[1].text, ]
            i['roles'] = roles  # Роли
            i = replace_code(i)

            # Управление (подразделение)
            self.implicitly_wait(1)
            if len(self.find_elements(By.ID, 'lbl_emp_depn')) == 1:
                i['branch'] = self.find_element(By.ID, 'lbl_emp_depn').text
            self.implicitly_wait(15)

            # Доступ в интернет
            if 'Интернет' not in i['roles'].keys():
                i['roles']['Интернет'] = ['INET_MIN_ACCESS']
            for inet in internet_hierarchy:
                if inet in i['roles']['Интернет']:
                    i['roles']['Интернет'] = [inet]

            # Проверка логина на наличие специальных символов
            i['temp_name'] = i['username']
            for j in i['username']:
                if j in alph_map.keys():
                    i['temp_name'] = i['temp_name'].replace(j, alph_map[j])
            text = translit(i['temp_name'], 'ru', reversed=True).split()
            if is_spec_alphabet(i['username']):
                i['login'] = f"{text[0]}.{text[1][0:2]}"
                i['login'] = i['login'].lower()
                i['user'] = f"{text[0]}_{text[1][0:2]}"
                i['user'] = i['user'].upper()
            else:
                i['login'] = f"{text[0]}.{text[1][0]}"
                i['login'] = i['login'].lower()
                i['user'] = f"{text[0]}_{text[1][0]}"
                i['user'] = i['user'].upper()

            if i['filial'] != 'АО "Жилищный строительный сберегательный банк "Отбасы банк"':
                i['login'] = f"{i['filial_c']}.{i['login']}"
                i['user'] = f"{i['filial_c']}_{i['user']}"

            i['login'] = str(i['login']).replace('\'', '').lower()
            i['user'] = str(i['user']).replace('\'', '').upper()

            i['mail'] = f'{i["login"]}@otbasybank.kz'  # Почта
            i = replace_filial(i)
            print(i)
            self.implicitly_wait(1)
            if len(self.find_elements(By.XPATH, '//table[@class="table table"]')) == 1:
                replaced_roles_table = self.find_elements(By.XPATH,
                                                          '//table[@class="table table"]/tbody/tr[not(@style)]/td[5]/div/input[@checked]/parent::div/parent::td/parent::tr')
                roles = {}
                for role in replaced_roles_table:
                    self.implicitly_wait(1)
                    current_line = role.find_elements(By.TAG_NAME, 'td')
                    if len(current_line[5].find_elements(By.XPATH, 'div/input[not(@checked)]')) == 1:
                        continue
                    self.implicitly_wait(15)
                    current_line = role.find_elements(By.TAG_NAME, 'td')
                    if current_line[0].text in roles.keys():
                        roles[current_line[0].text].append(current_line[1].text)
                    else:
                        roles[current_line[0].text] = [current_line[1].text, ]
                i['old_roles'] = i['roles']
                i['roles'] = roles  # Роли

                # Доступ в интернет
                if 'Интернет' not in i['roles'].keys():
                    i['roles']['Интернет'] = ['INET_MIN_ACCESS']
                for inet in internet_hierarchy:
                    if inet in i['roles']['Интернет']:
                        i['roles']['Интернет'] = [inet]

            self.implicitly_wait(15)
            self.close()
            sleep(1)
            # self.switch_to.window(main_window)
            for window in self.window_handles:
                self.switch_to.window(window)
                if 'Система' in self.title:
                    break

        return data


class ExplorerOtbasy(Ie):
    def __init__(self, options=None):
        if options:
            self.options = options
        else:
            self.options = selenium.webdriver.ie.options.Options()
            self.options.add_argument("--enable-smooth-scrolling")
            self.options.add_argument("--start-maximized")
            self.options.add_argument("--no-sandbox")
            self.options.add_argument("--disable-dev-shm-usage")
            self.options.add_argument("--disable-print-preview")
            self.options.add_argument("--disable-extensions")
            self.options.add_argument("--disable-notifications")
            self.options.add_argument("--ignore-ssl-errors=yes")
            self.options.add_argument("--ignore-certificate-errors")
            self.options.add_argument('--allow-insecure-localhost')
            self.options.ignore_zoom_level = True
            self.options.ignore_protected_mode_settings = True
        super().__init__(service=Service(IEDriverManager().install()), options=self.options)
        self.maximize_window()
        sleep(1)
        self.execute_script("document.body.style.zoom='100%'")
        self.implicitly_wait(15)
        self.proc = 'IEXPLORE.EXE'

    def bpm_change_(self, user: dict):
        self.get('http://bpm/site/Administration/Users.aspx')
        sleep(3)
        pythonRPA.keyboard.write('robot.ad', timing_after=1)
        pythonRPA.keyboard.press('TAB')
        pythonRPA.keyboard.write('Asd12345678', timing_after=1)
        pythonRPA.keyboard.press('TAB', timing_after=1)
        pythonRPA.keyboard.press('ENTER', timing_after=2)
        main_window = self.current_window_handle

        # Поиск сотрудника
        self.find_element(By.ID, 'TB_Search_FullName').send_keys(user['username'])  # Ввод ФИО
        self.find_element(By.ID, 'Button_Search').send_keys(keys.Keys.ENTER)  # Поиск
        sleep(2)
        # self.find_element(By.XPATH, '//td[text()="Работающий"]/parent::tr/td/a').click()
        self.implicitly_wait(1)
        if len(self.find_elements(By.ID, 'GV_IpotekaUsers')) < 1:
            short_name = user['username'].split(' ')
            short_name = f"{short_name[0]} {short_name[1]}"
            self.find_element(By.ID, 'TB_Search_FullName').clear()
            self.find_element(By.ID, 'TB_Search_FullName').send_keys(short_name)  # Ввод ФИО
            self.find_element(By.ID, 'Button_Search').send_keys(keys.Keys.ENTER)  # Поиск
        self.implicitly_wait(15)
        if self.find_element(By.XPATH, '//tr[not(@class)]/td[11]').text != ' ':
            user['colvir'] = True
            user['user'] = self.find_element(By.XPATH, '//tr[not(@class)]/td[11]').text
        if 'ПО «BPM»' not in user['roles'].keys():
            return
        self.find_element(By.XPATH, '//a[text()="Выбрать"]').send_keys(keys.Keys.ENTER)
        sleep(1)
        self.find_element(By.ID, 'IB_Update').send_keys(keys.Keys.ENTER)

        # Открытие справочника
        df = pd.read_excel('glossary.xlsx', sheet_name='BPM')
        codes = df['Code'].tolist()
        roles = df['Role'].tolist()
        groups = df['Group'].tolist()
        codes_roles_dict = {k: v for k, v in zip(codes, roles)}

        # Удаление старых ролей
        untouchable_roles = [
            '5-70-22 - Доступ на запрос в ПКБ',
            '5-70-19 - Доступ на запрос в ГЦВП',
            '5-70-60 - Доступ к карточке клиента'
        ]
        current_roles = self.find_elements(By.XPATH, '//select[@id="LB_UserRoles"]/option')
        for role, urole in zip(current_roles, untouchable_roles):
            if role.text in urole:
                user['roles']['ПО «BPM»'].append(urole)

        for _ in current_roles:
            select = Select(webelement=self.find_element(By.ID, 'LB_UserRoles'))
            # select.select_by_visible_text(sel.text)
            select.select_by_index(0)
            sleep(1)
            self.find_element(By.XPATH, '//input[@id="Button_DeleteRole"]').send_keys(keys.Keys.ENTER)
            sleep(1)

        # Добавление групп
        # for i in user['roles']['ПО «BPM»']:
        #     code = i.split(' - ')[0]
        #     group = None
        #     for j, k in zip(code, groups):
        #         if code == j:
        #             group = k
        #     self.find_element(By.ID, 'Button_AddExecutor').send_keys(keys.Keys.ENTER)
        #     sleep(2)
        #     for handle in self.window_handles:
        #         self.switch_to.window(handle)
        #         if 'ролей' in self.title:
        #             break
        #     try:
        #         selector = self.find_element(By.XPATH, f'//td[text()="{group}"]/preceding-sibling::td/a')
        #         selector.send_keys(keys.Keys.ENTER)
        #         self.find_element(By.ID, 'IB_Select').send_keys(keys.Keys.ENTER)
        #     except NoSuchElementException:
        #         self.find_element(By.ID, 'IB_CloseWindow').send_keys(keys.Keys.ENTER)
        #     sleep(2)
        #     for handle in self.window_handles:
        #         self.switch_to.window(handle)
        #         if 'ролей' not in self.title:
        #             break

        # Добавление ролей

        # current_roles = self.find_elements(By.XPATH, '//select[@id="LB_UserRoles"]/option')
        # current_roles = [el.text for el in current_roles]
        user['roles']['ПО «BPM»'] = set(user['roles']['ПО «BPM»'])
        if '5-70-12 - Кредитный администратор' in user['roles']['ПО «BPM»'] and '5-70-1 - Кредитный эксперт (роль - Кредитный эксперт)' in user['roles']['ПО «BPM»']:
            user['roles']['ПО «BPM»'].remove('5-70-1 - Кредитный эксперт (роль - Кредитный эксперт)')

        for i in user['roles']['ПО «BPM»']:
            code = i.split(' - ')[0]
            if code in codes_roles_dict:
                role = codes_roles_dict[code]
            else:
                print(f'Роль с кодом {code} не найдена')
                continue
            if 'Нет роли в BPM' in role:
                continue
            self.find_element(By.ID, 'Button_AddRole').send_keys(keys.Keys.ENTER)
            sleep(2)
            for handle in self.window_handles:
                self.switch_to.window(handle)
                if 'ролей' in self.title:
                    break
            selector = self.find_element(By.XPATH, f'//td[text()="{role}"]/preceding-sibling::td/a')
            selector.send_keys(keys.Keys.ENTER)
            self.find_element(By.ID, 'IB_Select').send_keys(keys.Keys.ENTER)
            sleep(2)
            for handle in self.window_handles:
                self.switch_to.window(handle)
                if 'ролей' not in self.title:
                    break
        sleep(2)
        self.find_element(By.ID, 'Button_OK').send_keys(keys.Keys.ENTER)
        sleep(5)
