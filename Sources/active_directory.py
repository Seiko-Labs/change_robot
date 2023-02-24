import pandas as pd
from ldap3 import Server, Connection, SIMPLE, SYNC, ALL, MODIFY_REPLACE
from openpyxl import load_workbook
from transliterate import translit
from config import *


def connect_to_ad() -> Connection:
    s = Server(
        'HCSBKKZ.LOC',
        port=636,
        use_ssl=True,
        get_info=ALL)
    c = Connection(
        s,
        auto_bind=True,
        client_strategy=SYNC,
        user='robot.ad@hcsbkkz.loc',
        password='Asd12345678',
        authentication=SIMPLE,
        check_names=True)
    return c


def ad_find_user(c: Connection, username: str):
    c.search(f'CN={username},OU=Block,DC=hcsbkkz,DC=loc', '(objectclass=person)', attributes='*')
    if c.entries:
        print(f"{username} - учётная запись заблокирована")
        return None

    filials = ['CA', 'Aktau', 'Aktobe', 'Almaty', 'Astana', 'Atyrau', 'Karaganda', 'Kokshetau', 'Kostanay',
               'Kyzylorda', 'Pavlodar', 'Petropavlovsk', 'Semey', 'Shymkent', 'Taldykorgan', 'Taraz', 'Turkestan',
               'Uralsk', 'Ust-Kamenogorsk']

    for filial in filials:
        dn = f'CN={username},OU=Users,OU={filial},DC=hcsbkkz,DC=loc'
        c.search(dn, '(objectclass=person)', attributes='*')
        if c.entries:
            return dn
    print(f"{username}. Учётная запись не существует")
    return None


def add_group_ad(user: dict):
    c = connect_to_ad()
    dn = ad_find_user(c, user['username'])
    if dn is None:
        print('Пользователь не найден в системе Active Directory. Права не добавлены')
        return

    if c.entries[0]['mailNickname'] != user['login']:
        user['login'] = str(c.entries[0]['mailNickname'])
        user['mail'] = str(c.entries[0]['mail'])
        user['user'] = str(c.entries[0]['mailNickname']).replace('.', '_').upper()

    # Открыть справочник в Excel
    df = pd.read_excel('glossary.xlsx', sheet_name='AD')
    roles = df['Role Names']
    paths = df['Active Directory Path']
    to_append = []

    for role_type in user['roles'].values():
        for role in role_type:
            role_num = role.split(' - ')[0]
            if role_num == '5-70-133':
                to_append.append('CN=development_segment,OU=LAN_segments,OU=Bank-group,DC=hcsbkkz,DC=loc')
            if str(role_num[0]).isnumeric():
                for excel_role, path in zip(roles, paths):
                    if role_num == excel_role:
                        to_append.append(path)

    if 'Интернет' in user['roles'].keys():
        internet_groups = [
            'CN=PALO_ALTO_TOP2,CN=Users,DC=hcsbkkz,DC=loc',
            'CN=INET_WITH_SOCIAL,CN=Users,DC=hcsbkkz,DC=loc',
            'CN=PALO_ALTO_WITHOUT_SOCIAL,CN=Users,DC=hcsbkkz,DC=loc',
            'CN=INET_STANDART,CN=Users,DC=hcsbkkz,DC=loc',
            'CN=INET_MIN_ACCESS,CN=Users,DC=hcsbkkz,DC=loc',
            'CN=INET_JUST_HCSBK,CN=Users,DC=hcsbkkz,DC=loc',
        ]
        c.extend.microsoft.remove_members_from_groups([dn], internet_groups)

    c.extend.microsoft.add_members_to_groups([dn], to_append)


def insert_to_ad(data):
    s = Server(
        'HCSBKKZ.LOC',
        port=636,
        use_ssl=True,
        get_info=ALL)
    c = Connection(
        s,
        auto_bind=True,
        client_strategy=SYNC,
        user='robot.ad@hcsbkkz.loc',
        password='Asd12345678',
        authentication=SIMPLE,
        check_names=True)

    default_password = 'Asd12345678+'

    f"CN=Группа распространения,OU=Exchange_group,OU=Bank-Group,DC=hcsbkkz,DC=loc"
    data['memberOf'] = [f"CN=Группа распространения,OU=Exchange_group,OU=Bank-Group,DC=hcsbkkz,DC=loc",
                        f"CN=Пользователи Банка,OU=Exchange_group,OU=Bank-Group,DC=hcsbkkz,DC=loc",
                        f"CN={data['filial_type']},OU=Exchange_group,OU=Bank-Group,DC=hcsbkkz,DC=loc",
                        f"CN={str(data['roles']['Интернет']).lower()},CN=Users,DC=hcsbkkz,DC=loc"]
    DN = f"CN={data['username']},OU=Users,OU={data['filial_ad']},DC=hcsbkkz,DC=loc"
    OBJECT_CLASS = ['top', 'person', 'organizationalPerson', 'user']

    attributes = {
        "displayName": data['username'],
        "sAMAccountName": data['login'],
        "userPrincipalName": f"{data['login']}@hcsbkkz.loc",
        "name": data['username'],
        'title': data['title'],
        'department': data['department'],
        'manager': f"CN={data['manager']},OU=Users,OU={data['filial_ad']},DC=hcsbkkz,DC=loc",
        "givenName": data['username'].split()[0],
        "sn": data['username'].split()[1],
        "company": 'АО "Отбасы банк"',
        'employeeID': data['ID'],
        'description': data['description']
    }
    if not c.add(dn=DN, object_class=OBJECT_CLASS, attributes=attributes):
        if c.result.get('description') == 'constraintViolation':
            change_login(data)
            attributes['sAMAccountName'] = data['login']
            attributes['userPrincipalName'] = f"{data['login']}@hcsbkkz.loc"
        elif c.result.get('description') == 'entryAlreadyExists':
            print('Creation error. Account may be created')
            # data['username'] += '1'
            # attributes['displayName'] = data['username']
            DN = f"CN={data['username']},OU=Users,OU={data['filial_ad']},DC=hcsbkkz,DC=loc"
        else:
            print(f"Ошибка: {c.result.get('description')}")
    c.add(dn=DN, object_class=OBJECT_CLASS, attributes=attributes)
    c.extend.microsoft.modify_password(DN, default_password)
    password_expire = {"pwdLastSet": (MODIFY_REPLACE, [0]), 'userAccountControl': (MODIFY_REPLACE, [512])}
    c.modify(dn=DN, changes=password_expire)
    c.extend.microsoft.unlock_account(DN)
    c.extend.microsoft.add_members_to_groups([DN], data['memberOf'])
    return data


if __name__ == '__main__':
    ...
    # insert_to_ad(data)
