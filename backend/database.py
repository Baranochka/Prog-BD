import re
from abc import ABC, abstractmethod
from typing import Optional, List, Any
from pyodbc import connect, Row
from datetime import datetime

# from Temp import update


def debug(*args, **kwargs) -> None:
    print(*args, **kwargs)


class Database(ABC):
    @abstractmethod
    def get_all_rows(self) -> ...: ...

    @abstractmethod
    def read_row_by_id(self, id: int) -> Any: ...

    @abstractmethod
    def read_rows_by_id(self, ids: List[int]) -> ...: ...


class MSSQL(Database):
    @staticmethod
    def __debug(*args, **kwargs) -> None:
        print(*args, **kwargs)

    def __init__(
            self,
            driver: str,
            server_name: str,
            database_name: str,
            username: str,
            password: str,
    ) -> None:
        try:
            self._connection = connect(
                f"DRIVER={driver};"
                f"Server={server_name};"
                f"Database={database_name};"
                f"UID={username};"
                f"PWD={password};"
                "Encrypt=no;"
            )
            self._cursor = self._connection.cursor()
            if self._connection:
                self.cond_conn = True
                # self.__debug("Successfully connected")
        except Exception as e:
            self.cond_conn = False
            self.__debug(e)

    @property
    def is_connected(self) -> bool:
        return self.cond_conn

    def get_all_rows(self) -> Optional[List[Row]]:
        try:
            self._cursor.execute("SELECT * FROM persons")
            result = self._cursor.fetchall()

            if result is not None:
                return result
            return None
        except Exception as e:
            self.__debug(e)
            return None

    def read_row_by_id(self, id: int) -> Optional[Row]:
        try:
            self._cursor.execute("SELECT * FROM persons WHERE id=?", (id,))
            row = self._cursor.fetchone()
            if row is not None:
                return row
            return None
        except Exception as e:
            self.__debug(e)
            return None

    def get_person(self,
                   surname: Optional[str],
                   name: Optional[str],
                   och: Optional[str],
                   birthdate: Optional[datetime]) -> Optional[List[List[str]]]:
        # try:
        # Начинаем с базового запроса
        query = "SELECT fru, last_lat, name_rus, nla, och, " \
                "oche, ctz1, dob, sex, pob, " \
                "cob, pas_ser, pas_num, pds, pde, " \
                "visa_priz, vis_ser, vis_num, d_poluch, vis_end, " \
                "tel_nom, d_enter, date_okon, mcs, mcn, " \
                "k, dnd, dog_obsh, rf, rfd, " \
                "mot, ser, nmr, vis_krat, vis_id, " \
                "gos_nap, proz, kontrakt, kont_start, kont_end, " \
                "gos_start, gos_end, star, d_poluch, str_poluch," \
                "city_poluch, num_prig, kpp, med, mot, " \
                "uch_st_st, pr, fr, o_p, prikaz, " \
                "prik_start, o_s, d_naym, email, prim " \
                "FROM persons"

        # Список условий и параметров
        conditions = []
        parameters = []

        # Добавляем условия в зависимости от переданных аргументов
        if surname is not None:
            if self.check_rus_eng(surname):
                conditions.append("fru = ?")
                parameters.append(surname)
            else:
                conditions.append("last_lat = ?")
                parameters.append(surname)
        if name is not None:
            if self.check_rus_eng(name):
                conditions.append("name_rus = ?")
                parameters.append(name)
            else:
                conditions.append("nla = ?")
                parameters.append(name)
        if och is not None:
            if self.check_rus_eng(och):
                conditions.append("och = ?")
                parameters.append(och)
            else:
                conditions.append("oche = ?")
                parameters.append(och)
        if birthdate != datetime(1900, 1, 1):
            conditions.append("dob = ?")
            parameters.append(birthdate)

        # Если есть условия, добавляем их к запросу
        if conditions:
            query += f" WHERE {' AND '.join(conditions)}"

        # debug(query)

        # Выполняем запрос
        self._cursor.execute(query, parameters)

        all_rows = self._cursor.fetchall()
        new_all_rows = []

        for row in all_rows:
            list_row = []
            # print(row)
            for field in row:
                # print(field)
                if isinstance(field, datetime):
                        list_row.append(field)
                else:
                    list_row.append(str(field).strip() if field is not None else "")
            new_all_rows.append(list_row)
        return new_all_rows

        # except Exception as e:
        # self.__debug(e)
        # return None

    def get_person_for_check(self,
                   surname: Optional[str],
                   name: Optional[str],
                   och: Optional[str],
                   birthdate: Optional[datetime]) -> Optional[List[List[str]]]:
        # try:
        # Начинаем с базового запроса
        query = "SELECT fru, last_lat, name_rus, nla, dob, pas_ser, pas_num, pds " \
                "FROM persons"

        # Список условий и параметров
        conditions = []
        parameters = []

        # Добавляем условия в зависимости от переданных аргументов
        if surname is not None:
            if self.check_rus_eng(surname):
                conditions.append("fru = ?")
                parameters.append(surname)
            else:
                conditions.append("last_lat = ?")
                parameters.append(surname)
        if name is not None:
            if self.check_rus_eng(name):
                conditions.append("name_rus = ?")
                parameters.append(name)
            else:
                conditions.append("nla = ?")
                parameters.append(name)
        if och is not None:
            if self.check_rus_eng(och):
                conditions.append("och = ?")
                parameters.append(och)
            else:
                conditions.append("oche = ?")
                parameters.append(och)
        if birthdate != datetime(1900, 1, 1):
            conditions.append("dob = ?")
            parameters.append(birthdate)

        # Если есть условия, добавляем их к запросу
        if conditions:
            # query += " WHERE " + " OR ".join(conditions)
            query += f" WHERE {' AND '.join(conditions)}"

        # debug(query)

        # Выполняем запрос
        self._cursor.execute(query, parameters)
        all_rows = self._cursor.fetchall()
        new_all_rows = []

        for row in all_rows:
            list_row = []
            for i, field in enumerate(row):
                if isinstance(field, datetime) or i in [4, 7]:
                    if field is None:
                        list_row.append("          ")
                    else:
                        date = field.strftime("%d%m%Y")  # Форматируем
                        list_row.append(date)
                else:
                    list_row.append(str(field).strip()
                                    if field is not None else "")
            new_all_rows.append(list_row)
        return new_all_rows

    def get_all_rows_for_check(self) -> Optional[List[Row]]:
        try:
            self._cursor.execute("SELECT fru, last_lat, name_rus, nla, dob, pas_ser, pas_num, pds FROM persons")
            all_rows = self._cursor.fetchall()

            new_all_rows = []

            for row in all_rows:
                list_row = []
                for i, field in enumerate(row):
                    if isinstance(field, datetime) or i in [4, 7]:
                        if field is None:
                            list_row.append("          ")
                        else:
                            date = field.strftime("%d%m%Y")  # Форматируем
                            list_row.append(date)
                    else:
                        list_row.append(str(field).strip()
                                        if field is not None else "")
                new_all_rows.append(list_row)
            return new_all_rows
        except Exception as e:
            self.__debug(e)
            return None

    def update_person(self, model=None,  surname=None, name=None, och=None, birthdate=None):


        query = "UPDATE persons SET "

        parameters = []
        updates = []
        update_par = ['fru = ?', 'last_lat = ?', 'name_rus = ?', 'nla = ?', 'och = ?', 'oche = ?', 'ctz1 = ?', 'dob = ?', 'sex = ?', 'pob = ?',
                   'cob = ?', 'pas_ser = ?', 'pas_num = ?', 'pds = ?', 'pde = ?', 'visa_priz = ?', 'vis_ser = ?', 'vis_num = ?',
                    'vis_start = ?', 'vis_end = ?', 'tel_nom = ?', 'd_enter = ?', 'date_okon = ?', 'mcs = ?', 'mcn = ?', 'k = ?',
                   'dnd = ?', 'dog_obsh = ?', 'rf = ?', 'rfd = ?', 'mot = ?', 'ser = ?', 'nmr = ?', 'vis_krat = ?', 'vis_id = ?',
                   'gos_nap = ?', 'proz = ?', 'kontrakt = ?', 'kont_start = ?', 'kont_end = ?', 'gos_start = ?', 'gos_end = ?',
                   'star = ?', 'd_poluch = ?', 'str_poluch = ?', 'city_poluch = ?', 'num_prig = ?', 'kpp = ?', 'med = ?', 'mot = ?',
                   'uch_st_st = ?', 'pr = ?', 'fr = ?', 'o_p = ?', 'prikaz = ?', 'prik_start = ?', 'o_s = ?', 'd_naym = ?', 'email = ?',
                   'prim = ?' ]

        for i, var in enumerate(update_par):
            if model.data_update[i] != '':
                updates.append(var)
                parameters.append(model.data_update[i])

        # Формируем запрос
        query += ", ".join(updates) + " WHERE last_lat = ?"
        parameters.append(surname)

        # Выводим для отладки (если нужно)
        print(query)
        print(parameters)

        # Выполняем запрос
        self._cursor.execute(query, parameters)
        self._connection.commit()  # Фиксируем изменения

    def check_rus_eng(self, text) -> bool:
        return bool(re.search(r'[а-яА-ЯёЁ]', text))

    def read_rows_by_id(self, ids: int) -> List[Row]:
        ...

    def read_row_by_surname(self, surname: str) -> Optional[Row]:
        self._cursor.execute("SELECT * FROM persons WHERE fru=?", (surname,))
        row = self._cursor.fetchone()
        if row is not None:
            return row
        return None

    def read_row_by_name(self) -> Row:
        ...

    def read_row_by_patronymic(self) -> Row:
        ...


if __name__ == "__main__":
    from configparser import ConfigParser
    import os
    config = ConfigParser()
    path = os.path.dirname(os.path.abspath(__file__))
    print(path)
    config.read(f"{path}\\config.ini")

    DRIVER = config["Test_Maks"]["DRIVER"]
    SERVER_NAME = config["Test_Maks"]["SERVER_NAME"]
    DATABASE_NAME = config["Test_Maks"]["DATABASE_NAME"]
    USERNAME = config["Test_Maks"]["USERNAME"]
    PASSWORD = config["Test_Maks"]["PASSWORD"]

    db = MSSQL(DRIVER, SERVER_NAME, DATABASE_NAME, USERNAME, PASSWORD)

    all_rows = db.get_person("АМИРИ", None, None, datetime(1900, 1, 1))
    debug(all_rows)
    # col = ["fru", "last_lat", "name_rus", "nla", "och", "oche", "ctz1", "dob",
    #        "sex", "pob", "cob", "pas_ser", "pas_num", "pds", "pde", "visa_priz", "vis_sev", "vis_num",
    #        "d_poluch", "vis_endv", "vtel_nom", "d_enter", "date_okon", "mcs", "mcn", "k",
    #        "dnd", "dog_obsh", "rf", "rfd", "motv", "ser", "nmr", "vis_krat", "vis_id", "gos_nap",
    #        "kont_start", "kontrakt", "kont_start", "kont_end", "gos_start", "gos_end", "star"]
    # i = 0
    # # for var, c in zip(all_rows[0], col):
    # #     debug(f"{i}. {c}: {var}")
    # #     i += 1
    # for var in all_rows[0]:
    #     debug(f'{i}: "{var}"')
    #     i += 1
    # all_rows = db.get_person_for_check("ХРАЙЗАТ", None, None, datetime(1900, 1, 1))
    # debug(all_rows)

    # db.update_person(surname="AHMED")
