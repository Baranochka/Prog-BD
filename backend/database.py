import sys
from abc import ABC, abstractmethod
from typing import Optional, List, Any
from pyodbc import connect, Row
from datetime import datetime


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
                self.__debug("Successfully connected")
        except Exception as e:
            self.cond_conn = False
            self.__debug(e)
            

    @property
    def is_connected(self) -> bool:
        return self.cond_conn
    
    def get_all_rows(self) -> Optional[list[Row]]:
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
    # def get_person(self,
    #                last_name: Optional[str],
    #                first_name: Optional[str],
    #                patronymic: Optional[str],
    #                birthday: Optional[datetime]) -> Optional[List[List[str]]]:
    #     try:
    #         self._cursor.execute(
    #             "SELECT fru, last_lat, name_rus, nla, och, oche, pob ,dob, "
    #             "sex, pob, cob, pas_ser, pas_num, pds, pde, vis_ser, vis_num, vis_start"
    #             ", vis_end, tel_nom, d_enter, date_okon, mcs, mcn, obsch, d_naym, dog_obsh "
    #             "FROM persons WHERE fru=? OR name_rus=? OR och=? OR dob=?",
    #             (last_name, first_name, patronymic, birthday))
    #         self.__debug(last_name, first_name, patronymic, birthday)
    #         all_rows = self._cursor.fetchall()
    #         new_all_rows = []
    #         list_row = []
    #
    #         for row in all_rows:
    #             self.__debug(row)
    #             list_row = []
    #             for field in row:
    #                 if isinstance(field, datetime):
    #                     # Преобразуем datetime в строку в нужном формате
    #                     list_row.append(field.strftime("%d"))  # Пример формата
    #                     list_row.append(field.strftime("%m"))
    #                     list_row.append(field.strftime("%Y"))
    #                 else:
    #                     list_row.append(str(field).strip() if field is not None else "")
    #             list_row[-2] = list_row[-2][-2:] # Замена 4-х значной даты с конца на 2 знака.
    #             if len(list_row[29]) == 11:
    #                 list_row[29] = list_row[29][1:]
    #             new_all_rows.append(list_row)
    #
    #         return new_all_rows
    #
    #     except Exception as e:
    #         self.__debug(e)
    #         return None

    def get_person(self,
                   surname: Optional[str],
                   name: Optional[str],
                   patronymic: Optional[str],
                   birthdate: Optional[datetime]) -> Optional[List[List[str]]]:
        # try:
            # Начинаем с базового запроса
            query = "SELECT fru, last_lat, name_rus, nla, och, oche, ctz1, dob, " \
                    "sex, pob, cob, pas_ser, pas_num, pds, pde, visa_priz, vis_ser, vis_num, " \
                    "d_poluch, vis_end, tel_nom, d_enter, date_okon, mcs, mcn, k, " \
                    "dnd, dog_obsh, rf, rfd, mot, ser, nmr, vis_krat, vis_id, gos_nap, " \
                    "kont_start, kontrakt, kont_start, kont_end " \
                    "FROM persons"

            # Список условий и параметров
            conditions = []
            parameters = []

            # Добавляем условия в зависимости от переданных аргументов
            if surname is not None:
                conditions.append("fru = ?")
                parameters.append(surname)
            if name is not None:
                conditions.append("name_rus = ?")
                parameters.append(name)
            if patronymic is not None:
                conditions.append("och = ?")
                parameters.append(patronymic)
            if birthdate != datetime(1900, 1, 1):
                conditions.append("dob = ?")
                parameters.append(birthdate)

            # Если есть условия, добавляем их к запросу
            if conditions:
                # query += " WHERE " + " OR ".join(conditions)
                query += f" WHERE {' AND '.join(conditions)}"

            debug(query)

            # Выполняем запрос
            self._cursor.execute(query, parameters)
            self.__debug(f"DEBUG args = '{surname}' '{name}' '{patronymic}' '{birthdate}'")
            all_rows = self._cursor.fetchall()
            new_all_rows = []

            for row in all_rows:
                list_row = []
                for i, field in enumerate(row):
                    #print(field,i)
                    if isinstance(field, datetime) or i in [7, 13, 14, 18, 19, 21, 22, 29, 30, 36, 38, 39]:
                        if field is None:
                            list_row.append("          ")
                        else:
                            date =  field.strftime("%d.%m.%Y")  # Форматируем
                            list_row.append(date)
                    else:
                        list_row.append(str(field).strip() if field is not None else "")
                if len(list_row[20]) == 11: # проверка номера на 11 символов
                    list_row[20] = list_row[20][1:]
                new_all_rows.append(list_row)
            return new_all_rows

        # except Exception as e:
            # self.__debug(e)
            # return None

    def read_rows_by_id(self, ids: int) -> list[Row]:
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
    config.read(f"{path}\\config.ini")

    DRIVER        = config["Test_Maks"]["DRIVER"]  
    SERVER_NAME   = config["Test_Maks"]["SERVER_NAME"]  
    DATABASE_NAME = config["Test_Maks"]["DATABASE_NAME"] 
    USERNAME      = config["Test_Maks"]["USERNAME"]
    PASSWORD      = config["Test_Maks"]["PASSWORD"]

    db = MSSQL(DRIVER, SERVER_NAME, DATABASE_NAME, USERNAME, PASSWORD)

    surname = "ФАМ"
    birthdate = datetime(1990, 1, 1)

    # is_find = db.is_person_exists(surname, None, None, birthdate)
    # # debug(is_find)
    # if is_find:
    #     all_rows = db.get_person(surname, None, None, birthdate)
    #
    #     for row in all_rows:
    #         debug(row)
    
    all_rows = db.get_person("ХАБИБИ", None, None, datetime(1900, 1, 1))
    debug(all_rows)
    debug(len(all_rows[0]))
