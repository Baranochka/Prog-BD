from datetime import datetime
import re as regexp
from pathlib import Path
from socket import gethostname

from sqlalchemy import create_engine, text
from sqlalchemy.engine import URL, Row
from sqlalchemy.orm import Session

from typing import Optional, Sequence, Tuple, Any, Union, List, Dict

from configparser import ConfigParser

from dataclasses import dataclass

HOSTNAME = gethostname()


def debug(*args, **kwargs) -> None:
    print(*args, **kwargs)


def get_connection_string(
    driver: str,
    database_name: str,
    username: str,
    password: str,
) -> str:
    """Функция для получения строки соединения к БД."""
    return (
        f"DRIVER={driver};"
        f"SERVER={HOSTNAME}\\SQLEXPRESS;"
        f"PORT=1433;"
        f"DATABASE={database_name};"
        f"UID={username};"
        f"PWD={password};"
        "Encrypt=no"
    )


def create_database_config(section: str, config: ConfigParser):
    """Функция для создания конфигурации БД."""
    driver = config.get(section, "DRIVER")
    database_name = config.get(section, "DATABASE_NAME")
    username = config.get(section, "USERNAME")
    password = config.get(section, "PASSWORD")

    connection_string = get_connection_string(
        driver,
        database_name,
        username,
        password,
    )
    connection_url = URL.create(
        "mssql+pyodbc",
        query={
            "odbc_connect": connection_string,
            "TrustServerCertificate": "yes",
        },
    )
    return connection_url


@dataclass(frozen=True)
class BaseDatabaseConfig:
    """Базовый класс конфигурации БД."""

    section: str
    connection_url: URL

    @classmethod
    def create(cls, section: str, config: ConfigParser):
        """Создает экземпляр класса с нужной конфигурацией."""
        connection_url = create_database_config(section, config)
        return cls(section=section, connection_url=connection_url)


@dataclass(frozen=True)
class DevelopmentDatabaseConfig(BaseDatabaseConfig):
    """Класс-конфигурация БД для разработки."""

    CONFIG_PATH = Path(__file__).resolve().parent / "config.ini"
    config = ConfigParser()
    config.read(CONFIG_PATH)

    section = "TEST"
    connection_url: URL = BaseDatabaseConfig.create(section, config).connection_url


@dataclass(frozen=True)
class ProductionDatabaseConfig(BaseDatabaseConfig):
    """Класс-конфигурация БД для продуктивной среды."""

    CONFIG_PATH = Path(__file__).resolve().parent / "config" / "config.ini"
    config = ConfigParser()
    config.read(CONFIG_PATH)

    section = "SQL Server"
    connection_url: URL = BaseDatabaseConfig.create(section, config).connection_url


engine = create_engine(DevelopmentDatabaseConfig.connection_url)


@dataclass(frozen=True)
class PersonsField:
    """Класс данных для понятного сопоставления полей в таблице 'persons'."""

    id = "num"  # Первичный ключ.
    surname = "fru"  # Поле 'Фамилия'.
    surname_in_latin = "last_lat"  # Поле 'Фамилия' на латинице.
    name = "name_rus"  # Поле 'Имя'.
    name_in_latin = "nla"  # Поле 'Имя' на латинице.
    patronymic = "och"  # Поле 'Отчество'.
    patronymic_in_latin = "oche"  # Поле 'Отчество' на латинице.
    citizenship = "ctz1"  # Поле 'Гражданство'.
    birthdate = "dob"  # Поле 'Дата рождения'.
    sex = "sex"  # Поле 'Пол'.
    state_of_birth = "pob"  # Поле 'Государство рождения'.
    birth_city = "cob"  # Поле 'Город рождения'.

    # Документ, удостоверяющий личность.
    passport_series = "pas_ser"  # Поле 'Серия паспорта'.
    passport_number = "pas_num"  # Поле 'Номер паспорта'.
    passport_issue_date = "pds"  # Поле 'Дата выдачи паспорта'.
    passport_expiration_date = "pde"  # Поле 'Дата окончания паспорта'.

    # Вид и реквизиты документа, подтверждающего право
    # на пребывание (проживание) в Российской Федерации.
    visa_requirement = "visa_priz"
    visa_series = "vis_ser"  # Поле 'Серия визы'.
    visa_number = "vis_num"  # Поле 'Номер визы'.
    visa_issue_date = "vis_start"  # Поле 'Дата выдачи визы'.
    visa_expiration_date = "vis_end"  # Поле 'Дата окончания визы'.

    telephone_number = "tel_nom"  # Поле 'Номер телефона'.

    entry_date = "d_enter"  # Поле 'Дата въезда'.
    departure_date = "date_okon"  # Поле 'Дата выезда'.

    migration_card_series = "mcs"  # Поле 'Серия миграционной карты'.
    migration_card_number = "mcn"  # Поле 'Номер миграционной карты'.

    room_number = "k"  # Поле 'Номер комнаты' в общежитии.
    start_date_of_hostel_contract = "dnd"  # Поле 'Дата начала договора с общежитием'.
    hostel_agreement = "dog_obsh"  # Поле 'Договор с общежитием'.

    availability_of_residence_permit_and_RVPO = "rf"  # Поле 'Наличие ВНЖ, РВПО'.
    date_of_issuance_of_residence_permit_and_RVPO = (
        "rfd"  # Поле 'Дата выдачи ВНЖ, РВПО'.
    )
    term_of_validity_of_residence_permit_and_RVPO = (
        "mot"  # Поле 'Срок действия ВНЖ, РВПО'.
    )
    residence_permit_and_RVPO_series = "ser"  # Поле 'Серия ВНЖ, РВПО'.
    residence_permit_and_RVPO_number = "nmr"  # Поле 'Номер ВНЖ, РВПО'.
    visa_multiplicity = "vis_krat"  # Поле 'Кратность визы'.
    visa_id = "vis_id"  # Поле 'Индентификатор визы'.
    study_direction = "gos_nap"  # Поле 'Направление обучения'.
    address_of_permanent_residence_in_the_country = (
        "proz"  # Поле 'Адрес постоянного проживания в стране'.
    )
    contract_number = "kontrakt"  # Поле 'Номер контракта'.
    term_of_study_from = "kont_start"  # Поле 'Срок обучения с'.
    term_of_study_on = "kont_end"  # Поле 'Срок обучения по'.
    date_of_commencement_of_the_contract_directions = (
        "gos_start"  # Поле 'Дата начала договора гос. Направления'.
    )
    termination_date_of_the_state_direction_contract = (
        "gos_end"  # Поле 'Дата окончания договора гос. Направления'.
    )
    previous_address = "star"  # Поле 'Прежний адрес'.
    date_of_visa_issuance = "d_poluch"  # Поле 'Дата получения визы'.
    visa_issuance_country = "str_poluch"  # Поле 'Страна выдачи визы'.
    visa_issuance_city = "city_poluch"  # Поле 'Город выдачи визы'.
    invitation_number = "num_prig"  # Поле 'Номер приглашения'.
    checkpoint = "kpp"  # Поле 'КПП'.
    date_of_medical_examination = "med"  # Поле 'Дата медицинского осмотра'.
    date_of_fingerprinting = "dak"  # Поле 'Дата дактилоскопии'.
    study_status = "uch_st_st"  # Поле 'Статус обучения'.
    study_program = "pr"  # Поле 'Программа обучения'.
    study_form = "fr"  # Поле 'Форма обучения'.
    specialization = "o_p"  # Поле 'Специальность'.
    order_number = "prikaz"  # Поле 'Номер приказа'.
    start_of_the_order = "prik_start"  # Поле 'Начало приказа'.
    status = "o_s"  # Поле 'Статус'.
    termination_date_of_the_employment_contract = (
        "d_naym"  # Поле 'Срок окончания договора найма'.
    )
    email = "email"  # Поле 'e-mail'.
    note = "prim"  # Поле 'Примечание'.


def execute_query(query: str) -> Optional[Sequence[Row[Tuple[Any, ...] | Any]]]:
    """Базовая функция, осуществляющая подключение к БД и исполняющая запрос."""
    try:
        with Session(engine) as session:
            debug(
                f"Подключение к базе данных '{DevelopmentDatabaseConfig.DATABASE_NAME}'..."
            )

            if ...:
                ...

            debug(f"Выполнение запроса '{query}'...")

            result = session.execute(text(query))

            debug(f"Успешное выполнение запроса '{query}'...")
            return result.all()
    except Exception as e:
        debug(f"Ошибка '{e}...'")
        return None


def fetch_all_data() -> Optional[Sequence[Row[Tuple[Any, ...] | Any]]]:
    """Получение всех данных из БД."""
    return execute_query("SELECT * FROM persons")


def fetch_person_by(
    logical_operator: str = "AND",
    **kwargs: Union[str, None],
) -> Optional[Sequence[Row[Tuple[Any, ...] | Any]]]:
    """
    Функция для получения данных о студенте.
    По умолчанию стоит точный фильтр.
    """
    if logical_operator not in ["AND", "OR"]:
        raise ValueError("logical_operator must be 'AND' or 'OR'")

    conditions = []
    debug_messages = []

    for key, value in kwargs.items():
        if value is not None:  # Проверяем, что значение не None
            if key == "surname":
                conditions.append(f"{PersonsField.surname} = '{value}'")
                debug_messages.append(f"Fetching person by surname: '{value}'")

            elif key == "name":
                conditions.append(f"{PersonsField.name} = '{value}'")
                debug_messages.append(f"Fetching person by name: '{value}'")

            elif key == "patronymic":
                conditions.append(f"{PersonsField.patronymic} = '{value}'")
                debug_messages.append(f"Fetching person by patronymic: '{value}'")

            elif key == "birthdate":
                conditions.append(f"{PersonsField.birthdate} = '{value}'")
                debug_messages.append(f"Fetching person by birthdate: '{value}'")

    if conditions:
        query = "SELECT * FROM persons WHERE " + f" {logical_operator} ".join(
            conditions
        )
        for message in debug_messages:
            debug(message)
        return execute_query(query)

    return None


BASE_QUERY = (
    "SELECT "
    "fru, last_lat, name_rus, nla, och, "
    "oche, ctz1, dob, sex, pob, "
    "cob, pas_ser, pas_num, pds, pde, "
    "visa_priz, vis_ser, vis_num, d_poluch, vis_end, "
    "tel_nom, d_enter, date_okon, mcs, mcn, "
    "k, dnd, dog_obsh, rf, rfd, "
    "mot, ser, nmr, vis_krat, vis_id, "
    "gos_nap, proz, kontrakt, kont_start, kont_end, "
    "gos_start, gos_end, star, d_poluch, str_poluch,"
    "city_poluch, num_prig, kpp, med, mot, "
    "uch_st_st, pr, fr, o_p, prikaz, "
    "prik_start, o_s, d_naym, email, prim "
    "FROM persons"
)


def check_rus_eng(text: str) -> bool:
    """Проверка текста на кириллицу."""
    return bool(regexp.search(r"[а-яА-ЯёЁ]", text))


def fetch_person(
    surname: Optional[str],
    name: Optional[str],
    patronymic: Optional[str],
    birthdate: Optional[datetime],
) -> Optional[List[List[str]]]:
    query = BASE_QUERY

    # Список условий и параметров
    conditions = []
    parameters = []

    # Добавляем условия в зависимости от переданных аргументов
    if surname is not None:
        parameters.append(surname)

        if check_rus_eng(surname):
            conditions.append(f"{PersonsField.surname} = ?")
        else:
            conditions.append(f"{PersonsField.surname_in_latin} = ?")

    if name is not None:
        parameters.append(name)

        if check_rus_eng(name):
            conditions.append(f"{PersonsField.name} = ?")
        else:
            conditions.append(f"{PersonsField.name_in_latin} = ?")

    if patronymic is not None:
        parameters.append(patronymic)

        if check_rus_eng(patronymic):
            conditions.append(f"{PersonsField.patronymic} = ?")
        else:
            conditions.append(f"{PersonsField.patronymic_in_latin} = ?")

    if birthdate != datetime(1900, 1, 1):
        conditions.append(f"{PersonsField.birthdate} = ?")
        parameters.append(birthdate)

    # Если есть условия, добавляем их к запросу
    if conditions:
        # query += " WHERE " + " OR ".join(conditions)
        query += f" WHERE {' AND '.join(conditions)}"

    # Выполняем запрос и возвращаем его
    result = execute_query(query)
    new_data = []

    for row in result:
        new_row = []
        for i, field in enumerate(row):
            if isinstance(field, datetime) or i in [4, 7]:
                if field is None:
                    # new_row.append("          ")
                    new_row.append(" " * 10)
                else:
                    date = field.strftime("%d%m%Y")  # Форматируем
                    new_row.append(date)
            else:
                new_row.append(str(field).strip() if field is not None else "")
        new_data.append(new_row)

    return new_data


def update_person(
    id: int,
    model,
) -> None:
    """Обновление записи в таблице 'persons' по ID."""
    query = "UPDATE persons SET "

    parameters = []
    updates = []
    update_par = [
        f"{PersonsField.id}",
        f"{PersonsField.surname}",
        f"{PersonsField.surname_in_latin}",
        f"{PersonsField.name}",
        f"{PersonsField.name_in_latin}",
        f"{PersonsField.patronymic}",
        f"{PersonsField.patronymic_in_latin}",
        f"{PersonsField.citizenship}",
        f"{PersonsField.birthdate}",
        f"{PersonsField.sex}",
        f"{PersonsField.state_of_birth}",
        f"{PersonsField.birth_city}",
        f"{PersonsField.passport_series}",
        f"{PersonsField.passport_number}",
        f"{PersonsField.passport_issue_date}",
        f"{PersonsField.passport_expiration_date}",
        f"{PersonsField.visa_requirement}",
        f"{PersonsField.visa_series}",
        f"{PersonsField.visa_number}",
        f"{PersonsField.visa_issue_date}",
        f"{PersonsField.visa_expiration_date}",
        f"{PersonsField.telephone_number}",
        f"{PersonsField.entry_date}",
        f"{PersonsField.departure_date}",
        f"{PersonsField.migration_card_series}",
        f"{PersonsField.migration_card_number}",
        f"{PersonsField.room_number}",
        f"{PersonsField.start_date_of_hostel_contract}",
        f"{PersonsField.hostel_agreement}",
        f"{PersonsField.availability_of_residence_permit_and_RVPO}",
        f"{PersonsField.date_of_issuance_of_residence_permit_and_RVPO}",
        f"{PersonsField.term_of_validity_of_residence_permit_and_RVPO}",
        f"{PersonsField.residence_permit_and_RVPO_series}",
        f"{PersonsField.residence_permit_and_RVPO_number}",
        f"{PersonsField.visa_multiplicity}",
        f"{PersonsField.visa_id}",
        f"{PersonsField.study_direction}",
        f"{PersonsField.address_of_permanent_residence_in_the_country}",
        f"{PersonsField.contract_number}",
        f"{PersonsField.term_of_study_from}",
        f"{PersonsField.term_of_study_on}",
        f"{PersonsField.date_of_commencement_of_the_contract_directions}",
        f"{PersonsField.termination_date_of_the_state_direction_contract}",
        f"{PersonsField.previous_address}",
        f"{PersonsField.date_of_visa_issuance}",
        f"{PersonsField.visa_issuance_country}",
        f"{PersonsField.visa_issuance_city}",
        f"{PersonsField.invitation_number}",
        f"{PersonsField.checkpoint}",
        f"{PersonsField.date_of_medical_examination}",
        f"{PersonsField.date_of_fingerprinting}",
        f"{PersonsField.study_status}",
        f"{PersonsField.study_program}",
        f"{PersonsField.study_form}",
        f"{PersonsField.specialization}",
        f"{PersonsField.order_number}",
        f"{PersonsField.start_of_the_order}",
        f"{PersonsField.status}",
        f"{PersonsField.termination_date_of_the_employment_contract}",
        f"{PersonsField.email}",
        f"{PersonsField.note}",
    ]

    for i, field in enumerate(update_par):
        if model.data_update[i] != "":
            updates.append(field)
            parameters.append(model.data_update[i])

    # Формируем запрос
    query += ", ".join(updates) + f" WHERE {PersonsField.id} = {id}"

    # Выводим для отладки (если нужно)
    debug(query)
    debug(parameters)

    # Выполняем запрос
    execute_query(query)
