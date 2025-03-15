from configparser import ConfigParser
from dataclasses import dataclass
from pathlib import Path
import os
import sys
from sqlalchemy import URL

engine = None

def get_connection_string(
    driver: str,
    server_name: str,
    database_name: str,
    username: str,
    password: str,
) -> str:
    """Функция для получения строки соединения к БД."""
    return (
        f"DRIVER={driver};"
        f"SERVER={server_name}\\SQLEXPRESS;"
        f"PORT=1433;"
        f"DATABASE={database_name};"
        f"UID={username};"
        f"PWD={password};"
        "Encrypt=no"
    )


def create_database_config(section: str, config: ConfigParser,  user, passw):
    """Функция для создания конфигурации БД."""
    driver = config.get(section, "DRIVER")
    server_name = config.get(section, "SERVER_NAME")
    database_name = config.get(section, "DATABASE_NAME")
    username = user
    password = passw

    connection_string = get_connection_string(
        driver,
        server_name,
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


@dataclass(frozen=False)
class BaseDatabaseConfig:
    """Базовый класс конфигурации БД."""

    section: str
    connection_url: URL
    username: str
    password: str

    @classmethod
    def create(cls, section: str, config: ConfigParser, username, password):
        """Создает экземпляр класса с нужной конфигурацией."""
        connection_url = create_database_config(section, config, username, password)
        return cls(section=section, connection_url=connection_url, username=username, password=password)


@dataclass(frozen=False)
class DevelopmentDatabaseConfig(BaseDatabaseConfig):
    """Класс-конфигурация БД для разработки."""
    def __init__(self, username, password):
        path_progs = Path(__file__).parent
        """Возвращает корректный путь к файлу, работает и в `.exe`, и в обычном запуске Python"""
        if getattr(sys, '_MEIPASS', False):  # Проверяем, запущен ли скрипт в режиме PyInstaller
            CONFIG_PATH = os.path.join(sys._MEIPASS, "config.ini")
        else:
            CONFIG_PATH = os.path.join(path_progs, "config\\config.ini")

        config = ConfigParser()
        config.read(CONFIG_PATH)

        section = "Test_Maks"
        self.connection_url: URL = BaseDatabaseConfig.create(section, config, username, password).connection_url


@dataclass(frozen=False)
class ProductionDatabaseConfig(BaseDatabaseConfig):
    """Класс-конфигурация БД для продуктивной среды."""
    def __init__(self, username, password):
        path_progs = Path(__file__).parent
        """Возвращает корректный путь к файлу, работает и в `.exe`, и в обычном запуске Python"""
        if getattr(sys, '_MEIPASS', False):  # Проверяем, запущен ли скрипт в режиме PyInstaller
            CONFIG_PATH = os.path.join(sys._MEIPASS, "config.ini")
        else:
            CONFIG_PATH = os.path.join(path_progs, "config\\config.ini")

        config = ConfigParser()
        config.read(CONFIG_PATH)

        section = "SQL Server"
        self.connection_url: URL = BaseDatabaseConfig.create(section, config, username, password).connection_url
