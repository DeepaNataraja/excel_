
from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    env_name: str = "Local"
    base_url: str = "http://localhost:8000"
    db_url: str = "mysql+pymysql://root@localhost:3306/task1"

    class Config:
        env_file = ".env"

settings= Settings()