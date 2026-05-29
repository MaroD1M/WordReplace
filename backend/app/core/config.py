from pydantic import BaseModel


class Settings(BaseModel):
    app_name: str = "WordReplace API"
    app_version: str = "2.0.0"
    database_url: str = "sqlite:////app/data/wordreplace.db"


settings = Settings()
