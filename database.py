from sqlalchemy import create_engine
from config import settings
from contextlib import contextmanager


engine = create_engine(settings.db_url)

@contextmanager
def DB_Connect():
    db = engine.connect()
    try:
        yield db
    finally:
            db.close()
