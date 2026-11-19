import unittest
from unittest.mock import patch

from sqlalchemy import create_engine, text
from sqlalchemy.pool import StaticPool

from app.database import SQLAlchemyConnection
from app.lookups import _get_all_livestock_ages


class LivestockAgeLookupTests(unittest.TestCase):
    def setUp(self):
        self.engine = create_engine(
            "sqlite://",
            connect_args={"check_same_thread": False},
            poolclass=StaticPool,
        )
        with self.engine.begin() as connection:
            connection.execute(text(
                "CREATE TABLE agetype ("
                "ageID TEXT, ageName TEXT, animalTypeID TEXT, animalID TEXT)"
            ))
            connection.execute(text(
                "INSERT INTO agetype VALUES "
                "('ageID01', '(၃) နှစ်အထက်', 'atID01', ''), "
                "('ageID07', 'ငယ်', 'atID03', 'lsID01'), "
                "('ageID08', 'လတ်', 'atID03', 'lsID01'), "
                "('ageID17', 'ကြီး', 'lsID07', 'atID03')"
            ))

        self.patch = patch(
            "app.lookups.get_db_connection",
            side_effect=lambda: SQLAlchemyConnection(self.engine.connect()),
        )
        self.patch.start()

    def tearDown(self):
        self.patch.stop()
        self.engine.dispose()

    def test_livestock_age_rows_are_returned_and_normalized(self):
        rows = _get_all_livestock_ages()
        self.assertEqual(
            rows,
            [
                {"ageID": "ageID07", "ageName": "ငယ်", "animalID": "lsID01"},
                {"ageID": "ageID08", "ageName": "လတ်", "animalID": "lsID01"},
                {"ageID": "ageID17", "ageName": "ကြီး", "animalID": "lsID07"},
            ],
        )


if __name__ == "__main__":
    unittest.main()
