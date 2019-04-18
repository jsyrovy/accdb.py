import pypyodbc


class Accdb:
    def __init__(self, path):
        """Class for communication with Access DB

        :param path: Access DB absolute file path"""
        pypyodbc.lowercase = False
        self._path = path
        self._conn = pypyodbc.connect(f"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};Dbq={path};")
        self._cursor = self._conn.cursor()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.commit()
        self.connection.close()

    @property
    def path(self):
        return self._path

    @property
    def connection(self):
        return self._conn

    @property
    def cursor(self):
        return self._cursor

    @staticmethod
    def convert_date(dt):
        """Returns string with datetime format used in Access

        :param dt: DateTime"""
        return f"#{dt.month}/{dt.day}/{dt.year} {dt.hour}:{dt.minute}:{dt.second}#"

    def commit(self):
        self.connection.commit()

    def execute(self, qry):
        self.cursor.execute(qry)

    def fetchall(self):
        return self.cursor.fetchall()

    def fetchone(self):
        return self.cursor.fetchone()

    def query(self, qry):
        """Returns list with query result

        :param qry: SQL query string"""
        self.execute(qry)
        return self.fetchall()

    def close(self):
        self.commit()
        self.cursor.close()
        self.connection.close()

    def qry_header(self, qry):
        self.execute(qry)
        return [s[0] for s in self.cursor.description]

    def table_header(self, table):
        return self.qry_header(f"SELECT * FROM {table};")
