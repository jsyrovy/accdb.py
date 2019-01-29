# accdb.py

Python class based on [pypyodbc](https://github.com/jiangwen365/pypyodbc) for communication with Microsoft Access Database.

## Prerequisites

[Pypyodbc](https://github.com/jiangwen365/pypyodbc) has to be installed.

```
pip install pypyodbc
```

## Usage

```python
import datetime

import accdb

# Access DB absolute file path
path = r"c:\my_db.accdb"

# You can use Accdb.convert_date(dt) to convert datetime to Access format
qry_insert = f"INSERT INTO my_first_table (item, dt) VALUES ('New Item', {accdb.Accdb.convert_date(datetime.datetime.now())});"
qry_select = "SELECT * FROM my_second_table;"

# Classic way
db = accdb.Accdb(path)
db.execute(qry_insert)
print(db.query(qry_select))
db.close()

del db

# Using with statement
with accdb.Accdb(path) as db:
    db.execute(qry_insert)
    print(db.query(qry_select))
```