import os
import sqlite3

DATA_DIR = "data"
BANKS_DB_FILENAMES = (
    os.path.join(DATA_DIR, "banks.db"),
    "banks.db",
)


class BanksDatabaseError(Exception):
    """Readable error for bank database problems shown in the UI."""


def get_banks_db_path():
    for path in BANKS_DB_FILENAMES:
        if os.path.exists(path):
            return path

    raise BanksDatabaseError(
        "База банков banks.db не найдена. Поместите файл banks.db в папку data "
        "или в корень проекта."
    )


def connect_banks_db():
    db_path = get_banks_db_path()

    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        ensure_banks_schema(conn)
        return conn
    except sqlite3.Error as exc:
        raise BanksDatabaseError(f"Не удалось открыть базу банков: {exc}") from exc


def ensure_banks_schema(conn):
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS banks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            inn TEXT,
            ogrn TEXT,
            address TEXT,
            source_text TEXT NOT NULL DEFAULT ''
        )
    """)

    existing_columns = {
        row[1]
        for row in cursor.execute("PRAGMA table_info(banks)").fetchall()
    }

    for column_name, column_type in {
        "name": "TEXT NOT NULL DEFAULT ''",
        "inn": "TEXT",
        "ogrn": "TEXT",
        "address": "TEXT",
        "source_text": "TEXT",
    }.items():
        if column_name not in existing_columns:
            cursor.execute(f"ALTER TABLE banks ADD COLUMN {column_name} {column_type}")

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS bank_aliases (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bank_id INTEGER NOT NULL,
            alias TEXT NOT NULL,
            FOREIGN KEY (bank_id) REFERENCES banks(id) ON DELETE CASCADE
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS bank_branches (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bank_id INTEGER NOT NULL,
            branch TEXT NOT NULL,
            FOREIGN KEY (bank_id) REFERENCES banks(id) ON DELETE CASCADE
        )
    """)

    cursor.execute("CREATE INDEX IF NOT EXISTS idx_banks_name ON banks(name)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_banks_inn ON banks(inn)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_banks_ogrn ON banks(ogrn)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_bank_aliases_alias ON bank_aliases(alias)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_bank_aliases_bank_id ON bank_aliases(bank_id)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_bank_branches_bank_id ON bank_branches(bank_id)")

    conn.commit()


def normalize_multiline(value):
    if isinstance(value, list):
        lines = value
    else:
        lines = str(value or "").splitlines()

    return [line.strip() for line in lines if line.strip()]


def row_to_bank(row):
    return {
        "id": row["id"],
        "name": row["name"] or "",
        "inn": row["inn"] or "",
        "ogrn": row["ogrn"] or "",
        "address": row["address"] or "",
    }


def build_source_text(data):
    name = (data.get("name") or "").strip()
    inn = (data.get("inn") or "").strip()
    ogrn = (data.get("ogrn") or "").strip()
    address = (data.get("address") or "").strip()

    return f"{name} ИНН {inn} ОГРН {ogrn} {address}".strip()


def quote_identifier(identifier):
    return f'"{identifier.replace(chr(34), chr(34) + chr(34))}"'


def get_banks_columns(cursor):
    return [
        {
            "name": row[1],
            "type": (row[2] or "").upper(),
            "notnull": bool(row[3]),
            "default": row[4],
            "pk": bool(row[5]),
        }
        for row in cursor.execute("PRAGMA table_info(banks)").fetchall()
    ]


def get_required_column_value(column, data):
    column_name = column["name"]

    if column_name == "source_text":
        return build_source_text(data)

    if column_name in data:
        return data.get(column_name) or ""

    column_type = column["type"]
    if "INT" in column_type:
        return 0
    if any(type_name in column_type for type_name in ("REAL", "FLOA", "DOUB")):
        return 0.0
    if "BLOB" in column_type:
        return b""

    return ""


def build_insert_values(cursor, data):
    base_values = {
        "name": (data.get("name") or "").strip(),
        "inn": (data.get("inn") or "").strip(),
        "ogrn": (data.get("ogrn") or "").strip(),
        "address": (data.get("address") or "").strip(),
        "source_text": build_source_text(data),
    }
    insert_values = {}

    for column in get_banks_columns(cursor):
        column_name = column["name"]

        if column["pk"]:
            continue

        if column_name in base_values:
            insert_values[column_name] = base_values[column_name]
        elif column["notnull"] and column["default"] is None:
            insert_values[column_name] = get_required_column_value(column, data)

    return insert_values


def search_banks(query, limit=100):
    conn = connect_banks_db()
    cursor = conn.cursor()
    query = (query or "").strip()

    if query:
        cursor.execute("""
            SELECT b.id, b.name, b.inn, b.ogrn, b.address,
                   GROUP_CONCAT(a.alias, ' ') AS aliases_text
            FROM banks b
            LEFT JOIN bank_aliases a ON a.bank_id = b.id
            GROUP BY b.id, b.name, b.inn, b.ogrn, b.address
            ORDER BY b.name COLLATE NOCASE
        """)
        query_value = query.casefold()
        rows = []

        for row in cursor.fetchall():
            searchable_text = " ".join((
                row["name"] or "",
                row["inn"] or "",
                row["ogrn"] or "",
                row["aliases_text"] or "",
            )).casefold()

            if query_value in searchable_text:
                rows.append(row_to_bank(row))

            if len(rows) >= limit:
                break
    else:
        cursor.execute("""
            SELECT id, name, inn, ogrn, address
            FROM banks
            ORDER BY name COLLATE NOCASE
            LIMIT ?
        """, (limit,))
        rows = [row_to_bank(row) for row in cursor.fetchall()]

    conn.close()
    return rows


def get_bank_details(bank_id):
    conn = connect_banks_db()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT id, name, inn, ogrn, address
        FROM banks
        WHERE id = ?
    """, (bank_id,))
    bank_row = cursor.fetchone()

    if not bank_row:
        conn.close()
        return None

    bank = row_to_bank(bank_row)

    cursor.execute("""
        SELECT alias
        FROM bank_aliases
        WHERE bank_id = ?
        ORDER BY alias COLLATE NOCASE
    """, (bank_id,))
    bank["aliases"] = [row["alias"] for row in cursor.fetchall()]

    cursor.execute("""
        SELECT branch
        FROM bank_branches
        WHERE bank_id = ?
        ORDER BY branch COLLATE NOCASE
    """, (bank_id,))
    bank["branches"] = [row["branch"] for row in cursor.fetchall()]

    conn.close()
    return bank


def add_bank(data):
    name = (data.get("name") or "").strip()
    if not name:
        raise BanksDatabaseError("Название банка обязательно для заполнения.")

    conn = connect_banks_db()
    cursor = conn.cursor()

    try:
        insert_values = build_insert_values(cursor, data)
        columns_sql = ", ".join(quote_identifier(column) for column in insert_values)
        placeholders_sql = ", ".join("?" for _ in insert_values)
        cursor.execute(
            f"INSERT INTO banks ({columns_sql}) VALUES ({placeholders_sql})",
            tuple(insert_values.values()),
        )
        bank_id = cursor.lastrowid
        save_bank_related_rows(cursor, bank_id, data)
        conn.commit()
        return bank_id
    except sqlite3.Error as exc:
        conn.rollback()
        raise BanksDatabaseError(f"Не удалось добавить банк: {exc}") from exc
    finally:
        conn.close()


def update_bank(bank_id, data):
    name = (data.get("name") or "").strip()
    if not name:
        raise BanksDatabaseError("Название банка обязательно для заполнения.")

    conn = connect_banks_db()
    cursor = conn.cursor()

    try:
        update_values = {
            "name": name,
            "inn": (data.get("inn") or "").strip(),
            "ogrn": (data.get("ogrn") or "").strip(),
            "address": (data.get("address") or "").strip(),
        }
        existing_columns = {column["name"] for column in get_banks_columns(cursor)}

        if "source_text" in existing_columns:
            update_values["source_text"] = build_source_text(data)

        set_sql = ", ".join(
            f"{quote_identifier(column)} = ?"
            for column in update_values
            if column in existing_columns
        )
        params = [
            value
            for column, value in update_values.items()
            if column in existing_columns
        ]
        params.append(bank_id)

        cursor.execute(
            f"UPDATE banks SET {set_sql} WHERE id = ?",
            tuple(params),
        )

        if cursor.rowcount == 0:
            raise BanksDatabaseError("Банк не найден в базе данных.")

        if "aliases" in data:
            cursor.execute("DELETE FROM bank_aliases WHERE bank_id = ?", (bank_id,))
            save_bank_aliases(cursor, bank_id, data.get("aliases"))

        if "branches" in data:
            cursor.execute("DELETE FROM bank_branches WHERE bank_id = ?", (bank_id,))
            save_bank_branches(cursor, bank_id, data.get("branches"))

        conn.commit()
    except BanksDatabaseError:
        conn.rollback()
        raise
    except sqlite3.Error as exc:
        conn.rollback()
        raise BanksDatabaseError(f"Не удалось сохранить изменения банка: {exc}") from exc
    finally:
        conn.close()


def delete_bank(bank_id):
    conn = connect_banks_db()
    cursor = conn.cursor()

    try:
        cursor.execute("DELETE FROM bank_aliases WHERE bank_id = ?", (bank_id,))
        cursor.execute("DELETE FROM bank_branches WHERE bank_id = ?", (bank_id,))
        cursor.execute("DELETE FROM banks WHERE id = ?", (bank_id,))

        if cursor.rowcount == 0:
            raise BanksDatabaseError("Банк не найден в базе данных.")

        conn.commit()
    except BanksDatabaseError:
        conn.rollback()
        raise
    except sqlite3.Error as exc:
        conn.rollback()
        raise BanksDatabaseError(f"Не удалось удалить банк: {exc}") from exc
    finally:
        conn.close()


def save_bank_related_rows(cursor, bank_id, data):
    save_bank_aliases(cursor, bank_id, data.get("aliases"))
    save_bank_branches(cursor, bank_id, data.get("branches"))


def save_bank_aliases(cursor, bank_id, aliases_value):
    aliases = normalize_multiline(aliases_value)

    cursor.executemany("""
        INSERT INTO bank_aliases (bank_id, alias)
        VALUES (?, ?)
    """, [(bank_id, alias) for alias in aliases])


def save_bank_branches(cursor, bank_id, branches_value):
    branches = normalize_multiline(branches_value)

    cursor.executemany("""
        INSERT INTO bank_branches (bank_id, branch)
        VALUES (?, ?)
    """, [(bank_id, branch) for branch in branches])
