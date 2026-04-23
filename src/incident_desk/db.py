"""SQLite persistence layer: schema, migrations, and CRUD helpers."""
from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import List, Optional, Tuple

from .config import DB_PATH


class DB:
    def __init__(self, path: Path = DB_PATH):
        self.path = path
        self.conn = sqlite3.connect(self.path)
        self.conn.row_factory = sqlite3.Row
        self._init_schema()

    def _init_schema(self):
        cur = self.conn.cursor()
        cur.executescript(
            """
            PRAGMA foreign_keys = ON;

            CREATE TABLE IF NOT EXISTS locations (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL UNIQUE
            );

            CREATE TABLE IF NOT EXISTS units (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL UNIQUE,
                category TEXT DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS incident_types (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL UNIQUE
            );

            CREATE TABLE IF NOT EXISTS driver_codes (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL UNIQUE,
                sort_order INTEGER DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS incidents (
                id INTEGER PRIMARY KEY,
                location_id INTEGER,
                type TEXT NOT NULL,
                reported_at TEXT NOT NULL,
                dispatched_at TEXT DEFAULT '',
                arrived_at TEXT DEFAULT '',
                cleared_at TEXT DEFAULT '',
                disposition TEXT DEFAULT '',
                car_number TEXT DEFAULT '',
                driver_code TEXT DEFAULT '',
                is_cleared INTEGER DEFAULT 0,
                FOREIGN KEY(location_id) REFERENCES locations(id)
            );

            CREATE TABLE IF NOT EXISTS incident_units (
                id INTEGER PRIMARY KEY,
                incident_id INTEGER NOT NULL,
                unit_id INTEGER NOT NULL,
                role TEXT NOT NULL CHECK(role in ('primary','backup')),
                FOREIGN KEY(incident_id) REFERENCES incidents(id) ON DELETE CASCADE,
                FOREIGN KEY(unit_id) REFERENCES units(id)
            );

            CREATE TABLE IF NOT EXISTS notes (
                id INTEGER PRIMARY KEY,
                incident_id INTEGER NOT NULL,
                ts TEXT NOT NULL,
                body TEXT NOT NULL,
                FOREIGN KEY(incident_id) REFERENCES incidents(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS billables (
                id INTEGER PRIMARY KEY,
                incident_id INTEGER NOT NULL,
                body TEXT NOT NULL,
                FOREIGN KEY(incident_id) REFERENCES incidents(id) ON DELETE CASCADE
            );
            """
        )
        self.conn.commit()
        self._migrate()

    def _migrate(self):
        cur = self.conn.cursor()
        for stmt in [
            "ALTER TABLE locations ADD COLUMN sort_order INTEGER DEFAULT 0",
            "ALTER TABLE incident_types ADD COLUMN sort_order INTEGER DEFAULT 0",
            "ALTER TABLE units ADD COLUMN sort_order INTEGER DEFAULT 0",
            "ALTER TABLE incidents ADD COLUMN car_number TEXT DEFAULT ''",
            "ALTER TABLE incidents ADD COLUMN driver_code TEXT DEFAULT ''",
        ]:
            try:
                cur.execute(stmt)
            except Exception:
                pass  # column already exists
        # Initialise sort_order for any rows that still have 0 (existing data)
        cur.execute("UPDATE locations SET sort_order = id WHERE sort_order = 0")
        cur.execute("UPDATE incident_types SET sort_order = id WHERE sort_order = 0")
        cur.execute("UPDATE units SET sort_order = id WHERE sort_order = 0")
        cur.execute("UPDATE driver_codes SET sort_order = id WHERE sort_order = 0")
        self.conn.commit()

    # ---- CRUD helpers
    def list_locations(self) -> List[sqlite3.Row]:
        return self.conn.execute("SELECT * FROM locations ORDER BY sort_order, id").fetchall()

    def add_location(self, name: str):
        self.conn.execute(
            "INSERT OR IGNORE INTO locations(name, sort_order) "
            "VALUES(?, (SELECT COALESCE(MAX(sort_order), 0) + 1 FROM locations))",
            (name.strip(),),
        )
        self.conn.commit()

    def rename_location(self, loc_id: int, new_name: str):
        self.conn.execute("UPDATE locations SET name=? WHERE id=?", (new_name.strip(), loc_id))
        self.conn.commit()

    def location_incident_count(self, loc_id: int) -> int:
        return self.conn.execute(
            "SELECT COUNT(*) FROM incidents WHERE location_id=?", (loc_id,)
        ).fetchone()[0]

    def delete_location(self, loc_id: int):
        self.conn.execute("DELETE FROM locations WHERE id=?", (loc_id,))
        self.conn.commit()

    def list_units(self) -> List[sqlite3.Row]:
        return self.conn.execute("SELECT * FROM units ORDER BY sort_order, id").fetchall()

    def list_units_with_availability(self) -> List[sqlite3.Row]:
        """Returns all units with a computed `available` column (1=available, 0=unavailable)."""
        return self.conn.execute("""
            SELECT u.*,
                CASE WHEN EXISTS (
                    SELECT 1 FROM incident_units iu
                    JOIN incidents i ON i.id = iu.incident_id
                    WHERE iu.unit_id = u.id AND iu.role = 'primary' AND i.is_cleared = 0
                ) THEN 0 ELSE 1 END AS available
            FROM units u ORDER BY u.sort_order, u.id
        """).fetchall()

    def list_available_units(self, exclude_incident_id: Optional[int] = None) -> List[sqlite3.Row]:
        """Returns units not currently assigned as primary on any active incident.
        exclude_incident_id: ignore assignments belonging to this incident (used when editing)."""
        if exclude_incident_id:
            return self.conn.execute("""
                SELECT u.* FROM units u
                WHERE NOT EXISTS (
                    SELECT 1 FROM incident_units iu
                    JOIN incidents i ON i.id = iu.incident_id
                    WHERE iu.unit_id = u.id AND iu.role = 'primary'
                      AND i.is_cleared = 0 AND i.id != ?
                ) ORDER BY u.sort_order, u.id
            """, (exclude_incident_id,)).fetchall()
        return self.conn.execute("""
            SELECT u.* FROM units u
            WHERE NOT EXISTS (
                SELECT 1 FROM incident_units iu
                JOIN incidents i ON i.id = iu.incident_id
                WHERE iu.unit_id = u.id AND iu.role = 'primary' AND i.is_cleared = 0
            ) ORDER BY u.sort_order, u.id
        """).fetchall()

    def add_unit(self, name: str, category: str = ""):
        self.conn.execute(
            "INSERT OR IGNORE INTO units(name, category, sort_order) "
            "VALUES(?,?, (SELECT COALESCE(MAX(sort_order), 0) + 1 FROM units))",
            (name.strip(), category.strip()),
        )
        self.conn.commit()

    def update_unit(self, unit_id: int, name: str, category: str):
        self.conn.execute("UPDATE units SET name=?, category=? WHERE id=?", (name.strip(), category.strip(), unit_id))
        self.conn.commit()

    def unit_incident_count(self, unit_id: int) -> int:
        return self.conn.execute(
            "SELECT COUNT(*) FROM incident_units WHERE unit_id=?", (unit_id,)
        ).fetchone()[0]

    def delete_unit(self, unit_id: int):
        self.conn.execute("DELETE FROM units WHERE id=?", (unit_id,))
        self.conn.commit()

    def list_incident_types(self) -> List[sqlite3.Row]:
        return self.conn.execute("SELECT * FROM incident_types ORDER BY sort_order, id").fetchall()

    def add_incident_type(self, name: str):
        self.conn.execute(
            "INSERT OR IGNORE INTO incident_types(name, sort_order) "
            "VALUES(?, (SELECT COALESCE(MAX(sort_order), 0) + 1 FROM incident_types))",
            (name.strip(),),
        )
        self.conn.commit()

    def set_sort_order(self, table: str, ordered_ids: list):
        """Assign sort_order 1..N to rows in the given order."""
        if table not in ("locations", "units", "incident_types", "driver_codes"):
            raise ValueError(f"unknown table: {table}")
        cur = self.conn.cursor()
        for idx, row_id in enumerate(ordered_ids, start=1):
            cur.execute(f"UPDATE {table} SET sort_order=? WHERE id=?", (idx, row_id))
        self.conn.commit()

    def rename_incident_type(self, type_id: int, new_name: str):
        self.conn.execute("UPDATE incident_types SET name=? WHERE id=?", (new_name.strip(), type_id))
        self.conn.commit()

    def incident_type_incident_count(self, type_name: str) -> int:
        return self.conn.execute(
            "SELECT COUNT(*) FROM incidents WHERE type=?", (type_name,)
        ).fetchone()[0]

    def delete_incident_type(self, type_id: int):
        self.conn.execute("DELETE FROM incident_types WHERE id=?", (type_id,))
        self.conn.commit()

    def list_driver_codes(self) -> List[sqlite3.Row]:
        return self.conn.execute("SELECT * FROM driver_codes ORDER BY sort_order, id").fetchall()

    def add_driver_code(self, name: str):
        self.conn.execute(
            "INSERT OR IGNORE INTO driver_codes(name, sort_order) "
            "VALUES(?, (SELECT COALESCE(MAX(sort_order), 0) + 1 FROM driver_codes))",
            (name.strip(),),
        )
        self.conn.commit()

    def rename_driver_code(self, code_id: int, new_name: str):
        self.conn.execute("UPDATE driver_codes SET name=? WHERE id=?", (new_name.strip(), code_id))
        self.conn.commit()

    def driver_code_incident_count(self, code_name: str) -> int:
        return self.conn.execute(
            "SELECT COUNT(*) FROM incidents WHERE driver_code=?", (code_name,)
        ).fetchone()[0]

    def delete_driver_code(self, code_id: int):
        self.conn.execute("DELETE FROM driver_codes WHERE id=?", (code_id,))
        self.conn.commit()

    def create_incident(self, location_id: Optional[int], type_name: str, reported_at: str,
                         dispatched_at: str = "", arrived_at: str = "", cleared_at: str = "",
                         disposition: str = "", is_cleared: int = 0,
                         primary_unit_id: Optional[int] = None, backup_unit_ids: Optional[List[int]] = None,
                         car_number: str = "", driver_code: str = "") -> int:
        cur = self.conn.cursor()
        cur.execute(
            """
            INSERT INTO incidents(location_id, type, reported_at, dispatched_at, arrived_at, cleared_at, disposition, car_number, driver_code, is_cleared)
            VALUES(?,?,?,?,?,?,?,?,?,?)
            """,
            (location_id, type_name, reported_at, dispatched_at, arrived_at, cleared_at, disposition, car_number, driver_code, is_cleared),
        )
        inc_id = cur.lastrowid
        if primary_unit_id:
            cur.execute(
                "INSERT INTO incident_units(incident_id, unit_id, role) VALUES(?,?, 'primary')",
                (inc_id, primary_unit_id),
            )
        if backup_unit_ids:
            for uid in backup_unit_ids:
                cur.execute(
                    "INSERT INTO incident_units(incident_id, unit_id, role) VALUES(?,?, 'backup')",
                    (inc_id, uid),
                )
        self.conn.commit()
        return inc_id

    def update_incident(self, inc_id: int, location_id: Optional[int], type_name: str, reported_at: str,
                         dispatched_at: str, arrived_at: str, cleared_at: str, disposition: str, is_cleared: int,
                         primary_unit_id: Optional[int], backup_unit_ids: List[int],
                         car_number: str = "", driver_code: str = ""):
        cur = self.conn.cursor()
        cur.execute(
            "UPDATE incidents SET location_id=?, type=?, reported_at=?, dispatched_at=?, arrived_at=?, cleared_at=?, disposition=?, car_number=?, driver_code=?, is_cleared=? WHERE id=?",
            (location_id, type_name, reported_at, dispatched_at, arrived_at, cleared_at, disposition, car_number, driver_code, is_cleared, inc_id),
        )
        # reset assignments
        cur.execute("DELETE FROM incident_units WHERE incident_id=?", (inc_id,))
        if primary_unit_id:
            cur.execute("INSERT INTO incident_units(incident_id, unit_id, role) VALUES(?,?, 'primary')", (inc_id, primary_unit_id))
        for uid in backup_unit_ids:
            cur.execute("INSERT INTO incident_units(incident_id, unit_id, role) VALUES(?,?, 'backup')", (inc_id, uid))
        self.conn.commit()

    def delete_incident(self, inc_id: int):
        self.conn.execute("DELETE FROM incidents WHERE id=?", (inc_id,))
        self.conn.commit()

    def add_note(self, inc_id: int, ts: str, body: str):
        self.conn.execute("INSERT INTO notes(incident_id, ts, body) VALUES(?,?,?)", (inc_id, ts, body))
        self.conn.commit()

    def list_notes(self, inc_id: int) -> List[sqlite3.Row]:
        return self.conn.execute("SELECT * FROM notes WHERE incident_id=? ORDER BY ts", (inc_id,)).fetchall()

    def add_billable(self, inc_id: int, body: str):
        self.conn.execute("INSERT INTO billables(incident_id, body) VALUES(?,?)", (inc_id, body))
        self.conn.commit()

    def list_billables(self, inc_id: int) -> List[sqlite3.Row]:
        return self.conn.execute("SELECT * FROM billables WHERE incident_id=? ORDER BY id", (inc_id,)).fetchall()

    def set_cleared(self, inc_id: int, cleared: bool, cleared_at: Optional[str] = None):
        cur = self.conn.cursor()
        cur.execute(
            "UPDATE incidents SET is_cleared=?, cleared_at=? WHERE id=?",
            (1 if cleared else 0, cleared_at or "", inc_id),
        )
        self.conn.commit()

    def fetch_board(self, loc_filter: Optional[int], type_filter: Optional[str],
                    start_date: Optional[str], end_date: Optional[str]) -> List[sqlite3.Row]:
        q = [
            "SELECT i.*, l.name as location_name,",
            "GROUP_CONCAT(CASE iu.role WHEN 'primary' THEN u.name END) as primary_units,",
            "GROUP_CONCAT(CASE iu.role WHEN 'backup' THEN u.name END) as backup_units",
            "FROM incidents i",
            "LEFT JOIN locations l ON l.id=i.location_id",
            "LEFT JOIN incident_units iu ON iu.incident_id=i.id",
            "LEFT JOIN units u ON u.id=iu.unit_id",
        ]
        where = []
        params: List[object] = []
        if loc_filter:
            where.append("i.location_id=?")
            params.append(loc_filter)
        if type_filter:
            where.append("i.type=?")
            params.append(type_filter)
        if start_date:
            where.append("date(i.reported_at) >= date(?)")
            params.append(start_date)
        if end_date:
            where.append("date(i.reported_at) <= date(?)")
            params.append(end_date)
        if where:
            q.append("WHERE " + " AND ".join(where))
        q.append("GROUP BY i.id ORDER BY i.is_cleared ASC, i.reported_at DESC")
        sql = "\n".join(q)
        return self.conn.execute(sql, params).fetchall()

    def get_incident(self, inc_id: int) -> sqlite3.Row:
        return self.conn.execute("SELECT * FROM incidents WHERE id=?", (inc_id,)).fetchone()

    def get_incident_assignments(self, inc_id: int) -> Tuple[Optional[int], List[int]]:
        primary = self.conn.execute(
            "SELECT unit_id FROM incident_units WHERE incident_id=? AND role='primary'",
            (inc_id,),
        ).fetchone()
        backups = [r[0] for r in self.conn.execute(
            "SELECT unit_id FROM incident_units WHERE incident_id=? AND role='backup'",
            (inc_id,),
        ).fetchall()]
        return (primary[0] if primary else None, backups)
