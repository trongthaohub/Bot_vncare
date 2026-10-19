import sqlite3
import os
from datetime import datetime

class Database:
    """Class quản lý cơ sở dữ liệu SQLite3 cho bot."""
    def __init__(self, db_path: str = None):
        if not db_path:
            db_dir = os.path.dirname(os.path.abspath(__file__))
            db_path = os.path.join(db_dir, "bot.db")
        
        # Đảm bảo thư mục chứa database tồn tại
        os.makedirs(os.path.dirname(os.path.abspath(db_path)), exist_ok=True)
        self.db_path = db_path
        self.init_db()

    def get_connection(self):
        """Tạo và trả về kết nối SQLite."""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def init_db(self):
        """Khởi tạo các bảng cơ sở dữ liệu."""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # Bảng users
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                user_id INTEGER PRIMARY KEY,
                username TEXT,
                first_name TEXT,
                joined_at TEXT
            )
        """)
        
        # Bảng messages
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS messages (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                message_text TEXT,
                created_at TEXT,
                FOREIGN KEY (user_id) REFERENCES users (user_id)
            )
        """)

        # Bảng departments (Khoa)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS departments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                hospital_id TEXT,
                hospital_url TEXT,
                dept_id TEXT,
                dept_name TEXT,
                updated_at TEXT,
                UNIQUE(hospital_id, dept_id)
            )
        """)

        # Bảng rooms (Phòng)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS rooms (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                hospital_id TEXT,
                dept_id TEXT,
                room_id TEXT,
                room_name TEXT,
                updated_at TEXT,
                UNIQUE(hospital_id, dept_id, room_id)
            )
        """)
        
        conn.commit()
        conn.close()

    def save_user(self, user_id: int, username: str, first_name: str):
        """Lưu hoặc cập nhật thông tin người dùng."""
        conn = self.get_connection()
        cursor = conn.cursor()
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        cursor.execute("""
            INSERT INTO users (user_id, username, first_name, joined_at)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(user_id) DO UPDATE SET
                username = excluded.username,
                first_name = excluded.first_name
        """, (user_id, username, first_name, now))
        
        conn.commit()
        conn.close()

    def log_message(self, user_id: int, text: str):
        """Lưu lịch sử tin nhắn."""
        conn = self.get_connection()
        cursor = conn.cursor()
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        cursor.execute("""
            INSERT INTO messages (user_id, message_text, created_at)
            VALUES (?, ?, ?)
        """, (user_id, text, now))
        
        conn.commit()
        conn.close()

    def save_departments_and_rooms(self, hospital_id: str, hospital_url: str, departments: list) -> int:
        """
        Lưu danh sách khoa và phòng vào database.
        Trả về tổng số phòng đã lưu.
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        total_rooms = 0

        for dept in departments:
            d_id = str(dept["dept_id"])
            d_name = str(dept["dept_name"])
            cursor.execute("""
                INSERT INTO departments (hospital_id, hospital_url, dept_id, dept_name, updated_at)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(hospital_id, dept_id) DO UPDATE SET
                    dept_name = excluded.dept_name,
                    hospital_url = excluded.hospital_url,
                    updated_at = excluded.updated_at
            """, (str(hospital_id), hospital_url, d_id, d_name, now))

            for room in dept.get("rooms", []):
                r_id = str(room["room_id"])
                r_name = str(room["room_name"])
                cursor.execute("""
                    INSERT INTO rooms (hospital_id, dept_id, room_id, room_name, updated_at)
                    VALUES (?, ?, ?, ?, ?)
                    ON CONFLICT(hospital_id, dept_id, room_id) DO UPDATE SET
                        room_name = excluded.room_name,
                        updated_at = excluded.updated_at
                """, (str(hospital_id), d_id, r_id, r_name, now))
                total_rooms += 1

        conn.commit()
        conn.close()
        return total_rooms

    def get_departments(self, hospital_id: str = None) -> list:
        """Lấy danh sách khoa theo bệnh viện hoặc tất cả."""
        conn = self.get_connection()
        cursor = conn.cursor()
        if hospital_id:
            cursor.execute("SELECT * FROM departments WHERE hospital_id = ? ORDER BY dept_name", (str(hospital_id),))
        else:
            cursor.execute("SELECT * FROM departments ORDER BY dept_name")
        rows = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return rows

    def get_rooms_by_dept(self, dept_id: str, hospital_id: str = None) -> list:
        """Lấy danh sách phòng thuộc một khoa."""
        conn = self.get_connection()
        cursor = conn.cursor()
        if hospital_id:
            cursor.execute(
                "SELECT * FROM rooms WHERE dept_id = ? AND hospital_id = ? ORDER BY room_name",
                (str(dept_id), str(hospital_id))
            )
        else:
            cursor.execute("SELECT * FROM rooms WHERE dept_id = ? ORDER BY room_name", (str(dept_id),))
        rows = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return rows

    def has_hospital_departments(self, hospital_id: str) -> bool:
        """Kiểm tra bệnh viện đã có dữ liệu khoa trong DB chưa."""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT 1 FROM departments WHERE hospital_id = ? LIMIT 1", (str(hospital_id),))
        row = cursor.fetchone()
        conn.close()
        return row is not None

    def get_full_departments_tree(self, hospital_id: str) -> list:
        """Lấy toàn bộ danh sách khoa kèm danh sách phòng tương ứng từ DB."""
        depts = self.get_departments(hospital_id)
        result = []
        for d in depts:
            rooms = self.get_rooms_by_dept(d["dept_id"], hospital_id)
            result.append({
                "dept_id": d["dept_id"],
                "dept_name": d["dept_name"],
                "rooms_count": len(rooms),
                "rooms": [
                    {"room_id": r["room_id"], "room_name": r["room_name"]}
                    for r in rooms
                ]
            })
        return result


