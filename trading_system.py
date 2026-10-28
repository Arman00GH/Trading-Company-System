import hashlib
import openpyxl
from openpyxl.chart import LineChart, Reference
from colorama import init , Fore , Style
import json
import mysql.connector
import re
import os
from datetime import datetime, timezone
from pymongo import MongoClient, ASCENDING
from pymongo.errors import PyMongoError
from dotenv import load_dotenv


load_dotenv()
init(autoreset=True)

ROLE_USER = 0
ROLE_ADMIN = 1
ROLE_ACCOUNTANT = 2
ROLE_CEO = 3
ROLE_NAMES = {ROLE_USER: "User", ROLE_ADMIN: "Admin", ROLE_ACCOUNTANT: "Accountant", ROLE_CEO: "CEO"}
ALL_ROLES = tuple(ROLE_NAMES)
ROLE_CHOICES_TEXT = ", ".join(f"{code}={name}" for code, name in ROLE_NAMES.items())   # "0=User, 1=Admin, ..."

# Who may use which menu entries. Shared by show_menu() and menu_choice() so they cannot drift apart.
BALANCE_ROLES = (ROLE_ADMIN, ROLE_ACCOUNTANT, ROLE_CEO)
LOW_STOCK_ROLES = (ROLE_ADMIN, ROLE_CEO)

CONTRACT_TYPE_BUY = 0
CONTRACT_TYPE_SELL = 1

CONTRACT_STATUS_PENDING = 0
CONTRACT_STATUS_APPROVED = 1
CONTRACT_STATUS_REJECTED = 2
CONTRACT_STATUS_NAMES = ("Pending", "Approved", "Rejected")   # index == status code

DEFAULT_CEO_PHONE = os.getenv("DEFAULT_CEO_PHONE", "09123456789")
DEFAULT_CEO_PASSWORD = os.getenv("DEFAULT_CEO_PASSWORD", "")
DEFAULT_CEO_AGE = int(os.getenv("DEFAULT_CEO_AGE", "40"))

LOW_STOCK_THRESHOLD = 5
STARTING_WALLET_BALANCE = 20_000_000
STARTING_COMPANY_BALANCE = 100_000_000
MIN_PASSWORD_LENGTH = 4

CONTRACT_FEE = 5_000
FREE_CONTRACTS = 5           
PROFIT_LOSS_WINDOW = 10       # profit/loss and the chart look at the last N approved contracts
CHART_BAR_WIDTH = 50          # characters in the longest text-chart bar
MIN_AGE = 10
MAX_AGE = 120

PHONE_LENGTH = 11
PHONE_PREFIX = "09"

MIN_RATING = 1
MAX_RATING = 5

MONGODB_TIMEOUT_MS = 3_000

# Allowed values
PRODUCT_TYPES = ("Phone", "Camera", "Laptop")
VIDEO_RESOLUTIONS = ("1080p", "4K", "6K", "8K")
SENSOR_TYPES = ("Full Frame", "APS-C", "Micro Four Thirds", "Medium Format")
SQL_IDENTIFIER = re.compile(r"[A-Za-z0-9_$-]{1,64}")   # safe database

USERS_FILE_JSON = "Users.json"
PRODUCTS_FILE_JSON = "Products.json"
CONTRACTS_FILE_JSON = "contracts.json"
BALANCE_FILE_JSON = "company_balance.json"
BALANCE_HISTORY_FILE_JSON = "balance_history.json"

PROFIT_LOSS_CHART_FILE = "profit_loss_chart.xlsx"

MONGODB_URI = os.getenv("MONGODB_URI", "mongodb://localhost:27017/")
MONGODB_DATABASE = os.getenv("MONGODB_DATABASE", "trading_system")
# ============================================================
# MY FUNCTIONS
# ============================================================

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()


def is_valid_phone(phone):
    return len(phone) == PHONE_LENGTH and phone.startswith(PHONE_PREFIX) and phone.isdigit()


def role_name(role):
    return ROLE_NAMES.get(role, "Unknown")


def contract_fee(contract_count):
    """Fee charged on a user's next contract, given how many contracts they already have."""
    return CONTRACT_FEE if contract_count >= FREE_CONTRACTS else 0


def password_long_enough(password):
    """True if the password meets the minimum length; otherwise prints why and returns False."""
    if len(password) >= MIN_PASSWORD_LENGTH:
        return True
    print(Fore.RED + f"Password must be at least {MIN_PASSWORD_LENGTH} characters.")
    return False


def _sql_list(values):
    return ",".join(str(v) for v in values)          # 0,1,2,3


def _sql_enum(values):
    return ",".join(f"'{v}'" for v in values)        # 'Phone','Camera','Laptop'


def ask_int(request):
    "Keeps asking until it gets the correct response"
    while True:
        text = input(request)
        try:
            return int(text.strip())
        except ValueError:
            print(Fore.RED + "Please enter a whole number.")


def ask_nonempty(request):
    while True:
        text = input(request).strip()
        if text:
            return text
        print(Fore.RED + "This cannot be empty.")

###########################################
#CLASSES
###########################################

class Database:
    """MySQL persistence layer for the trading-system Python application."""

    def __init__(self, host=None, user=None, password=None, database=None):
        self.host = host 
        self.user = user 
        self.password = password
        self.database = database
        if not SQL_IDENTIFIER.fullmatch(self.database):
            raise ValueError("Invalid database name: use only letters, digits, '_', '$' and '-' (max 64)")
        self.conn = None

    def connect_server(self):
        if self.conn and self.conn.is_connected():
            return
        self.conn = mysql.connector.connect(
            host=self.host,
            user=self.user,
            password=self.password,
        )

    def connect(self):
        if self.conn and self.conn.is_connected():
            return
        self.conn = mysql.connector.connect(
            host=self.host,
            user=self.user,
            password=self.password,
            database=self.database,
        )

    def create_database(self):
        self.connect_server()
        cur = self.conn.cursor()
        cur.execute(
            f"CREATE DATABASE IF NOT EXISTS `{self.database}` "
            "CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
        )
        self.conn.commit()
        cur.close()
        self.conn.close()
        self.conn = None

    def create_tables(self):
        self.create_database()
        self.connect()
        statements = [
            f"""
            CREATE TABLE IF NOT EXISTS users (
                id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
                name VARCHAR(100) NOT NULL,
                phone CHAR({PHONE_LENGTH}) NOT NULL UNIQUE,
                age TINYINT UNSIGNED NOT NULL,
                role TINYINT UNSIGNED NOT NULL DEFAULT {ROLE_USER},
                contract_count INT UNSIGNED NOT NULL DEFAULT 0,
                wallet_balance DECIMAL(18,2) NOT NULL DEFAULT {STARTING_WALLET_BALANCE},
                password_hash CHAR(64) NOT NULL DEFAULT '',
                CHECK (age BETWEEN {MIN_AGE} AND {MAX_AGE}),
                CHECK (role IN ({_sql_list(ALL_ROLES)})),
                CHECK (wallet_balance >= 0),
                CHECK (contract_count >= 0)
            )
            """,
            f"""
            CREATE TABLE IF NOT EXISTS products (
                id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
                name VARCHAR(150) NOT NULL,
                category ENUM({_sql_enum(PRODUCT_TYPES)}) NOT NULL,
                price DECIMAL(18,2) NOT NULL,
                quantity INT UNSIGNED NOT NULL DEFAULT 0,
                pending_listing BOOLEAN NOT NULL DEFAULT FALSE,
                CHECK (price >= 0)
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS phone_details (
                product_id INT UNSIGNED PRIMARY KEY,
                storage INT UNSIGNED NOT NULL,
                battery_capacity INT UNSIGNED NOT NULL,
                camera_mp DECIMAL(6,2) NOT NULL,
                ram_size INT UNSIGNED NOT NULL,
                FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE CASCADE,
                CHECK (storage > 0 AND battery_capacity > 0 AND camera_mp > 0 AND ram_size > 0)
            )
            """,
            f"""
            CREATE TABLE IF NOT EXISTS camera_details (
                product_id INT UNSIGNED PRIMARY KEY,
                megapixels DECIMAL(6,2) NOT NULL,
                sensor_type VARCHAR(50) NOT NULL,
                lens_mount VARCHAR(100) NOT NULL,
                video_resolution ENUM({_sql_enum(VIDEO_RESOLUTIONS)}) NOT NULL,
                FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE CASCADE,
                CHECK (megapixels > 0)
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS laptop_details (
                product_id INT UNSIGNED PRIMARY KEY,
                cpu VARCHAR(150) NOT NULL,
                ram_size INT UNSIGNED NOT NULL,
                storage INT UNSIGNED NOT NULL,
                gpu VARCHAR(150) NOT NULL,
                screen_size DECIMAL(5,2) NOT NULL,
                FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE CASCADE,
                CHECK (ram_size > 0 AND storage > 0 AND screen_size > 0)
            )
            """,
            f"""
            CREATE TABLE IF NOT EXISTS contracts (
                id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
                user_id INT UNSIGNED NOT NULL,
                product_id INT UNSIGNED NOT NULL,
                type TINYINT UNSIGNED NOT NULL,
                quantity INT UNSIGNED NOT NULL,
                unit_price DECIMAL(18,2) NOT NULL,
                status TINYINT UNSIGNED NOT NULL DEFAULT 0,
                FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE RESTRICT,
                FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE RESTRICT,
                CHECK (type IN ({CONTRACT_TYPE_BUY},{CONTRACT_TYPE_SELL})),
                CHECK (status IN ({CONTRACT_STATUS_PENDING},{CONTRACT_STATUS_APPROVED},{CONTRACT_STATUS_REJECTED})),
                CHECK (quantity > 0),
                CHECK (unit_price > 0)
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS company_balance (
                id TINYINT UNSIGNED PRIMARY KEY,
                balance DECIMAL(18,2) NOT NULL,
                CHECK (balance >= 0)
            )
            """,

            #Timestamp = saving in UTC
            """
            CREATE TABLE IF NOT EXISTS balance_history (
                id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
                balance DECIMAL(18,2) NOT NULL,
                created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                CHECK (balance >= 0)
            )
            """,
        ]

        cur = self.conn.cursor()
        try:
            for statement in statements:
                cur.execute(statement)

            self._ensure_auto_increment_ids(cur)
            self._ensure_pending_listing_column(cur)

            cur.execute(
                f"""
                INSERT INTO company_balance (id, balance)
                VALUES (1, {STARTING_COMPANY_BALANCE})
                ON DUPLICATE KEY UPDATE id = id
                """
            )
            cur.execute(
                "INSERT INTO balance_history (balance) "
                "SELECT balance FROM company_balance WHERE id = 1 "
                "AND NOT EXISTS (SELECT 1 FROM balance_history)"
            )
            self.conn.commit()
        finally:
            cur.close()


    """
    اسم همان کلیدها را از information_schema می‌خواند.
    موقتاً DROP FOREIGN KEY می‌کند.
    روی هر سه جدول id را AUTO_INCREMENT می‌کند.
    همان پنج قید را با اسم‌های ثابت (fk_phone_product و بقیه) برمی‌گرداند.
    """


    def _ensure_auto_increment_ids(self, cur):
        """Upgrade databases made by older versions without breaking FKs."""
        cur.execute(
            "SELECT table_name FROM information_schema.columns "
            "WHERE table_schema = %s AND column_name = 'id' "
            "AND table_name IN ('users', 'products', 'contracts') "
            "AND extra LIKE '%%auto_increment%%'",
            (self.database,),
        )
        if len(cur.fetchall()) == 3:
            return

        # MySQL does not allow a referenced primary key to be altered while a
        # foreign key points to it.  Drop only this application's five known
        # constraints, perform the one-time upgrade, then restore them.
        cur.execute(
            "SELECT table_name, constraint_name FROM information_schema.key_column_usage "
            "WHERE constraint_schema = %s AND table_name IN "
            "('phone_details', 'camera_details', 'laptop_details', 'contracts') "
            "AND referenced_table_name IN ('users', 'products')",
            (self.database,),
        )
        foreign_keys = cur.fetchall()
        for table, constraint in foreign_keys:
            if SQL_IDENTIFIER.fullmatch(table) and SQL_IDENTIFIER.fullmatch(constraint):
                cur.execute(f"ALTER TABLE `{table}` DROP FOREIGN KEY `{constraint}`")

        for table in ("users", "products", "contracts"):
            cur.execute(f"ALTER TABLE `{table}` MODIFY `id` INT UNSIGNED NOT NULL AUTO_INCREMENT")

        cur.execute("ALTER TABLE phone_details ADD CONSTRAINT fk_phone_product FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE CASCADE")
        cur.execute("ALTER TABLE camera_details ADD CONSTRAINT fk_camera_product FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE CASCADE")
        cur.execute("ALTER TABLE laptop_details ADD CONSTRAINT fk_laptop_product FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE CASCADE")
        cur.execute("ALTER TABLE contracts ADD CONSTRAINT fk_contract_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE RESTRICT")
        cur.execute("ALTER TABLE contracts ADD CONSTRAINT fk_contract_product FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE RESTRICT")
    
    
    """
    DO not show the pending objects
    """
    def _ensure_pending_listing_column(self, cur):
        """Hide products introduced by an unapproved sell request."""
        cur.execute(
            "SELECT 1 FROM information_schema.columns "
            "WHERE table_schema = %s AND table_name = 'products' "
            "AND column_name = 'pending_listing'",
            (self.database,),
        )
        if not cur.fetchone():
            cur.execute(
                "ALTER TABLE products ADD COLUMN pending_listing "
                "BOOLEAN NOT NULL DEFAULT FALSE"
            )

    @staticmethod
    def _number(value):
        """Convert MySQL Decimal values to int/float accepted by the existing classes."""
        if value is None:
            return None
        if hasattr(value, "as_tuple"):
            if value == int(value):
                return int(value)
            return float(value)
        return value

    def save_state(self, users, market):
        self.connect()
        cur = self.conn.cursor()

        try:
            if self.conn.in_transaction:
                self.conn.rollback()
            self.conn.start_transaction()

            # ==========================================================
            # USERS
            # ==========================================================
            for u in users:
                cur.execute(
                    """
                    SELECT id
                    FROM users
                    WHERE id = %s
                    """,
                    (u.id,)
                )

                exists = cur.fetchone()

                if exists:
                    cur.execute(
                        """
                        UPDATE users
                        SET
                            name = %s,
                            phone = %s,
                            age = %s,
                            role = %s,
                            contract_count = %s,
                            wallet_balance = %s,
                            password_hash = %s
                        WHERE id = %s
                        """,
                        (
                            u.name,
                            u.phone,
                            u.age,
                            u.role,
                            u.contract_count,
                            u.wallet_balance,
                            u.password_hash,
                            u.id
                        )
                    )
                else:
                    cur.execute(
                        """
                        INSERT INTO users
                        (
                            id,
                            name,
                            phone,
                            age,
                            role,
                            contract_count,
                            wallet_balance,
                            password_hash
                        )
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                        """,
                        (
                            u.id,
                            u.name,
                            u.phone,
                            u.age,
                            u.role,
                            u.contract_count,
                            u.wallet_balance,
                            u.password_hash
                        )
                    )

            # ==========================================================
            # PRODUCTS + PHONE/CAMERA/LAPTOP DETAILS
            # ==========================================================
            for p in market.products:
                cur.execute(
                    """
                    UPDATE products
                    SET name = %s,
                        category = %s,
                        price = %s,
                        quantity = %s
                    WHERE id = %s
                    """,
                    (
                        p.name,
                        p.category,
                        p.price,
                        p.quantity,
                        p.id
                    )
                )

                if cur.rowcount == 0:
                    cur.execute(
                        """
                        INSERT INTO products
                        (id, name, category, price, quantity)
                        VALUES (%s,%s,%s,%s,%s)
                        """,
                        (
                            p.id,
                            p.name,
                            p.category,
                            p.price,
                            p.quantity
                        )
                    )

                category = p.category.lower()

                if category == "phone":
                    cur.execute(
                        """
                        UPDATE phone_details
                        SET storage = %s,
                            battery_capacity = %s,
                            camera_mp = %s,
                            ram_size = %s
                        WHERE product_id = %s
                        """,
                        (
                            p.storage,
                            p.battery_capacity,
                            p.camera_MP,
                            p.ram_size,
                            p.id
                        )
                    )

                    if cur.rowcount == 0:
                        cur.execute(
                            """
                            INSERT INTO phone_details
                            (product_id, storage, battery_capacity,
                             camera_mp, ram_size)
                            VALUES (%s,%s,%s,%s,%s)
                            """,
                            (
                                p.id,
                                p.storage,
                                p.battery_capacity,
                                p.camera_MP,
                                p.ram_size
                            )
                        )

                elif category == "camera":
                    cur.execute(
                        """
                        UPDATE camera_details
                        SET megapixels = %s,
                            sensor_type = %s,
                            lens_mount = %s,
                            video_resolution = %s
                        WHERE product_id = %s
                        """,
                        (
                            p.megapixels,
                            p.sensor_type,
                            p.lens_mount,
                            p.video_resolution,
                            p.id
                        )
                    )

                    if cur.rowcount == 0:
                        cur.execute(
                            """
                            INSERT INTO camera_details
                            (product_id, megapixels, sensor_type,
                             lens_mount, video_resolution)
                            VALUES (%s,%s,%s,%s,%s)
                            """,
                            (
                                p.id,
                                p.megapixels,
                                p.sensor_type,
                                p.lens_mount,
                                p.video_resolution
                            )
                        )

                elif category == "laptop":
                    cur.execute(
                        """
                        UPDATE laptop_details
                        SET cpu = %s,
                            ram_size = %s,
                            storage = %s,
                            gpu = %s,
                            screen_size = %s
                        WHERE product_id = %s
                        """,
                        (
                            p.cpu,
                            p.ram_size,
                            p.storage,
                            p.gpu,
                            p.Screen_size,
                            p.id
                        )
                    )

                    if cur.rowcount == 0:
                        cur.execute(
                            """
                            INSERT INTO laptop_details
                            (product_id, cpu, ram_size, storage,
                             gpu, screen_size)
                            VALUES (%s,%s,%s,%s,%s,%s)
                            """,
                            (
                                p.id,
                                p.cpu,
                                p.ram_size,
                                p.storage,
                                p.gpu,
                                p.Screen_size
                            )
                        )

            # ==========================================================
            # CONTRACTS
            # ==========================================================
            for c in market.contracts:
                cur.execute(
                    """
                    UPDATE contracts
                    SET user_id = %s,
                        product_id = %s,
                        type = %s,
                        quantity = %s,
                        unit_price = %s,
                        status = %s
                    WHERE id = %s
                    """,
                    (
                        c.user_id,
                        c.product_id,
                        c.type,
                        c.quantity,
                        c.unit_price,
                        c.status,
                        c.id
                    )
                )

                if cur.rowcount == 0:
                    cur.execute(
                        """
                        INSERT INTO contracts
                        (id, user_id, product_id, type,
                         quantity, unit_price, status)
                        VALUES (%s,%s,%s,%s,%s,%s,%s)
                        """,
                        (
                            c.id,
                            c.user_id,
                            c.product_id,
                            c.type,
                            c.quantity,
                            c.unit_price,
                            c.status
                        )
                    )

            # ==========================================================
            # COMPANY BALANCE
            # ==========================================================
            cur.execute(
                """
                INSERT INTO company_balance (id, balance)
                VALUES (1, %s)
                ON DUPLICATE KEY UPDATE
                    balance = VALUES(balance)
                """,
                (market.company_balance,)
)
            # ==========================================================
            # BALANCE HISTORY
            # ==========================================================
            cur.execute("SELECT COUNT(*) FROM balance_history")
            stored_count = cur.fetchone()[0]

            if stored_count < len(market.balance_history):
                new_values = market.balance_history[stored_count:]

                for value in new_values:
                    cur.execute(
                        """
                        INSERT INTO balance_history (balance)
                        VALUES (%s)
                        """,
                        (value,)
                    )

            self.conn.commit()

        except Exception:
            self.conn.rollback()
            raise

        finally:
            cur.close()

    def fetch_one(self, statement, params=()):
        self.connect()
        cur = self.conn.cursor(dictionary=True)
        try:
            cur.execute(statement, params)
            return cur.fetchone()
        finally:
            cur.close()

    def fetch_all(self, statement, params=()):
        self.connect()
        cur = self.conn.cursor(dictionary=True)
        try:
            cur.execute(statement, params)
            return cur.fetchall()
        finally:
            cur.close()

    def execute(self, statement, params=()):
        self.connect()
        cur = self.conn.cursor()
        try:
            cur.execute(statement, params)
            self.conn.commit()
            return cur.rowcount
        except Exception:
            self.conn.rollback()
            raise
        finally:
            cur.close()

    def close(self):
        if self.conn and self.conn.is_connected():
            self.conn.close()
            self.conn = None

    
#=========================================================================================================================================================================

class Product:
    next_id = 101
    def __init__(self , name , category ,price, quantity, product_id = None ):
        self._price = None
        self._quantity = None
        if product_id is None:
            self._id = Product.next_id
            Product.next_id += 1
        else:
            self._id = product_id
        self.name = name
        self.category = category
        self.price = price
        self.quantity = quantity

    @property
    def id(self):
        return self._id

    @property
    def name(self):
        return self._name
    @name.setter
    def name(self, value):
        if not isinstance(value , str):
            raise TypeError("Name must be a string")
        value = value.strip()
        if not value:
            raise ValueError("Name can't be empty")
        self._name = value

    @property
    def category(self):
        return self._category
    @category.setter
    def category(self,value):
        if not isinstance(value , str):
            raise TypeError("Category must be a string")
        if not value or value.strip() == "":
            raise ValueError("Category cannot be empty")
        if not value.strip().replace(" " , "").isalpha():
            raise ValueError("Category must contain only letters")
        if value.strip().lower() not in [t.lower() for t in PRODUCT_TYPES]:
            raise TypeError("Category must be laptop , phone or camera")
        self._category = value

    @property
    def price(self):
        return self._price
    @price.setter
    def price(self, value):
        if not isinstance(value , (int , float)):
            raise TypeError("Price must be a number")
        if value < 0:
            raise ValueError("Price can't be negative")
        self._price = value

    @property
    def quantity(self):
        return self._quantity
    @quantity.setter
    def quantity(self, value):
        if not isinstance(value , (int)):
            raise TypeError("Quantity must be an integer")
        if value < 0:
            raise ValueError("Quantity cannot be negative")
        self._quantity = value





#############################################################################################################################################



class Phone(Product):
    def __init__(self, name, category, price, quantity, product_id , storage , battery_capacity , camera_MP , ram_size ):
        super().__init__(name, category, price, quantity, product_id)
        self._storage = storage
        self._battery_capacity = battery_capacity
        self._camera_MP = camera_MP
        self._ram_size = ram_size

    @property
    def storage(self):
        return self._storage
    @storage.setter
    def storage(self,value):
        if not isinstance(value , int) or value <=0:
            raise TypeError("Storage must be positive integer")
        self._storage = value

    @property
    def battery_capacity(self):
        return self._battery_capacity
    @battery_capacity.setter
    def battery_capacity(self,value):
        if not isinstance(value , int) or value <=0:
            raise TypeError("Battery capacity must be positive integer")
        self._battery_capacity = value

    @property
    def camera_MP(self):
        return self._camera_MP
    @camera_MP.setter
    def camera_MP(self,value):
        if not isinstance(value , (int, float)) or value <=0:
            raise TypeError("Camera resolution must be a positive number")
        self._camera_MP = value

    @property
    def ram_size(self):
        return self._ram_size
    @ram_size.setter
    def ram_size(self,value):
        if not isinstance(value , int) or value <=0:
            raise TypeError("RAM size must be positive integer")
        self._ram_size = value




#############################################################################################################################################





class Camera(Product):
    def __init__(self, name, category, price, quantity, product_id , megapixels , sensor_type , lens_mount , video_resolution):
        super().__init__(name, category, price, quantity, product_id)
        self._megapixels = megapixels
        self._sensor_type = sensor_type
        self._lens_mount = lens_mount
        self._video_resolution = video_resolution

    @property
    def megapixels(self):
        return self._megapixels
    @megapixels.setter
    def megapixels(self,value):
        if not isinstance(value , (int , float)) or value <=0:
            raise TypeError("Megapixels must be positive")
        self._megapixels = value

    @property
    def sensor_type(self):
        return self._sensor_type
    @sensor_type.setter
    def sensor_type(self,value):
        if not isinstance(value , str):
            raise TypeError("Sensor type must be a word.")
        valid_types = [t.lower() for t in SENSOR_TYPES]
        if value.lower() not in valid_types:
            raise ValueError("Invalid sensor type")
        self._sensor_type = value

    @property
    def lens_mount(self):
        return self._lens_mount
    @lens_mount.setter
    def lens_mount(self,value):
        if not isinstance(value , str) or not value.strip():
            raise TypeError("Lens mount must be a word and cannot be empty.")
        self._lens_mount = value

    @property
    def video_resolution(self):
        return self._video_resolution
    @video_resolution.setter
    def video_resolution(self,value):
        if value not in VIDEO_RESOLUTIONS:
            raise ValueError("Invalid video resolution")
        self._video_resolution = value



#############################################################################################################################################



class Laptop(Product):
    def __init__(self , name , category , price , quantity , product_id , cpu , ram_size , storage , gpu , Screen_size):
        super().__init__(name, category, price, quantity, product_id)
        self._cpu = cpu
        self._ram_size = ram_size
        self._storage = storage
        self._gpu = gpu
        self._Screen_size = Screen_size

    @property
    def cpu(self):
        return self._cpu
    @cpu.setter
    def cpu(self,value):
        if not isinstance(value , str) or not value.strip():
            raise TypeError("CPU must be a word and cannot be empty.")
        self._cpu = value

    @property
    def ram_size(self):
        return self._ram_size
    @ram_size.setter
    def ram_size(self,value):
        if not isinstance(value , int) or value <=0:
            raise TypeError("RAM size must be positive integer")
        self._ram_size = value

    @property
    def storage(self):
        return self._storage
    @storage.setter
    def storage(self,value):
        if not isinstance(value , int) or value <=0:
            raise TypeError("Storage must be positive integer")
        self._storage = value

    @property
    def gpu(self):
        return self._gpu
    @gpu.setter
    def gpu(self,value):
        if not isinstance(value , str) or not value.strip():
            raise TypeError("GPU must be a word and cannot be empty.")
        self._gpu = value

    @property
    def Screen_size(self):
        return self._Screen_size
    @Screen_size.setter
    def Screen_size(self,value):
        if not isinstance(value , (int, float)) or value <=0:
            raise TypeError("Screen size must be a positive number")
        self._Screen_size = value


#############################################################################################################################################


class User:

    next_id = 1

    def __init__(self, name, phone, age, role=ROLE_USER, user_id=None):
        if user_id is None:
            self._id = User.next_id
            User.next_id += 1
        else:
            self._id = user_id

        self.name = name
        self.phone = phone
        self.age = age
        self.role = role
        self._contract_count = 0
        self._wallet_balance = STARTING_WALLET_BALANCE
        self._password_hash = ""  # set right after creation, see register() in LOGIN / REGISTER / ACCOUNT FUNCTIONS

    @property
    def id(self):
        return self._id
    
    @property
    def name(self):
        return self._name
    @name.setter
    def name(self, value):
        if not isinstance(value , str):
            raise TypeError("Name must be a string")
        value = value.strip()
        if not value:
            raise ValueError("Name can't be empty")
        self._name = value

    @property
    def phone(self):
        return self._phone
    @phone.setter
    def phone(self,value):
        if not is_valid_phone(value):
            raise ValueError(f"Phone must be {PHONE_LENGTH} digits, starting with {PHONE_PREFIX}")
        self._phone = value

    @property
    def age(self):
        return self._age
    @age.setter
    def age(self,value):
        if not isinstance(value , int):
            raise TypeError("Age must be an integer")
        
        if value < MIN_AGE or value > MAX_AGE:
            raise ValueError(f"Age must be between {MIN_AGE} and {MAX_AGE}")
        self._age = value

    @property
    def role(self):
        return self._role
    @role.setter
    def role (self,value):
        if value not in ALL_ROLES:
            raise ValueError("Invalid role")
        self._role = value

    @property
    def wallet_balance(self):
        return self._wallet_balance
    @wallet_balance.setter
    def wallet_balance(self, value):
        if not isinstance(value , (int , float)):
            raise TypeError("Wallet balance must be a number")
        if value <0:
            raise ValueError("Wallet balance can't be negative")
        self._wallet_balance = value

    @property
    def contract_count(self):
        return self._contract_count
    @contract_count.setter
    def contract_count(self, value):
        if not isinstance(value, int):
            raise TypeError("Contract count must be an integer")
        if value < 0:
            raise ValueError("Contract count cannot be negative")
        self._contract_count = value

    @property
    def password_hash(self):
        return self._password_hash
    @password_hash.setter
    def password_hash(self, value):
        if not isinstance(value, str):
            raise TypeError("Password hash must be a string")
        self._password_hash = value

    def print_info(self):
        print(f"ID: {self.id} | Name: {self.name} | Phone: {self.phone} | "
              f"Role: {role_name(self.role)} | Contracts: {self._contract_count} | "
              f"Wallet: {self._wallet_balance} Toman")


#========================================================================================================================================================================================


class Contract:
    """A pending/approved/rejected buy or sell request between a User and the Market."""

    next_id = 1001

    def __init__(self, user_id, product_id, contract_type, quantity, unit_price,
                 contract_id=None, status=CONTRACT_STATUS_PENDING):
        self._user_id = user_id
        self._product_id = product_id
        self.type = contract_type
        self.quantity = quantity
        self.unit_price = unit_price
        self.status = status

        if contract_id is None:
            self._id = Contract.next_id
            Contract.next_id += 1
        else:
            self._id = contract_id
    @property
    def id(self):
        return self._id
    
    @property
    def user_id(self):
        return self._user_id

    @property
    def product_id(self):
        return self._product_id

    @property
    def type(self):
        return self._type
    @type.setter
    def type(self,value):
        if value not in (CONTRACT_TYPE_BUY, CONTRACT_TYPE_SELL):
            raise ValueError("Invalid contract type")
        self._type = value

    @property
    def quantity(self):
        return self._quantity
    @quantity.setter
    def quantity(self,value):
        if value <= 0:
            raise ValueError("Quantity must be positive")
        self._quantity = value

    @property
    def unit_price(self):
        return self._unit_price
    @unit_price.setter
    def unit_price(self,value):
        if value <= 0:
            raise ValueError("Price must be positive")
        self._unit_price = value

    @property
    def status(self):
        return self._status

    @status.setter
    def status(self, value):
        if value not in (CONTRACT_STATUS_PENDING,CONTRACT_STATUS_APPROVED,CONTRACT_STATUS_REJECTED):
            raise ValueError("Invalid status")
        self._status = value





#==========================================================================================================================================================================================


class LegacyJsonMarket:
    """Temporary in-memory model used only to import pre-MySQL JSON files.

    The interactive application uses QueryMarket below, which reads and writes
    MySQL for every operation.  This class must not be used by menus.
    """
    def __init__(self):
        self.products = []
        self.contracts = []
        self.company_balance = STARTING_COMPANY_BALANCE
        self.balance_history = [self.company_balance]

    def export_profit_loss_chart_excel(self, values, start, base):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Profit-Loss"

        sheet.append(["Approval #", "Company Balance (Toman)", "Change From Base (%)"])
        for i, val in enumerate(values, start=start):
            percent = None if base == 0 else round((val - base) / base * 100, 2)
            sheet.append([i, val, percent])

        chart = LineChart()
        chart.title = "Profit / Loss Over Last Approvals"
        chart.x_axis.title = "Approval #"
        chart.y_axis.title = "Company Balance (Toman)"
        chart.style = 12

        last_row = sheet.max_row
        data_ref = Reference(sheet, min_col=2, min_row=1, max_row=last_row)
        cats_ref = Reference(sheet, min_col=1, min_row=2, max_row=last_row)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        sheet.add_chart(chart, "E2")

        workbook.save(PROFIT_LOSS_CHART_FILE)
        print(Fore.GREEN + f"Excel chart saved to {PROFIT_LOSS_CHART_FILE}")
#===============================================================
# LOAD USERS FROM JSON (one-time migration to MySQL)
#===============================================================
    def load_users(self):
        users = []
        try:
            with open(USERS_FILE_JSON, "r", encoding="utf-8") as f:
                rows = json.load(f)
                if not isinstance(rows , list):
                    return
        except (FileNotFoundError, json.JSONDecodeError):
            return users

        for row in rows:
            try:
                user_id = int(row["id"])
                name, phone = row["name"], row["phone"]
                age = int(row["age"])
                role = int(row["role"])
                user = User(name, phone, age, role, user_id=user_id)
                user.contract_count = int(row["contractCount"])
                user.wallet_balance = int(row["walletBalance"])
                user.password_hash = row.get("passwordHash", "")
                users.append(user)
            except (KeyError, ValueError, TypeError):
                continue #instead of crashing

        if users:
            User.next_id = max(u.id for u in users) + 1
        return users
#===============================================================
# LOAD PRODUCTS FROM JSON (one-time migration to MySQL)
#===============================================================
    def load_products(self):
        try:
            with open(PRODUCTS_FILE_JSON, "r", encoding="utf-8") as f:
                rows = json.load(f)
                if not isinstance(rows , list):
                    return
        except (FileNotFoundError, json.JSONDecodeError):
            return

        for row in rows:
            try:
                product_id = int(row["id"])
                name, category = row["name"], row["category"]
                price, quantity = int(row["price"]), int(row["quantity"])
                category_key = category.strip().lower()

                if category_key == "phone":
                    product = Phone(name, category, price, quantity, product_id,
                                     storage=int(row["storage"]),
                                     battery_capacity=int(row["batteryCapacity"]),
                                     camera_MP=row["cameraMP"],
                                     ram_size=int(row["ramSize"]))
                elif category_key == "camera":
                    product = Camera(name, category, price, quantity, product_id,
                                      megapixels=row["megapixels"],
                                      sensor_type=row["sensorType"],
                                      lens_mount=row["lensMount"],
                                      video_resolution=row["videoResolution"])
                elif category_key == "laptop":
                    product = Laptop(name, category, price, quantity, product_id,
                                      cpu=row["cpu"], ram_size=int(row["ramSize"]),
                                      storage=int(row["storage"]), gpu=row["gpu"],
                                      Screen_size=row["screenSize"])
                else:
                    continue

                self.products.append(product)
            except (KeyError, ValueError, TypeError):
                continue

        if self.products:
            Product.next_id = max(p.id for p in self.products) + 1

    def load_contracts(self):
        try:
            with open(CONTRACTS_FILE_JSON, "r", encoding="utf-8") as f:
                rows = json.load(f)
                if not isinstance(rows , list):
                    return
        except (FileNotFoundError, json.JSONDecodeError):
            return

        for row in rows:
            try:
                contract_id = int(row["id"])
                user_id, product_id = int(row["userId"]), int(row["productId"])
                contract_type = int(row["type"])
                status = int(row["status"])
                quantity = int(row["quantity"])
                total_price = int(row["totalPrice"])
                unit_price = total_price // quantity if quantity else 0
                self.contracts.append(Contract(user_id, product_id, contract_type, quantity,
                                                unit_price, contract_id=contract_id, status=status))
            except (KeyError, ValueError, TypeError):
                continue

        if self.contracts:
            Contract.next_id = max(c.id for c in self.contracts) + 1
# #===============================================================
# # LOAD BALANCE FROM JSON (one-time migration to MySQL)
# #===============================================================
    def load_balance(self):
        try:
            with open(BALANCE_FILE_JSON, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.company_balance = int(data["balance"])
        except (FileNotFoundError, json.JSONDecodeError, KeyError, ValueError):
            self.company_balance = STARTING_COMPANY_BALANCE

    def load_balance_history(self):
        try:
            with open(BALANCE_HISTORY_FILE_JSON, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.balance_history = [int(x) for x in data]
        except (FileNotFoundError, json.JSONDecodeError, ValueError):
            return


# ============================================================
# LOGIN / REGISTER / ACCOUNT FUNCTIONS
# ============================================================

class ReviewStore:
    "MongoDB storage for review text and rating values."

    def __init__(self, uri=MONGODB_URI, database=MONGODB_DATABASE):
        self.client = MongoClient(uri, serverSelectionTimeoutMS=MONGODB_TIMEOUT_MS)
        self.client.admin.command("ping")
        self.reviews = self.client[database]["product_reviews"]
        # A user can revise their review, but cannot create duplicates for one product.
        self.reviews.create_index([("product_id", ASCENDING), ("user_id", ASCENDING)], unique=True)
        self.reviews.create_index([("product_id", ASCENDING), ("created_at", -1)])

    def add_review(self, product_id, user_id, rating, text):
        self.reviews.update_one(
            {"product_id": product_id, "user_id": user_id},
            {"$set": {"rating": rating, "text": text.strip(), "updated_at": datetime.now(timezone.utc)},
             "$setOnInsert": {"created_at": datetime.now(timezone.utc)}},
            upsert=True,
        )

    def summaries(self, product_ids):
        if not product_ids:
            return {}
        rows = self.reviews.aggregate([
            {"$match": {"product_id": {"$in": list(product_ids)}}},
            {"$group": {"_id": "$product_id", "score": {"$avg": "$rating"}, "review_count": {"$sum": 1}}},
        ])
        return {row["_id"]: {"score": row["score"], "review_count": row["review_count"]} for row in rows}

    def product_reviews(self, product_id):
        return list(self.reviews.find(
            {"product_id": product_id}, {"_id": 0, "user_id": 1, "rating": 1, "text": 1, "created_at": 1}
        ).sort("created_at", -1))

    def close(self):
        self.client.close()


class QueryMarket:
    "Database-backed market facade; it does not keep product/contract lists."

    PRODUCT_QUERY = """
        SELECT p.id, p.name, p.category, p.price, p.quantity,
               pd.storage AS phone_storage, pd.battery_capacity, pd.camera_mp, pd.ram_size AS phone_ram,
               cd.megapixels, cd.sensor_type, cd.lens_mount, cd.video_resolution,
               ld.cpu, ld.ram_size AS laptop_ram, ld.storage AS laptop_storage, ld.gpu, ld.screen_size
        FROM products p
        LEFT JOIN phone_details pd ON pd.product_id = p.id
        LEFT JOIN camera_details cd ON cd.product_id = p.id
        LEFT JOIN laptop_details ld ON ld.product_id = p.id
        WHERE p.pending_listing = FALSE
    """

    def __init__(self, db, review_store):
        self.db = db
        self.review_store = review_store

    @staticmethod
    def _number(value):
        return Database._number(value)

    def _print_product(self, p):
        tag = Fore.YELLOW + "  [LOW STOCK]" if 0 < p["quantity"] <= LOW_STOCK_THRESHOLD else ""
        text = (f"{p['id']}. {p['name']} ({p['category']}) - Price: {self._number(p['price'])} Toman - "
                f"Stock: {p['quantity']}{tag}")
        category = p["category"].lower()
        if category == "phone":
            text += (f" - Storage: {p['phone_storage']}GB - Battery Capacity: {p['battery_capacity']}mAh"
                     f" - Camera Megapixels: {self._number(p['camera_mp'])} - Ram Size: {p['phone_ram']}GB")
        elif category == "camera":
            text += (f" - Megapixels: {self._number(p['megapixels'])} - Sensor: {p['sensor_type']}"
                     f" - Lens: {p['lens_mount']} - Video: {p['video_resolution']}")
        else:
            text += (f" - CPU: {p['cpu']} - RAM: {p['laptop_ram']}GB - Storage: {p['laptop_storage']}GB"
                     f" - GPU: {p['gpu']} - Screen: {self._number(p['screen_size'])} inches")
        summary = p.get("review_summary")
        if summary:
            text += f" - Score: {summary['score']:.1f}/5 ({summary['review_count']} reviews)"
        else:
            text += " - Score: No reviews"
        print(text)

    def _with_review_summaries(self, rows):
        try:
            summaries = self.review_store.summaries(row["id"] for row in rows) if self.review_store else {}
        except PyMongoError as e:
            summaries = {}
            print(Fore.YELLOW + f"Reviews are temporarily unavailable: {e}")
        for row in rows:
            row["review_summary"] = summaries.get(row["id"])
        return rows

    def show_warehouse(self):
        rows = self._with_review_summaries(self.db.fetch_all(self.PRODUCT_QUERY + " ORDER BY p.quantity = 0, p.id"))
        print(Fore.CYAN + "\n=== Warehouse ===")
        if not rows:
            print("Empty.")
        for row in rows:
            self._print_product(row)

    def search_products(self, keyword):
        keyword = keyword.strip()
        rows = self._with_review_summaries(self.db.fetch_all(self.PRODUCT_QUERY + " AND (p.name LIKE %s OR p.category LIKE %s) ORDER BY p.id",
                                 (f"%{keyword}%", f"%{keyword}%")))
        print(Fore.CYAN + f"\n=== Results for '{keyword}' ===")
        if not rows:
            print("No matching products.")
        for row in rows:
            self._print_product(row)

    def show_low_stock_report(self):
        rows = self._with_review_summaries(self.db.fetch_all(self.PRODUCT_QUERY + " AND p.quantity BETWEEN 1 AND %s ORDER BY p.quantity, p.id",
                                 (LOW_STOCK_THRESHOLD,)))
        print(Fore.CYAN + f"\n=== Low Stock Report (<= {LOW_STOCK_THRESHOLD} units) ===")
        if not rows:
            print("Nothing is low on stock.")
        for row in rows:
            self._print_product(row)

    def create_buy_contract(self, user, product_id, qty):
        if qty <= 0:
            print(Fore.RED + "Invalid quantity!")
            return
        product = self.db.fetch_one(
            "SELECT id, price, quantity FROM products "
            "WHERE id = %s AND pending_listing = FALSE",
            (product_id,),
        )
        if not product:
            print(Fore.RED + "Invalid product")
            return
        if product["quantity"] < qty:
            print(Fore.RED + "Insufficient stock!")
            return
        # Do not manufacture IDs with MAX(id) + 1: two users can submit at
        # once and receive the same ID.  contracts.id is AUTO_INCREMENT.
        self.db.execute(
            "INSERT INTO contracts (user_id, product_id, type, quantity, unit_price, status) "
            "VALUES (%s, %s, %s, %s, %s, %s)",
            (user.id, product_id, CONTRACT_TYPE_BUY, qty, product["price"], CONTRACT_STATUS_PENDING),
        )
        print(Fore.GREEN + "Buy contract registered (pending approval).")

    def create_sell_contract(self, user, product_type, name, price, qty, details):
        if product_type not in PRODUCT_TYPES or qty <= 0 or price <= 0:
            print(Fore.RED + "Quantity and price must be positive!")
            return
        name = name.strip()
        if not name:
            print(Fore.RED + "Product name cannot be empty!")
            return

        try:
            self._validate_sell_details(product_type, details)
        except (KeyError, TypeError, ValueError) as e:
            print(Fore.RED + f"Invalid product details: {e}")
            return

        self.db.connect()
        cur = self.db.conn.cursor(dictionary=True)
        try:
            if self.db.conn.in_transaction:
                self.db.conn.rollback()
            self.db.conn.start_transaction()
            # Only the chosen type's detail keys are read; building all three
            # parameter tuples up front raised KeyError for every sell request.
            match_queries = {
                "Phone": (
                    "SELECT p.id FROM products p JOIN phone_details d ON d.product_id = p.id "
                    "WHERE p.category = %s AND LOWER(p.name) = LOWER(%s) AND p.price = %s "
                    "AND d.storage = %s AND d.battery_capacity = %s AND d.camera_mp = %s "
                    "AND d.ram_size = %s FOR UPDATE",
                    ("storage", "battery_capacity", "camera_MP", "ram_size"),
                ),
                "Camera": (
                    "SELECT p.id FROM products p JOIN camera_details d ON d.product_id = p.id "
                    "WHERE p.category = %s AND LOWER(p.name) = LOWER(%s) AND p.price = %s "
                    "AND d.megapixels = %s AND d.sensor_type = %s AND d.lens_mount = %s "
                    "AND d.video_resolution = %s FOR UPDATE",
                    ("megapixels", "sensor_type", "lens_mount", "video_resolution"),
                ),
                "Laptop": (
                    "SELECT p.id FROM products p JOIN laptop_details d ON d.product_id = p.id "
                    "WHERE p.category = %s AND LOWER(p.name) = LOWER(%s) AND p.price = %s "
                    "AND d.cpu = %s AND d.ram_size = %s AND d.storage = %s AND d.gpu = %s "
                    "AND d.screen_size = %s FOR UPDATE",
                    ("cpu", "ram_size", "storage", "gpu", "Screen_size"),
                ),
            }
            query, detail_keys = match_queries[product_type]
            params = (product_type, name, price, *(details[key] for key in detail_keys))
            cur.execute(query, params)
            row = cur.fetchone()
            if row:
                product_id = row["id"]
            else:
                cur.execute("INSERT INTO products (name, category, price, quantity, pending_listing) VALUES (%s, %s, %s, 0, TRUE)",
                            (name, product_type, price))
                product_id = cur.lastrowid
                kind = product_type.lower()
                if kind == "phone":
                    cur.execute("INSERT INTO phone_details VALUES (%s, %s, %s, %s, %s)",
                                (product_id, details["storage"], details["battery_capacity"], details["camera_MP"], details["ram_size"]))
                elif kind == "camera":
                    cur.execute("INSERT INTO camera_details VALUES (%s, %s, %s, %s, %s)",
                                (product_id, details["megapixels"], details["sensor_type"], details["lens_mount"], details["video_resolution"]))
                else:
                    cur.execute("INSERT INTO laptop_details VALUES (%s, %s, %s, %s, %s, %s)",
                                (product_id, details["cpu"], details["ram_size"], details["storage"], details["gpu"], details["Screen_size"]))
            cur.execute("INSERT INTO contracts (user_id, product_id, type, quantity, unit_price, status) VALUES (%s,%s,%s,%s,%s,%s)",
                        (user.id, product_id, CONTRACT_TYPE_SELL, qty, price, CONTRACT_STATUS_PENDING))
            self.db.conn.commit()
            print(Fore.GREEN + "Sell contract registered (pending approval).")
        except Exception as e:
            self.db.conn.rollback()
            print(Fore.RED + f"Couldn't create sell contract: {e}")
        finally:
            cur.close()

    @staticmethod
    def _validate_sell_details(product_type, details):
        """Validate details before beginning a database transaction."""
        positive_numbers = {
            "Phone": ("storage", "battery_capacity", "camera_MP", "ram_size"),
            "Camera": ("megapixels",),
            "Laptop": ("ram_size", "storage", "Screen_size"),
        }
        for field in positive_numbers[product_type]:
            if not isinstance(details[field], (int, float)) or details[field] <= 0:
                raise ValueError(f"{field} must be a positive number")
        for field in {"Phone": (), "Camera": ("sensor_type", "lens_mount", "video_resolution"), "Laptop": ("cpu", "gpu")}[product_type]:
            if not isinstance(details[field], str) or not details[field].strip():
                raise ValueError(f"{field} cannot be empty")
        if product_type == "Camera" and details["video_resolution"] not in VIDEO_RESOLUTIONS:
            raise ValueError(f"video_resolution must be {', '.join(VIDEO_RESOLUTIONS[:-1])}, or {VIDEO_RESOLUTIONS[-1]}")
        if product_type == "Camera" and details["sensor_type"] not in SENSOR_TYPES:
            raise ValueError(f"Sensor type must be {', '.join(SENSOR_TYPES[:-1])}, or {SENSOR_TYPES[-1]}")

    def _show_contracts(self, rows, heading):
        print(Fore.CYAN + heading)
        if not rows:
            print("Nothing found.")
        for c in rows:
            print(f"Contract {c['id']} | User {c['user_id']} | Product {c['product_id']} | "
                  f"{'Buy' if c['type'] == CONTRACT_TYPE_BUY else 'Sell'} | Qty: {c['quantity']} | "
                  f"Total: {self._number(c['unit_price']) * c['quantity']} Toman | Status: {CONTRACT_STATUS_NAMES[c['status']]}")


    def show_products(self):
        rows = self.db.fetch_all(
            "SELECT id, name, category, price, quantity FROM products "
            "WHERE pending_listing = FALSE ORDER BY id"
        )
        print(Fore.CYAN + "\n=== Products ===")
        if not rows:
            print(Fore.YELLOW + "There are no products.")
            return
        for p in rows:
            print(f"{p['id']}. {p['name']} ({p['category']}) - "
                f"Price: {self._number(p['price'])} Toman - Stock: {p['quantity']}")
        

    def show_pending_contracts(self):
        self._show_contracts(self.db.fetch_all("SELECT * FROM contracts WHERE status = %s ORDER BY id", (CONTRACT_STATUS_PENDING,)), "\n=== Pending Contracts ===")

    def show_my_contracts(self, user):
        self._show_contracts(self.db.fetch_all("SELECT * FROM contracts WHERE user_id = %s ORDER BY id", (user.id,)), "\n=== Your Contracts ===")

    def add_review(self, user, product_id, rating, text):
        if not self.review_store:
            print(Fore.YELLOW + "Reviews are temporarily unavailable.")
            return
        if rating not in range(MIN_RATING, MAX_RATING + 1):
            print(Fore.RED + f"Rating must be from {MIN_RATING} to {MAX_RATING}.")
            return
        if not text.strip():
            print(Fore.RED + "Review text cannot be empty.")
            return
        # Only customers with an approved purchase can review a product.
        purchase = self.db.fetch_one(
            "SELECT id FROM contracts WHERE user_id = %s AND product_id = %s AND type = %s AND status = %s LIMIT 1",
            (user.id, product_id, CONTRACT_TYPE_BUY, CONTRACT_STATUS_APPROVED),
        )
        if not purchase:
            print(Fore.RED + "You can review only products you have bought and received.")
            return
        try:
            self.review_store.add_review(product_id, user.id, rating, text)
        except PyMongoError as e:
            print(Fore.YELLOW + f"Reviews are temporarily unavailable: {e}")
            return
        print(Fore.GREEN + "Review saved. Submitting another review updates your previous one.")

    def show_product_reviews(self, product_id):
        if not self.review_store:
            print(Fore.YELLOW + "Reviews are temporarily unavailable.")
            return
        product = self.db.fetch_one("SELECT name FROM products WHERE id = %s", (product_id,))
        if not product:
            print(Fore.RED + "Invalid product.")
            return
        try:
            reviews = self.review_store.product_reviews(product_id)
            summary = self.review_store.summaries([product_id]).get(product_id)
        except PyMongoError as e:
            print(Fore.YELLOW + f"Reviews are temporarily unavailable: {e}")
            return
        print(Fore.CYAN + f"\n=== Reviews: {product['name']} ===")
        if not reviews:
            print("No reviews yet.")
            return
        print(f"Score: {summary['score']:.1f}/5 from {summary['review_count']} review(s)")
        for review in reviews:
            created = review["created_at"].astimezone().strftime("%Y-%m-%d")
            print(f"{review['rating']}/5 | User {review['user_id']} | {created}\n  {review['text']}")

    def approve_contract(self, contract_id, approve, current_user=None):
        self.db.connect()
        cur = self.db.conn.cursor(dictionary=True)

        def reject(message):
            cur.execute("UPDATE contracts SET status = %s WHERE id = %s", (CONTRACT_STATUS_REJECTED, contract_id))
            self.db.conn.commit()
            print(Fore.RED + message)

        try:
            if self.db.conn.in_transaction:
                self.db.conn.rollback()
            self.db.conn.start_transaction()
            cur.execute("SELECT * FROM contracts WHERE id = %s AND status = %s FOR UPDATE", (contract_id, CONTRACT_STATUS_PENDING))
            c = cur.fetchone()
            if not c:
                print(Fore.RED + "Contract not found or already processed.")
                self.db.conn.rollback()
                return
            if not approve:
                reject("Contract rejected.")
                return
            cur.execute("SELECT quantity FROM products WHERE id = %s FOR UPDATE", (c["product_id"],))
            product = cur.fetchone()
            cur.execute("SELECT wallet_balance, contract_count FROM users WHERE id = %s FOR UPDATE", (c["user_id"],))
            user = cur.fetchone()
            cur.execute("SELECT balance FROM company_balance WHERE id = 1 FOR UPDATE")
            company = cur.fetchone()
            if not product or not user or not company:
                reject("Related product or user not found.")
                return
            total = self._number(c["unit_price"]) * c["quantity"]
            fee = contract_fee(user["contract_count"])
            wallet = self._number(user["wallet_balance"])
            balance =  self._number(company["balance"])
            if (c["type"] == CONTRACT_TYPE_BUY and (wallet < fee + total or product["quantity"] < c["quantity"])) or (c["type"] == CONTRACT_TYPE_SELL and (wallet < fee or balance < total)):
                reject("Approval failed: insufficient funds or stock.")
                return
            if c["type"] == CONTRACT_TYPE_BUY:
                new_wallet = wallet - fee - total
                new_qty = product["quantity"] - c["quantity"]
                new_balance = balance + total + fee
            else:
                new_wallet = wallet - fee + total
                new_qty = product["quantity"] + c["quantity"]
                new_balance = balance - total + fee

            cur.execute("UPDATE users SET wallet_balance = %s, contract_count = contract_count + 1 WHERE id = %s", (new_wallet, c["user_id"]))
            cur.execute(
                "UPDATE products SET quantity = %s, pending_listing = FALSE WHERE id = %s",
                (new_qty, c["product_id"]),
            )
            cur.execute("UPDATE company_balance SET balance = %s WHERE id = 1", (new_balance,))
            cur.execute("INSERT INTO balance_history (balance) VALUES (%s)", (new_balance,))
            cur.execute("UPDATE contracts SET status = %s WHERE id = %s", (CONTRACT_STATUS_APPROVED, contract_id))
            self.db.conn.commit()
            if current_user and current_user.id == c["user_id"]:
                current_user.wallet_balance, current_user.contract_count = new_wallet, user["contract_count"] + 1
            print(Fore.GREEN + "Contract approved and applied.")
        except Exception:
            self.db.conn.rollback()
            raise

        finally:
            cur.close()

    def show_company_balance(self):
        row = self.db.fetch_one("SELECT balance FROM company_balance WHERE id = 1")
        print(f"Company Balance: {self._number(row['balance'])} Toman")

    def _recent_balances(self, count=PROFIT_LOSS_WINDOW + 1):   # N approvals need N+1 balances
        return [self._number(r["balance"]) for r in self.db.fetch_all("SELECT balance FROM balance_history ORDER BY id DESC LIMIT %s", (count,))][::-1]

    def show_profit_loss(self):
        values = self._recent_balances()
        if len(values) < PROFIT_LOSS_WINDOW + 1 or values[0] == 0:
            print(Fore.YELLOW + f"Not enough data yet (need at least {PROFIT_LOSS_WINDOW} approved contracts).")
            return
        percent = (values[-1] - values[0]) / values[0] * 100
        print((Fore.GREEN if percent >= 0 else Fore.RED) + f"Profit/Loss (last {PROFIT_LOSS_WINDOW} contracts): {percent:.2f}%")

    def show_profit_loss_chart(self):
        values = self._recent_balances()
        if len(values) < 2:
            print(Fore.YELLOW + "Not enough data yet."); return
        base, low, high = values[0], min(values), max(values)
        print(Fore.CYAN + f"\n=== Profit/Loss Chart (Last {len(values) - 1} Approvals) ===")
        for i, value in enumerate(values):
            bar = "#" * ((value - low) * CHART_BAR_WIDTH // max(high - low, 1))
            print(f"C{i:>3}: {value:>12} | {bar} {(value - base) / base * 100 if base else 0:.1f}%")
        LegacyJsonMarket.export_profit_loss_chart_excel(self, values, 0, base)

def _user_from_row(row):
    if not row:
        return None
    user = User(row["name"], row["phone"], int(row["age"]), int(row["role"]), user_id=int(row["id"]))
    user.contract_count = int(row["contract_count"])
    user.wallet_balance = Database._number(row["wallet_balance"])
    user.password_hash = row["password_hash"] or ""
    return user


def ensure_ceo_exists(db):
    if db.fetch_one("SELECT id FROM users WHERE role = %s LIMIT 1", (ROLE_CEO,)):
        return
    if not DEFAULT_CEO_PASSWORD:
        print(Fore.RED + "No CEO account exists. Set DEFAULT_CEO_PASSWORD in your .env file, then restart.")
        return
    ceo = User("CEO", DEFAULT_CEO_PHONE, DEFAULT_CEO_AGE, ROLE_CEO)
    ceo.password_hash = hash_password(DEFAULT_CEO_PASSWORD)
    db.execute("INSERT INTO users (name, phone, age, role, password_hash) "
               "VALUES (%s, %s, %s, %s, %s)",
               (ceo.name, ceo.phone, ceo.age, ceo.role, ceo.password_hash))
    print(Fore.YELLOW + "No CEO account found -- created one.")
    print(Fore.YELLOW + f"  Phone:    {DEFAULT_CEO_PHONE}")
    print(Fore.YELLOW + "  Log in with the CEO password configured in your local .env file.")


def register(db):
    name = ask_nonempty("Name: ")
    phone = ask_nonempty("Phone: ")
    age = ask_int("Age: ")

    if db.fetch_one("SELECT id FROM users WHERE phone = %s", (phone,)):
        print(Fore.RED + "That phone number is already registered.")
        return

    try:
        user = User(name, phone, age)
    except ValueError as e:
        print(Fore.RED + "Could not register:", e)
        return

    password = input("Choose a password: ")
    confirm = input("Confirm password: ")
    if password != confirm:
        print(Fore.RED + "Passwords do not match.")
        return
    if not password_long_enough(password):
        return
    user.password_hash = hash_password(password)

    db.execute("INSERT INTO users (name, phone, age, role, password_hash) "
               "VALUES (%s, %s, %s, %s, %s)",
               (user.name, user.phone, user.age, user.role, user.password_hash))
    print(Fore.GREEN + "Registered! You can now log in.")

def login(db):
    phone = input("Phone: ")
    user = _user_from_row(db.fetch_one("SELECT * FROM users WHERE phone = %s", (phone,)))

    if user is None:
        print(Fore.RED + "User not found.")
        return None

    if user.password_hash == "":
        print(Fore.YELLOW + "No password set for this account yet -- let's create one.")
        password = input("Choose a password: ")
        if not password_long_enough(password):
            return None
        user.password_hash = hash_password(password)
        db.execute("UPDATE users SET password_hash = %s WHERE id = %s", (user.password_hash, user.id))
    else:
        password = input("Password: ")
        if hash_password(password) != user.password_hash:
            print(Fore.RED + "Invalid phone or password.")
            return None

    print(Fore.GREEN + f"Welcome {user.name} ({role_name(user.role)})")
    return user


def change_password(db, user):
    current = input("Current password: ")
    if hash_password(current) != user.password_hash:
        print(Fore.RED + "Current password is incorrect.")
        return
    new_password = input("New password: ")
    if not password_long_enough(new_password):
        return
    user.password_hash = hash_password(new_password)
    db.execute("UPDATE users SET password_hash = %s WHERE id = %s", (user.password_hash, user.id))
    print(Fore.GREEN + "Password updated.")


def change_user_role(db, current_user):
    print(Fore.CYAN + "\n=== All Users ===")
    for row in db.fetch_all("SELECT * FROM users ORDER BY id"):
        _user_from_row(row).print_info()

    phone = input("\nUser Phone: ")
    new_role = ask_int(f"New Role ({ROLE_CHOICES_TEXT}): ")

    target = _user_from_row(db.fetch_one("SELECT * FROM users WHERE phone = %s", (phone,)))

    if not target:
        print(Fore.RED + "User not found.")
        return
    if target.phone == current_user.phone:
        print(Fore.RED + "CEO cannot change their own role!")
        return
    if new_role not in ALL_ROLES:
        print(Fore.RED + "Invalid role.")
        return

    target.role = new_role
    db.execute("UPDATE users SET role = %s WHERE id = %s", (new_role, target.id))
    print(Fore.GREEN + "Role changed successfully!")


# ============================================================
# MENU
# ============================================================

def show_menu(user):
    print(Fore.CYAN + "\n--- Menu ---")
    print("1. View Warehouse")
    print("2. Search Products")
    print("3. Buy")
    print("4. Sell")
    print("5. My Wallet Balance")
    print("6. My Contracts")
    print("7. Change Password")
    if user.role in BALANCE_ROLES:
        print("8. Company Balance")
    if user.role in LOW_STOCK_ROLES:
        print("9. Low Stock Report")
    if user.role == ROLE_ADMIN:
        print("10. Approve/Reject Contracts")
    if user.role == ROLE_ACCOUNTANT:
        print("9. Calculate Profit/Loss")
        print("10. Show Profit/Loss Chart")
    if user.role == ROLE_CEO:
        print("10. Change User Role")
    print("11. Product Reviews & Scores")
    print("0. Logout")


def menu_choice(choice, user, market, db):
    "Returns False when the user chooses to log out, True otherwise."

    if choice == 0:
        return False

    elif choice == 1:
        market.show_warehouse()

    elif choice == 2:
        keyword = input("Search (name or category): ")
        market.search_products(keyword)

    elif choice == 3:
        market.show_warehouse()
        product_id = ask_int("\nProduct ID: ")
        qty = ask_int("Quantity: ")
        market.create_buy_contract(user, product_id, qty)

    elif choice == 4:
        print(Fore.CYAN + "\nWhat are you selling?")
        for number, type_name in enumerate(PRODUCT_TYPES, start=1):
            print(f"{number}. {type_name}")
        type_choice = ask_int("Choice: ")
        product_type = dict(enumerate(PRODUCT_TYPES, start=1)).get(type_choice) #we use get to recieve None if the choice is invalid instead of raising an exception

        if product_type is None:
            print(Fore.RED + "Invalid choice.")
        else:
            name = ask_nonempty("Name: ")
            price = ask_int("Price: ")
            qty = ask_int("Quantity: ")

            if product_type == "Phone":
                details = {
                    "storage": ask_int("Storage (GB): "),
                    "battery_capacity": ask_int("Battery Capacity (mAh): "),
                    "camera_MP": ask_int("Camera Megapixels: "),
                    "ram_size": ask_int("RAM Size (GB): "),
                }
            elif product_type == "Camera":
                details = {
                    "megapixels": ask_int("Megapixels: "),
                    "sensor_type": ask_nonempty(f"Sensor Type ({'/'.join(SENSOR_TYPES)}): "),
                    "lens_mount": ask_nonempty("Lens Mount: "),
                    "video_resolution": ask_nonempty(f"Video Resolution ({'/'.join(VIDEO_RESOLUTIONS)}): "),
                }
            else:
                details = {
                    "cpu": ask_nonempty("CPU: "),
                    "ram_size": ask_int("RAM Size (GB): "),
                    "storage": ask_int("Storage (GB): "),
                    "gpu": ask_nonempty("GPU: "),
                    "Screen_size": ask_int("Screen Size (inches): "),
                }

            market.create_sell_contract(user, product_type, name, price, qty, details)

    elif choice == 5:
        print(f"Your Wallet Balance: {user.wallet_balance} Toman")

    elif choice == 6:
        market.show_my_contracts(user)

    elif choice == 7:
        change_password(db, user)

    elif choice == 8 and user.role in BALANCE_ROLES:
        market.show_company_balance()

    elif choice == 9 and user.role in LOW_STOCK_ROLES:
        market.show_low_stock_report()

    elif choice == 10 and user.role == ROLE_ADMIN:
        market.show_pending_contracts()
        contract_id = ask_int("\nContract ID: ")
        approve = ask_int("1 = Approve, 0 = Reject: ") == 1
        market.approve_contract(contract_id, approve, user)

    elif choice == 9 and user.role == ROLE_ACCOUNTANT:
        market.show_profit_loss()

    elif choice == 10 and user.role == ROLE_ACCOUNTANT:
        market.show_profit_loss_chart()

    elif choice == 10 and user.role == ROLE_CEO:
        change_user_role(db, user)

    elif choice == 11:
        print("1. Add/update a review   2. View product reviews")
        review_choice = ask_int("Choice: ")
        if review_choice not in (1, 2):
            print(Fore.RED + "Invalid choice.")
        else:
            market.show_products()
            product_id = ask_int("Product ID: ")
            existance = db.fetch_one("SELECT id FROM products WHERE id = %s " ,(product_id,))
            if not existance:
                print(Fore.RED , "Invalid Product")
            elif review_choice == 1:
                rating = ask_int(f"Rating ({MIN_RATING}-{MAX_RATING}): ")
                text = ask_nonempty("Review: ")
                market.add_review(user, product_id, rating, text)
            else:
                market.show_product_reviews(product_id)

    else:
        print(Fore.RED + "Invalid choice.")

    return True

#====================================================================
#  Main
#====================================================================
def main():
    db = Database(
        host=os.getenv("MYSQL_HOST", "127.0.0.1"),
        user=os.getenv("MYSQL_USER", "root"),
        password=os.getenv("MYSQL_PASSWORD", ""),
        database=os.getenv("MYSQL_DATABASE", "trading_system"),
    )

    try:
        db.create_tables()
    except Exception as e:
        print(Fore.RED + "Database initialization/loading failed.")
        print(Fore.RED + f"Error: {e}")
        return

    # One-time migration: if the SQL database is empty but the old JSON
    # files exist, load the old data and immediately copy it into MySQL.
    is_empty = not db.fetch_one("SELECT id FROM users LIMIT 1") and not db.fetch_one("SELECT id FROM products LIMIT 1")
    if is_empty:
        JSON_market = LegacyJsonMarket()
        JSON_users = JSON_market.load_users()
        JSON_market.load_products()
        JSON_market.load_contracts()
        JSON_market.load_balance()
        JSON_market.load_balance_history()

        if JSON_users or JSON_market.products or JSON_market.contracts:
            db.save_state(JSON_users, JSON_market)
            print(Fore.GREEN + "Existing JSON data migrated to MySQL.")

    try:
        review_store = ReviewStore()
    except PyMongoError as e:
        # Reviews are an optional feature; a Mongo outage must not make the
        # database-backed warehouse and contracts unusable.
        review_store = None
        print(Fore.YELLOW + "MongoDB connection failed; reviews are unavailable.")
        print(Fore.YELLOW + f"Error: {e}")

    market = QueryMarket(db, review_store)
    ensure_ceo_exists(db)

    print(Style.BRIGHT + Fore.CYAN + "=== Trading Company System ===")
    current_user = None

    while True:
        if current_user is None:
            print(Fore.CYAN + "\n1. Login   2. Register   3. Exit")
            choice = ask_int("Choice: ")

            if choice == 1:
                current_user = login(db)

            elif choice == 2:
                register(db)

            elif choice == 3:
                break

            else:
                print(Fore.RED + "Invalid choice.")

        else:
            show_menu(current_user)
            choice = ask_int("Choice: ")
            still_logged_in = menu_choice(
                choice, current_user, market, db
            )

            if not still_logged_in:
                current_user = None

    db.close()
    if review_store:
        review_store.close()
    print(Fore.CYAN + "Goodbye!")


if __name__ == "__main__":
    main()
