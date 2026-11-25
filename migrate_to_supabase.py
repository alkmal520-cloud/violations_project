import os
import re
from urllib.parse import quote_plus, unquote
import pandas as pd
from dotenv import load_dotenv
from sqlalchemy import create_engine, inspect, text

load_dotenv(override=True)


def clean_supabase_url(raw_url: str) -> str:
    if not raw_url:
        raise ValueError("DATABASE_URL غير موجود في ملف .env")

    val = raw_url.strip().strip('"').strip("'")
    val = re.split(r"\s*\]\s*\(https?://", val, maxsplit=1)[0]
    val = val.replace("]", "").replace("[", "").strip()

    match = re.search(r"(postgres(?:ql)?(?:\+psycopg2)?://\S+)", val)
    if match:
        val = match.group(1)

    if "://" in val and "@" in val:
        scheme_part, rest = val.split("://", 1)
        creds_part, host_and_db = rest.rsplit("@", 1)
        if ":" in creds_part:
            user, raw_pwd = creds_part.split(":", 1)
            safe_pwd = quote_plus(unquote(raw_pwd))
            val = f"{scheme_part}://{user}:{safe_pwd}@{host_and_db}"

    if val.startswith("postgres://"):
        val = "postgresql+psycopg2://" + val[len("postgres://") :]
    elif val.startswith("postgresql://") and "+psycopg2" not in val:
        val = "postgresql+psycopg2://" + val[len("postgresql://") :]

    if "sslmode=" not in val:
        sep = "&" if "?" in val else "?"
        val = f"{val}{sep}sslmode=require"

    return val


SUPABASE_URL = clean_supabase_url(os.getenv("DATABASE_URL", ""))

LOCAL_USER = os.getenv("DB_USER", "postgres")
LOCAL_PASSWORD = quote_plus(unquote(os.getenv("DB_PASSWORD", "roo71900")))
LOCAL_HOST = os.getenv("DB_HOST", "localhost")
LOCAL_PORT = os.getenv("DB_PORT", "5432")
LOCAL_DB = os.getenv("DB_NAME", "violations_db")

LOCAL_DB_URL = (
    f"postgresql+psycopg2://{LOCAL_USER}:{LOCAL_PASSWORD}@"
    f"{LOCAL_HOST}:{LOCAL_PORT}/{LOCAL_DB}"
)

local_engine = create_engine(LOCAL_DB_URL, pool_pre_ping=True)
supabase_engine = create_engine(SUPABASE_URL, pool_pre_ping=True)

local_inspector = inspect(local_engine)
all_local_tables = local_inspector.get_table_names()

preferred_order = [
    "tbl_cases",
    "tbl_offenders",
    "tbl_locations",
    "tbl_executionstatus",
    "tbl_judgements",
    "tbl_reconciliation",
    "tbl_violationtypes",
    "tbl_Archive",
    "tbl_audit_logs",
]

ordered = [t for t in preferred_order if t in all_local_tables]
others = [t for t in all_local_tables if t not in ordered]
tables_to_migrate = ordered + others

print(f"عدد الجداول الموجودة في violations_db للنقل: {len(tables_to_migrate)}")

for table_name in tables_to_migrate:
    try:
        cols = [c["name"] for c in local_inspector.get_columns(table_name)]
        if not cols:
            print(f"تخطي {table_name}: لا توجد أعمدة.")
            continue

        cols_sql = ", ".join([f'"{c}"' for c in cols])
        with local_engine.connect() as conn:
            total_rows = conn.execute(
                text(f'SELECT COUNT(*) FROM "{table_name}"')
            ).scalar()

        print(f"\n--- جدول: {table_name} (عدد السجلات: {total_rows}) ---")

        if not total_rows:
            # إنشاء جدول فارغ بنفس الأعمدة على supabase إذا لم يكن له سجلات
            empty_df = pd.read_sql(
                text(f'SELECT {cols_sql} FROM "{table_name}" LIMIT 0'),
                local_engine,
            )
            empty_df.to_sql(
                table_name,
                con=supabase_engine,
                if_exists="replace",
                index=False,
            )
            print("تم إنشاء الجدول فارغاً على supabase.")
            continue

        chunk_size = 250
        transferred = 0
        first_chunk = True
        query_sql = text(f'SELECT {cols_sql} FROM "{table_name}"')

        for chunk in pd.read_sql(query_sql, local_engine, chunksize=chunk_size):
            if_exists_mode = "replace" if first_chunk else "append"
            chunk.to_sql(
                table_name,
                con=supabase_engine,
                if_exists=if_exists_mode,
                index=False,
                method="multi",
            )
            first_chunk = False
            transferred += len(chunk)
            print(f"  تم نقل {transferred} / {total_rows}")

        print(f"اكتمل نقل الجدول {table_name} بنجاح.")

    except Exception as e:
        print(f"خطأ في الجدول {table_name}: {e}")

print("\n=== انتهى النقل إلى supabase ===")