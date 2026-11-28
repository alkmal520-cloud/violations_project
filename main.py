import os
import math
import io
import glob
import json
import subprocess
from datetime import datetime, timedelta
from typing import Optional
from fastapi import FastAPI, Depends, HTTPException, Query, Request, Response, status, UploadFile, File
from fastapi.responses import HTMLResponse, StreamingResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from pydantic import BaseModel
from sqlalchemy import create_engine, Column, BigInteger, Text, Boolean, Float, func, or_
from sqlalchemy.orm import declarative_base, sessionmaker, Session
from passlib.context import CryptContext
import jwt
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import pythoncom
    import win32com.client
    WIA_AVAILABLE = True
except Exception:
    pythoncom = None
    win32com = None
    WIA_AVAILABLE = False

# استيرادات توليد ملفات PDF باللغة العربية
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import arabic_reshaper
from bidi.algorithm import get_display

# --- تسجيل خط عربي من النظام لمنع ظاهرة المربعات السوداء ---
ARABIC_FONT_NAME = 'Helvetica'
POSSIBLE_FONT_PATHS = [
    r"C:\Windows\Fonts\arial.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/noto/NotoNaskhArabic-Regular.ttf",
]
for font_path in POSSIBLE_FONT_PATHS:
    if os.path.exists(font_path):
        try:
            pdfmetrics.registerFont(TTFont('ArabicFont', font_path))
            ARABIC_FONT_NAME = 'ArabicFont'
            break
        except Exception:
            pass

def ar_text(text):
    if not text or str(text).strip() in ["", "None", "nan"]:
        return "-"
    reshaped_text = arabic_reshaper.reshape(str(text))
    bidi_text = get_display(reshaped_text)
    return bidi_text

# --- إعدادات الأمان والتشفير العالمي ---
SECRET_KEY = os.getenv("JWT_SECRET_KEY", "SUPER_SECURE_ENTERPRISE_KEY_2026_987654321_ROO7")
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 480

pwd_context = CryptContext(schemes=["pbkdf2_sha256"], deprecated="auto")
security = HTTPBearer()

def verify_password(plain_password, hashed_password):
    return pwd_context.verify(plain_password, hashed_password)

def get_password_hash(password):
    return pwd_context.hash(password)

def create_access_token(data: dict, expires_delta: Optional[timedelta] = None):
    to_encode = data.copy()
    expire = datetime.utcnow() + (expires_delta or timedelta(minutes=15))
    to_encode.update({"exp": expire})
    return jwt.encode(to_encode, SECRET_KEY, algorithm=ALGORITHM)

def decode_access_token(token: str):
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        return payload
    except jwt.ExpiredSignatureError:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="انتهت صلاحية الجلسة الأمنية، يرجى إعادة تسجيل الدخول")
    except jwt.PyJWTError:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="رمز المصادقة غير صالح")

# --- اتصال قاعدة البيانات PostgreSQL (محلي / شبكة محلية / سحابي عبر متغيرات البيئة) ---
POSTGRES_USER = os.getenv("DB_USER", "postgres")
POSTGRES_PASSWORD = os.getenv("DB_PASSWORD", "roo71900")
POSTGRES_HOST = os.getenv("DB_HOST", "localhost")
POSTGRES_PORT = os.getenv("DB_PORT", "5432")
POSTGRES_DB = os.getenv("DB_NAME", "violations_db")

RAW_DATABASE_URL = (os.getenv("DATABASE_URL") or "").strip()
if RAW_DATABASE_URL:
    if RAW_DATABASE_URL.startswith("postgres://"):
        SQLALCHEMY_DATABASE_URL = RAW_DATABASE_URL.replace("postgres://", "postgresql+psycopg2://", 1)
    elif RAW_DATABASE_URL.startswith("postgresql://") and "+psycopg2" not in RAW_DATABASE_URL:
        SQLALCHEMY_DATABASE_URL = RAW_DATABASE_URL.replace("postgresql://", "postgresql+psycopg2://", 1)
    else:
        SQLALCHEMY_DATABASE_URL = RAW_DATABASE_URL
else:
    SQLALCHEMY_DATABASE_URL = f"postgresql+psycopg2://{POSTGRES_USER}:{POSTGRES_PASSWORD}@{POSTGRES_HOST}:{POSTGRES_PORT}/{POSTGRES_DB}"

engine_connect_args = {}
if "sslmode=" not in SQLALCHEMY_DATABASE_URL and ("supabase.co" in SQLALCHEMY_DATABASE_URL or "neon.tech" in SQLALCHEMY_DATABASE_URL or "render.com" in SQLALCHEMY_DATABASE_URL or os.getenv("DB_SSLMODE") == "require"):
    engine_connect_args["sslmode"] = "require"

engine = create_engine(
    SQLALCHEMY_DATABASE_URL,
    connect_args=engine_connect_args,
    pool_pre_ping=True,
    pool_size=10,
    max_overflow=20
)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

from sqlalchemy import inspect as sa_inspect


def _get_table_columns(table_name: str) -> set:
    try:
        inspector = sa_inspect(engine)
        if table_name in inspector.get_table_names():
            return {col["name"] for col in inspector.get_columns(table_name)}
    except Exception:
        pass
    return set()


# --- النماذج والجداول السيادية (متوافقة مع أنواع قاعدة البيانات السحابية والبيانات النصية) ---
class User(Base):
    __tablename__ = "tbl_users"
    UserID = Column(BigInteger, primary_key=True, index=True, autoincrement=True)
    Username = Column(Text, unique=True, index=True, nullable=False)
    PasswordHash = Column(Text, nullable=False)
    FullName = Column(Text, nullable=True)
    Role = Column(Text, nullable=False, default="DataEntry")
    IsActive = Column(Boolean, nullable=True, default=True)
    CreatedAt = Column(Text, nullable=True)


class AuditLog(Base):
    __tablename__ = "tbl_audit_logs"
    LogID = Column(BigInteger, primary_key=True, index=True, autoincrement=True)
    Username = Column(Text, index=True, nullable=True)
    ActionType = Column(Text, index=True, nullable=True)
    TargetTable = Column(Text, nullable=True)
    Details = Column(Text, nullable=True)
    Timestamp = Column(Text, nullable=True)


class Case(Base):
    __tablename__ = "tbl_cases"
    CaseID = Column(BigInteger, primary_key=True, index=True, autoincrement=True)
    SerialNo = Column(Text, nullable=True)
    VariableCode = Column(Text, index=True, nullable=True)
    VariableDate = Column(Text, nullable=True)
    CaseNumber = Column(Text, index=True, nullable=True)
    CaseDate = Column(Text, nullable=True)
    NationalID = Column(Text, nullable=True)
    year = Column(Text, index=True, nullable=True)
    StopNumber = Column(Text, nullable=True)
    AdminNumber = Column(Text, nullable=True)
    DecisionNumber = Column(Text, index=True, nullable=True)
    DecisionDate = Column(Text, nullable=True)
    AreaSize = Column(Text, nullable=True)
    Notes = Column(Text, nullable=True)
    Latitude = Column(Text, nullable=True)
    Longitude = Column(Text, nullable=True)
    OffenderID = Column(BigInteger, nullable=True)
    LocationID = Column(BigInteger, nullable=True)


class Offender(Base):
    __tablename__ = "tbl_offenders"
    OffenderID = Column(BigInteger, primary_key=True, index=True, autoincrement=True)
    CaseID = Column(BigInteger, nullable=True)
    OffenderName = Column(Text, index=True, nullable=True)
    NationalID = Column(Text, index=True, nullable=True)
    OffenderAddress = Column(Text, nullable=True)
    Phone = Column(Text, nullable=True)
    Notes = Column(Text, nullable=True)


class Location(Base):
    __tablename__ = "tbl_locations"
    LocationID = Column(BigInteger, primary_key=True, index=True, autoincrement=True)
    CaseID = Column(BigInteger, nullable=True)
    StreetName = Column(Text, index=True, nullable=True)
    DistrictName = Column(Text, index=True, nullable=True)
    FullAddress = Column(Text, nullable=True)
    Landmark = Column(Text, nullable=True)


class ExecutionStatus(Base):
    __tablename__ = "tbl_executionstatus"
    ExecutionStatusID = Column(BigInteger, primary_key=True, index=True, autoincrement=True)
    StatusID = Column(BigInteger, nullable=True)
    CaseID = Column(BigInteger, nullable=True)
    ExecutionStatusName = Column(Text, index=True, nullable=True)
    procedure = Column(Text, nullable=True)
    ExecutionLevel = Column("Execution level", Text, nullable=True)
    Notes = Column(Text, nullable=True)
    IsClosed = Column(Boolean, nullable=True)


class Judgement(Base):
    __tablename__ = "tbl_judgements"
    JudgementID = Column(BigInteger, primary_key=True, index=True, autoincrement=True)
    CaseID = Column(BigInteger, nullable=True)
    HasAppeal = Column(Boolean, nullable=True)
    CourtDecision = Column(Text, nullable=True)
    CourtDecisionDate = Column(Text, nullable=True)
    DeputyGovernorDecision = Column("Deputy Governor Decision", Text, nullable=True)
    DateDeputyGovernorDecision = Column("Date Deputy Governor's Decision", Text, nullable=True)
    FinalJudgement = Column(Text, nullable=True)
    FinalJudgementDate = Column(Text, nullable=True)
    TheFine = Column("The fine", Text, nullable=True)
    NonExecutionNumber = Column(Text, nullable=True)
    DepartmentIssueDate = Column("Department issue date", Text, nullable=True)
    FelonyNumber = Column(Text, nullable=True)
    Notes = Column(Text, nullable=True)


class Reconciliation(Base):
    __tablename__ = "tbl_reconciliation"
    ReconciliationID = Column(BigInteger, primary_key=True, index=True, autoincrement=True)
    CaseID = Column(BigInteger, nullable=True)
    ReconciliationNumber = Column(Text, nullable=True)
    ReconciliationDate = Column(Text, nullable=True)
    FineAmount = Column(Text, nullable=True)
    NonExecutionNumber = Column(Text, nullable=True)
    DepartmentIssueDate = Column(Text, nullable=True)
    Notes = Column(Text, nullable=True)


class ViolationType(Base):
    __tablename__ = "tbl_violationtypes"
    ViolationTypeID = Column(BigInteger, primary_key=True, index=True, autoincrement=True)
    CaseID = Column(BigInteger, nullable=True)
    ViolationTypeName = Column(Text, nullable=True)
    ViolationDescription = Column(Text, nullable=True)


class Archive(Base):
    __tablename__ = "tbl_Archive"
    ArchiveID = Column(BigInteger, primary_key=True, index=True, autoincrement=True)
    CaseID = Column(BigInteger, nullable=True)
    FileType = Column(Text, nullable=True)
    FileName = Column(Text, nullable=True)
    FilePath = Column(Text, nullable=True)
    UploadDate = Column(Text, nullable=True)


try:
    Base.metadata.create_all(bind=engine)
except Exception as db_init_err:
    print(f"Warning during metadata create_all: {db_init_err}")

app = FastAPI(title="نظام إدارة المحاضر - المنظومة السيادية المؤمنة (السحابة)", version="26.11.28")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
os.makedirs(STATIC_DIR, exist_ok=True)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def log_audit(db: Session, username: str, action: str, table: str, details: str):
    try:
        log = AuditLog(
            Username=username or "مجهول",
            ActionType=action,
            TargetTable=table,
            Details=details,
            Timestamp=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        )
        db.add(log)
        db.commit()
    except Exception:
        try:
            db.rollback()
        except Exception:
            pass


def get_current_user_payload(credentials: HTTPAuthorizationCredentials = Depends(security)):
    return decode_access_token(credentials.credentials)


def build_filtered_cases_query(
    db: Session,
    q: Optional[str] = "",
    var_code: Optional[str] = "",
    case_no: Optional[str] = "",
    decision_no: Optional[str] = "",
    national_id: Optional[str] = "",
    lat: Optional[str] = "",
    lon: Optional[str] = "",
    offender: Optional[str] = "",
    address: Optional[str] = "",
    v_type: Optional[str] = "",
    status: Optional[str] = "",
    district: Optional[str] = "",
    vtype: Optional[str] = "",
    year: Optional[str] = None,
):
    case_cols = _get_table_columns("tbl_cases")
    off_cols = _get_table_columns("tbl_offenders")
    loc_cols = _get_table_columns("tbl_locations")
    exec_cols = _get_table_columns("tbl_executionstatus")
    judg_cols = _get_table_columns("tbl_judgements")
    recon_cols = _get_table_columns("tbl_reconciliation")
    vtype_cols = _get_table_columns("tbl_violationtypes")

    query = db.query(Case, Offender, Location, ExecutionStatus, Judgement, Reconciliation, ViolationType)

    if "OffenderID" in case_cols and "OffenderID" in off_cols:
        query = query.outerjoin(Offender, Case.OffenderID == Offender.OffenderID)
    elif "CaseID" in off_cols:
        query = query.outerjoin(Offender, Case.CaseID == Offender.CaseID)
    else:
        query = query.outerjoin(Offender, False)

    if "LocationID" in case_cols and "LocationID" in loc_cols:
        query = query.outerjoin(Location, Case.LocationID == Location.LocationID)
    elif "CaseID" in loc_cols:
        query = query.outerjoin(Location, Case.CaseID == Location.CaseID)
    else:
        query = query.outerjoin(Location, False)

    if "CaseID" in exec_cols:
        query = query.outerjoin(ExecutionStatus, Case.CaseID == ExecutionStatus.CaseID)
    else:
        query = query.outerjoin(ExecutionStatus, False)

    if "CaseID" in judg_cols:
        query = query.outerjoin(Judgement, Case.CaseID == Judgement.CaseID)
    else:
        query = query.outerjoin(Judgement, False)

    if "CaseID" in recon_cols:
        query = query.outerjoin(Reconciliation, Case.CaseID == Reconciliation.CaseID)
    else:
        query = query.outerjoin(Reconciliation, False)

    if "CaseID" in vtype_cols:
        query = query.outerjoin(ViolationType, Case.CaseID == ViolationType.CaseID)
    else:
        query = query.outerjoin(ViolationType, False)

    if var_code and var_code.strip() and "VariableCode" in case_cols:
        query = query.filter(Case.VariableCode.ilike(f"%{var_code.strip()}%"))
    if case_no and case_no.strip() and "CaseNumber" in case_cols:
        query = query.filter(Case.CaseNumber.ilike(f"%{case_no.strip()}%"))
    if decision_no and decision_no.strip() and "DecisionNumber" in case_cols:
        query = query.filter(Case.DecisionNumber.ilike(f"%{decision_no.strip()}%"))

    if national_id and national_id.strip():
        nat_q = f"%{national_id.strip()}%"
        nat_filters = []
        if "NationalID" in off_cols:
            nat_filters.append(Offender.NationalID.ilike(nat_q))
        if "NationalID" in case_cols:
            nat_filters.append(Case.NationalID.ilike(nat_q))
        if nat_filters:
            query = query.filter(or_(*nat_filters))

    if lat and lat.strip() and "Latitude" in case_cols:
        query = query.filter(Case.Latitude.ilike(f"%{lat.strip()}%"))
    if lon and lon.strip() and "Longitude" in case_cols:
        query = query.filter(Case.Longitude.ilike(f"%{lon.strip()}%"))

    if offender and offender.strip() and "OffenderName" in off_cols:
        query = query.filter(Offender.OffenderName.ilike(f"%{offender.strip()}%"))

    if address and address.strip():
        addr_filter = f"%{address.strip()}%"
        addr_filters = []
        if "FullAddress" in loc_cols:
            addr_filters.append(Location.FullAddress.ilike(addr_filter))
        if "StreetName" in loc_cols:
            addr_filters.append(Location.StreetName.ilike(addr_filter))
        if "OffenderAddress" in off_cols:
            addr_filters.append(Offender.OffenderAddress.ilike(addr_filter))
        if addr_filters:
            query = query.filter(or_(*addr_filters))

    effective_vtype = (v_type or vtype or "").strip()
    if effective_vtype and "ViolationTypeName" in vtype_cols:
        query = query.filter(ViolationType.ViolationTypeName.ilike(f"%{effective_vtype}%"))

    if q and q.strip():
        clean_q = q.strip()
        search_filter = f"%{clean_q}%"
        filters = []
        if "VariableCode" in case_cols:
            filters.append(Case.VariableCode.ilike(search_filter))
        if "CaseNumber" in case_cols:
            filters.append(Case.CaseNumber.ilike(search_filter))
        if "DecisionNumber" in case_cols:
            filters.append(Case.DecisionNumber.ilike(search_filter))
        if "Latitude" in case_cols:
            filters.append(Case.Latitude.ilike(search_filter))
        if "Longitude" in case_cols:
            filters.append(Case.Longitude.ilike(search_filter))
        if "NationalID" in case_cols:
            filters.append(Case.NationalID.ilike(search_filter))
        if "OffenderName" in off_cols:
            filters.append(Offender.OffenderName.ilike(search_filter))
        if "NationalID" in off_cols:
            filters.append(Offender.NationalID.ilike(search_filter))
        if "FullAddress" in loc_cols:
            filters.append(Location.FullAddress.ilike(search_filter))
        if "StreetName" in loc_cols:
            filters.append(Location.StreetName.ilike(search_filter))
        if "DistrictName" in loc_cols:
            filters.append(Location.DistrictName.ilike(search_filter))
        if "ViolationTypeName" in vtype_cols:
            filters.append(ViolationType.ViolationTypeName.ilike(search_filter))
        if filters:
            query = query.filter(or_(*filters))

    if status and status.strip():
        st_val = status.strip()
        st_filters = []
        if "ExecutionStatusName" in exec_cols:
            st_filters.append(ExecutionStatus.ExecutionStatusName == st_val)
        if "procedure" in exec_cols:
            st_filters.append(ExecutionStatus.procedure == st_val)
        if st_filters:
            query = query.filter(or_(*st_filters))

    if district and district.strip() and "DistrictName" in loc_cols:
        query = query.filter(Location.DistrictName.ilike(f"%{district.strip()}%"))

    if year is not None and str(year).strip() != "" and "year" in case_cols:
        query = query.filter(Case.year == str(year).strip())

    return query.order_by(Case.CaseID.desc())


class LoginPayload(BaseModel):
    username: str
    password: str


class UserPayload(BaseModel):
    Username: str
    Password: Optional[str] = ""
    FullName: Optional[str] = ""
    Role: Optional[str] = "DataEntry"
    IsActive: Optional[bool] = True


class CaseFullPayload(BaseModel):
    SerialNo: Optional[str] = None
    VariableCode: Optional[str] = ""
    VariableDate: Optional[str] = ""
    CaseNumber: Optional[str] = ""
    CaseDate: Optional[str] = ""
    year: Optional[str] = None
    StopNumber: Optional[str] = ""
    AdminNumber: Optional[str] = ""
    DecisionNumber: Optional[str] = ""
    DecisionDate: Optional[str] = ""
    AreaSize: Optional[str] = None
    Notes: Optional[str] = ""
    Latitude: Optional[str] = ""
    Longitude: Optional[str] = ""
    OffenderName: Optional[str] = ""
    OffenderNationalID: Optional[str] = None
    OffenderAddress: Optional[str] = ""
    Phone: Optional[str] = None
    OffenderNotes: Optional[str] = ""
    StreetName: Optional[str] = ""
    DistrictName: Optional[str] = ""
    FullAddress: Optional[str] = ""
    Landmark: Optional[str] = ""
    ExecutionStatusName: Optional[str] = "تنفيذ جزئي"
    procedure: Optional[str] = ""
    ExecutionLevel: Optional[str] = ""
    ExecNotes: Optional[str] = ""
    HasAppeal: Optional[bool] = False
    CourtDecision: Optional[str] = ""
    CourtDecisionDate: Optional[str] = None
    DeputyGovernorDecision: Optional[str] = None
    DateDeputyGovernorDecision: Optional[str] = None
    FinalJudgement: Optional[str] = None
    FinalJudgementDate: Optional[str] = None
    TheFine: Optional[str] = None
    NonExecutionNumber: Optional[str] = None
    DepartmentIssueDate: Optional[str] = None
    FelonyNumber: Optional[str] = None
    JudgementNotes: Optional[str] = ""
    ReconciliationNumber: Optional[str] = ""
    ReconciliationDate: Optional[str] = ""
    FineAmount: Optional[str] = None
    ReconNotes: Optional[str] = ""
    ViolationTypeName: Optional[str] = ""
    ViolationDescription: Optional[str] = ""

# --- مسارات المصادقة والمستخدمين ---
@app.post("/api/auth/login")
def api_login(payload: LoginPayload, db: Session = Depends(get_db)):
    clean_username = (payload.username or "").strip()
    clean_password = payload.password or ""
    if not clean_username or not clean_password:
        raise HTTPException(status_code=400, detail="يرجى إدخال اسم المستخدم وكلمة المرور")

    try:
        user = db.query(User).filter(User.Username == clean_username).first()
    except Exception as e:
        raise HTTPException(status_code=503, detail=f"تعذر الاتصال بقاعدة البيانات: {str(e)}")

    if not user or not user.IsActive:
        log_audit(db, clean_username, "LOGIN_FAILED", "tbl_users", "محاولة دخول لحساب غير موجود أو معطل")
        raise HTTPException(status_code=401, detail="اسم المستخدم غير موجود أو الحساب معطل أمنياً")

    if not verify_password(clean_password, user.PasswordHash):
        log_audit(db, clean_username, "LOGIN_FAILED", "tbl_users", "محاولة دخول فاشلة")
        raise HTTPException(status_code=401, detail="اسم المستخدم أو كلمة المرور غير صحيحة")

    token_expires = timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
    access_token = create_access_token(
        data={"sub": user.Username, "role": user.Role, "fullname": user.FullName or ""},
        expires_delta=token_expires
    )
    log_audit(db, user.Username, "LOGIN_SUCCESS", "tbl_users", "تم تسجيل الدخول بنجاح")
    return {
        "status": "success",
        "access_token": access_token,
        "token_type": "bearer",
        "username": user.Username,
        "fullname": user.FullName or "",
        "role": user.Role
    }

@app.get("/api/admin/users")
def get_all_users(user_data: dict = Depends(get_current_user_payload), db: Session = Depends(get_db)):
    if user_data.get("role") != "Admin":
        raise HTTPException(status_code=403, detail="غير مسموح")
    users = db.query(User).all()
    return {"status": "success", "data": [{"UserID": u.UserID, "Username": u.Username, "FullName": u.FullName or "", "Role": u.Role, "IsActive": u.IsActive, "CreatedAt": u.CreatedAt or ""} for u in users]}

@app.post("/api/admin/users")
def create_user(payload: UserPayload, user_data: dict = Depends(get_current_user_payload), db: Session = Depends(get_db)):
    if user_data.get("role") != "Admin":
        raise HTTPException(status_code=403, detail="غير مسموح")
    clean_uname = (payload.Username or "").strip()
    if not clean_uname:
        return {"status": "error", "message": "اسم المستخدم مطلوب"}
    existing = db.query(User).filter(User.Username == clean_uname).first()
    if existing:
        return {"status": "error", "message": "اسم المستخدم موجود مسبقاً"}
    hashed_pwd = get_password_hash(payload.Password if payload.Password else "SecurePass@2026")
    new_user = User(
        Username=clean_uname,
        PasswordHash=hashed_pwd,
        FullName=payload.FullName,
        Role=payload.Role,
        IsActive=payload.IsActive,
        CreatedAt=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    )
    db.add(new_user)
    db.commit()
    log_audit(db, user_data.get("sub"), "INSERT", "tbl_users", f"إنشاء مستخدم: {clean_uname}")
    return {"status": "success", "message": "تم إنشاء المستخدم بنجاح"}

@app.put("/api/admin/users/{user_id}")
def update_user(user_id: int, payload: UserPayload, user_data: dict = Depends(get_current_user_payload), db: Session = Depends(get_db)):
    if user_data.get("role") != "Admin":
        raise HTTPException(status_code=403, detail="غير مسموح")
    u = db.query(User).filter(User.UserID == user_id).first()
    if not u:
        raise HTTPException(status_code=404, detail="غير موجود")
    u.FullName = payload.FullName
    u.Role = payload.Role
    u.IsActive = payload.IsActive
    if payload.Password and payload.Password.strip():
        u.PasswordHash = get_password_hash(payload.Password)
    db.commit()
    log_audit(db, user_data.get("sub"), "UPDATE", "tbl_users", f"تعديل المستخدم ID: {user_id}")
    return {"status": "success", "message": "تم التعديل بنجاح"}

@app.delete("/api/admin/users/{user_id}")
def delete_user(user_id: int, user_data: dict = Depends(get_current_user_payload), db: Session = Depends(get_db)):
    if user_data.get("role") != "Admin":
        raise HTTPException(status_code=403, detail="غير مسموح")
    u = db.query(User).filter(User.UserID == user_id).first()
    if u:
        uname = u.Username
        db.delete(u)
        db.commit()
        log_audit(db, user_data.get("sub"), "DELETE", "tbl_users", f"حذف المستخدم: {uname}")
    return {"status": "success", "message": "تم الحذف بنجاح"}

@app.get("/api/users/stats")
def get_user_activity_stats(user_data: dict = Depends(get_current_user_payload), db: Session = Depends(get_db)):
    if user_data.get("role") != "Admin":
        raise HTTPException(status_code=403, detail="غير مسموح")
    user_records_count = db.query(AuditLog.Username, func.count(AuditLog.LogID)).filter(AuditLog.ActionType == "INSERT", AuditLog.TargetTable == "tbl_cases").group_by(AuditLog.Username).all()
    records_summary = {uname: cnt for uname, cnt in user_records_count if uname}
    activity_counts = db.query(AuditLog.Username, AuditLog.ActionType, func.count(AuditLog.LogID)).group_by(AuditLog.Username, AuditLog.ActionType).all()
    users_summary = {}
    for uname, action, cnt in activity_counts:
        if not uname:
            continue
        if uname not in users_summary:
            users_summary[uname] = {"INSERT": 0, "UPDATE": 0, "DELETE": 0, "LOGIN": 0, "TOTAL_ADDED_RECORDS": records_summary.get(uname, 0)}
        if action in users_summary[uname]:
            users_summary[uname][action] = cnt
    logs = db.query(AuditLog).order_by(AuditLog.LogID.desc()).limit(100).all()
    return {
        "status": "success",
        "summary": users_summary,
        "records_summary": records_summary,
        "logs": [{"LogID": l.LogID, "Username": l.Username, "ActionType": l.ActionType, "TargetTable": l.TargetTable, "Details": l.Details, "Timestamp": l.Timestamp} for l in logs]
    }

@app.get("/api/next-serial")
def get_next_serial(user_data: dict = Depends(get_current_user_payload), db: Session = Depends(get_db)):
    max_serial = db.query(func.max(Case.SerialNo)).scalar() or 0
    return {"next_serial": int(max_serial) + 1}

@app.get("/api/stats")
def get_stats(user_data: dict = Depends(get_current_user_payload), db: Session = Depends(get_db)):
    total_cases = db.query(func.count(Case.CaseID)).scalar() or 0
    variable_coded_count = db.query(func.count(Case.CaseID)).filter(Case.VariableCode.isnot(None), Case.VariableCode != "", Case.VariableCode != "-").scalar() or 0
    status_counts = db.query(ExecutionStatus.ExecutionStatusName, func.count(ExecutionStatus.CaseID)).group_by(ExecutionStatus.ExecutionStatusName).all()
    districts = [d[0] for d in db.query(Location.DistrictName).distinct().all() if d[0]]
    violation_types = [v[0] for v in db.query(ViolationType.ViolationTypeName).distinct().all() if v[0]]
    return {
        "status": "success",
        "total_cases": total_cases,
        "variable_coded_count": variable_coded_count,
        "status_breakdown": {str(st[0]): st[1] for st in status_counts if st[0]},
        "districts": districts,
        "violation_types": violation_types
    }

BACKUP_DIR = os.getenv("BACKUP_DIR", os.path.join(BASE_DIR, "backups"))

@app.post("/api/admin/backup")
def create_backup(user_data: dict = Depends(get_current_user_payload), db: Session = Depends(get_db)):
    if user_data.get("role") != "Admin":
        raise HTTPException(status_code=403, detail="غير مسموح")
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        backup_filename = f"Sovereign_Backup_{timestamp}.json"
        backup_path = os.path.join(BACKUP_DIR, backup_filename)
        backup_data = {
            "timestamp": timestamp,
            "cases": [{"CaseID": c.CaseID, "SerialNo": c.SerialNo, "VariableCode": c.VariableCode, "CaseNumber": c.CaseNumber, "year": c.year, "DecisionNumber": c.DecisionNumber} for c in db.query(Case).all()],
            "offenders": [{"OffenderID": o.OffenderID, "CaseID": o.CaseID, "OffenderName": o.OffenderName, "NationalID": o.NationalID} for o in db.query(Offender).all()],
            "users": [{"UserID": u.UserID, "Username": u.Username, "Role": u.Role} for u in db.query(User).all()]
        }
        with open(backup_path, "w", encoding="utf-8") as f:
            json.dump(backup_data, f, ensure_ascii=False, indent=4)
        log_audit(db, user_data.get("sub"), "BACKUP", "database", f"تم إنشاء نسخة احتياطية: {backup_filename}")
        return {"status": "success", "message": f"تم إنشاء النسخة الاحتياطية بنجاح: {backup_filename}"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/admin/backups")
def list_backups(user_data: dict = Depends(get_current_user_payload)):
    if user_data.get("role") != "Admin":
        raise HTTPException(status_code=403, detail="غير مسموح")
    files = []
    if os.path.exists(BACKUP_DIR):
        for f in os.listdir(BACKUP_DIR):
            f_path = os.path.join(BACKUP_DIR, f)
            if os.path.isfile(f_path):
                files.append({
                    "filename": f,
                    "size": f"{os.path.getsize(f_path) / (1024*1024):.2f} MB",
                    "time": datetime.fromtimestamp(os.path.getmtime(f_path)).strftime('%Y-%m-%d %H:%M:%S')
                })
    return {"status": "success", "backups": files}

@app.get("/api/admin/backup/download/{filename}")
def download_backup(filename: str, token: Optional[str] = Query(None), db: Session = Depends(get_db)):
    user_payload = None
    if token:
        user_payload = decode_access_token(token)
    if not user_payload or user_payload.get("role") != "Admin":
        raise HTTPException(status_code=403, detail="غير مسموح")
    safe_name = os.path.basename(filename)
    file_path = os.path.join(BACKUP_DIR, safe_name)
    if os.path.exists(file_path):
        return FileResponse(file_path, filename=safe_name, media_type='application/json')
    raise HTTPException(status_code=404, detail="الملف غير موجود")

def _safe_text(val, default=""):
    if val is None:
        return default
    s = str(val).strip()
    return default if s.lower() in ("", "none", "nan") else s


def _safe_int_or_none(val):
    if val is None or val == "":
        return None
    try:
        if isinstance(val, float) and math.isnan(val):
            return None
        return int(float(str(val).strip()))
    except Exception:
        return None


def _safe_float_or_none(val):
    if val is None or val == "":
        return None
    try:
        f = float(str(val).strip())
        return None if math.isnan(f) else f
    except Exception:
        return None


def _has_attr(obj, attr_name):
    return obj is not None and hasattr(obj, attr_name)


def _get_attr(obj, attr_name, default=None):
    if obj is None or not hasattr(obj, attr_name):
        return default
    return getattr(obj, attr_name)


@app.get("/api/cases/search")
def search_cases(
    q: Optional[str] = "",
    var_code: Optional[str] = "",
    case_no: Optional[str] = "",
    decision_no: Optional[str] = "",
    national_id: Optional[str] = "",
    lat: Optional[str] = "",
    lon: Optional[str] = "",
    offender: Optional[str] = "",
    address: Optional[str] = "",
    v_type: Optional[str] = "",
    status: Optional[str] = "",
    district: Optional[str] = "",
    vtype: Optional[str] = "",
    year: Optional[str] = None,
    page: int = Query(1, ge=1),
    limit: int = Query(50, ge=1, le=500),
    user_data: dict = Depends(get_current_user_payload),
    db: Session = Depends(get_db)
):
    try:
        query = build_filtered_cases_query(
            db=db,
            q=q,
            var_code=var_code,
            case_no=case_no,
            decision_no=decision_no,
            national_id=national_id,
            lat=lat,
            lon=lon,
            offender=offender,
            address=address,
            v_type=v_type,
            status=status,
            district=district,
            vtype=vtype,
            year=year,
        )

        total_count = query.count()
        results = query.offset((page - 1) * limit).limit(limit).all()

        data = []
        for row_tuple in results:
            r_case = row_tuple[0] if len(row_tuple) > 0 else None
            r_off = row_tuple[1] if len(row_tuple) > 1 else None
            r_loc = row_tuple[2] if len(row_tuple) > 2 else None
            r_exec = row_tuple[3] if len(row_tuple) > 3 else None
            r_judg = row_tuple[4] if len(row_tuple) > 4 else None
            r_recon = row_tuple[5] if len(row_tuple) > 5 else None
            r_vtype = row_tuple[6] if len(row_tuple) > 6 else None

            if not r_case:
                continue

            raw_nat_id = _get_attr(r_off, "NationalID") or _get_attr(r_case, "NationalID")
            nat_id_str = _safe_text(raw_nat_id, "")
            if nat_id_str.endswith(".0"):
                nat_id_str = nat_id_str[:-2]

            raw_phone = _get_attr(r_off, "Phone")
            phone_str = _safe_text(raw_phone, "")
            if phone_str.endswith(".0"):
                phone_str = phone_str[:-2]

            exec_status_val = _get_attr(r_exec, "ExecutionStatusName") or _get_attr(r_exec, "procedure") or "تنفيذ جزئي"

            data.append({
                "CaseID": _get_attr(r_case, "CaseID"),
                "SerialNo": _safe_text(_get_attr(r_case, "SerialNo"), ""),
                "VariableCode": _safe_text(_get_attr(r_case, "VariableCode"), "-"),
                "VariableDate": _safe_text(_get_attr(r_case, "VariableDate"), ""),
                "CaseNumber": _safe_text(_get_attr(r_case, "CaseNumber"), "-"),
                "CaseDate": _safe_text(_get_attr(r_case, "CaseDate"), ""),
                "DecisionNumber": _safe_text(_get_attr(r_case, "DecisionNumber"), "-"),
                "DecisionDate": _safe_text(_get_attr(r_case, "DecisionDate"), ""),
                "year": _safe_text(_get_attr(r_case, "year"), ""),
                "StopNumber": _safe_text(_get_attr(r_case, "StopNumber"), ""),
                "AdminNumber": _safe_text(_get_attr(r_case, "AdminNumber"), ""),
                "AreaSize": _safe_float_or_none(_get_attr(r_case, "AreaSize")) or 0.0,
                "Notes": _safe_text(_get_attr(r_case, "Notes"), ""),
                "Latitude": _safe_text(_get_attr(r_case, "Latitude"), ""),
                "Longitude": _safe_text(_get_attr(r_case, "Longitude"), ""),
                "OffenderName": _safe_text(_get_attr(r_off, "OffenderName"), "غير مسجل") if r_off else "غير مسجل",
                "OffenderNationalID": nat_id_str,
                "OffenderAddress": _safe_text(_get_attr(r_off, "OffenderAddress"), ""),
                "Phone": phone_str,
                "OffenderNotes": _safe_text(_get_attr(r_off, "Notes"), ""),
                "StreetName": _safe_text(_get_attr(r_loc, "StreetName"), ""),
                "DistrictName": _safe_text(_get_attr(r_loc, "DistrictName"), ""),
                "FullAddress": _safe_text(_get_attr(r_loc, "FullAddress"), "") or _safe_text(_get_attr(r_loc, "StreetName"), ""),
                "Landmark": _safe_text(_get_attr(r_loc, "Landmark"), ""),
                "ExecutionStatusName": _safe_text(exec_status_val, "تنفيذ جزئي"),
                "procedure": _safe_text(_get_attr(r_exec, "procedure"), ""),
                "ExecutionLevel": _safe_text(_get_attr(r_exec, "ExecutionLevel"), ""),
                "ExecNotes": _safe_text(_get_attr(r_exec, "Notes"), ""),
                "HasAppeal": bool(_get_attr(r_judg, "HasAppeal", False)),
                "CourtDecision": _safe_text(_get_attr(r_judg, "CourtDecision"), ""),
                "CourtDecisionDate": _get_attr(r_judg, "CourtDecisionDate"),
                "DeputyGovernorDecision": _get_attr(r_judg, "DeputyGovernorDecision"),
                "DateDeputyGovernorDecision": _get_attr(r_judg, "DateDeputyGovernorDecision"),
                "FinalJudgement": _get_attr(r_judg, "FinalJudgement"),
                "FinalJudgementDate": _get_attr(r_judg, "FinalJudgementDate"),
                "TheFine": _get_attr(r_judg, "TheFine"),
                "NonExecutionNumber": _get_attr(r_judg, "NonExecutionNumber"),
                "DepartmentIssueDate": _get_attr(r_judg, "DepartmentIssueDate"),
                "FelonyNumber": _get_attr(r_judg, "FelonyNumber"),
                "JudgementNotes": _safe_text(_get_attr(r_judg, "Notes"), ""),
                "ReconciliationNumber": _safe_text(_get_attr(r_recon, "ReconciliationNumber"), ""),
                "ReconciliationDate": _safe_text(_get_attr(r_recon, "ReconciliationDate"), ""),
                "FineAmount": _get_attr(r_recon, "FineAmount"),
                "ReconNotes": _safe_text(_get_attr(r_recon, "Notes"), ""),
                "ViolationTypeName": _safe_text(_get_attr(r_vtype, "ViolationTypeName"), ""),
                "ViolationDescription": _safe_text(_get_attr(r_vtype, "ViolationDescription"), ""),
            })

        return {
            "status": "success",
            "total": total_count,
            "page": page,
            "limit": limit,
            "total_pages": math.ceil(total_count / limit) if total_count > 0 else 1,
            "data": data,
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"خطأ أثناء جلب بيانات المحاضر: {str(e)}")


# --- مسارات الإضافة والتعديل والحذف والأرشيف المتوافقة مع السحابة وملف models ---
@app.post("/api/cases")
def create_case(payload: CaseFullPayload, user_data: dict = Depends(get_current_user_payload), db: Session = Depends(get_db)):
    if user_data.get("role") == "Viewer":
        raise HTTPException(status_code=403, detail="غير مسموح")
    try:
        case_num_clean = _safe_text(payload.CaseNumber, "")
        year_clean = _safe_text(payload.year, "")
        if case_num_clean and case_num_clean != "-":
            q_dup = db.query(Case).filter(Case.CaseNumber == case_num_clean)
            if year_clean and _has_attr(Case, "year"):
                q_dup = q_dup.filter(Case.year == year_clean)
            if q_dup.first():
                return {"status": "error", "message": f"عذراً، رقم المحضر ({case_num_clean}) لنفس السنة مسجل مسبقاً ولا يمكن تكراره لنفس السنة!"}

        offender_id_val = None
        if _has_attr(Case, "OffenderID"):
            off_kwargs = {}
            if _has_attr(Offender, "OffenderName"):
                off_kwargs["OffenderName"] = _safe_text(payload.OffenderName, None)
            if _has_attr(Offender, "NationalID"):
                nat_s = _safe_text(payload.OffenderNationalID, None)
                off_kwargs["NationalID"] = nat_s
            if _has_attr(Offender, "OffenderAddress"):
                off_kwargs["OffenderAddress"] = _safe_text(payload.OffenderAddress, None)
            if _has_attr(Offender, "Phone"):
                off_kwargs["Phone"] = _safe_text(payload.Phone, None)
            if _has_attr(Offender, "Notes"):
                off_kwargs["Notes"] = _safe_text(payload.OffenderNotes, None)
            new_off = Offender(**off_kwargs)
            db.add(new_off)
            db.flush()
            offender_id_val = new_off.OffenderID

        location_id_val = None
        if _has_attr(Case, "LocationID"):
            loc_kwargs = {}
            if _has_attr(Location, "StreetName"):
                loc_kwargs["StreetName"] = _safe_text(payload.StreetName, None)
            if _has_attr(Location, "DistrictName"):
                loc_kwargs["DistrictName"] = _safe_text(payload.DistrictName, None)
            if _has_attr(Location, "FullAddress"):
                loc_kwargs["FullAddress"] = _safe_text(payload.FullAddress, None)
            if _has_attr(Location, "Landmark"):
                loc_kwargs["Landmark"] = _safe_text(payload.Landmark, None)
            new_loc = Location(**loc_kwargs)
            db.add(new_loc)
            db.flush()
            location_id_val = new_loc.LocationID

        case_kwargs = {}
        if _has_attr(Case, "SerialNo"):
            case_kwargs["SerialNo"] = _safe_int_or_none(payload.SerialNo)
        if _has_attr(Case, "VariableCode"):
            case_kwargs["VariableCode"] = _safe_text(payload.VariableCode, None)
        if _has_attr(Case, "VariableDate"):
            case_kwargs["VariableDate"] = _safe_text(payload.VariableDate, None)
        if _has_attr(Case, "CaseNumber"):
            case_kwargs["CaseNumber"] = case_num_clean or None
        if _has_attr(Case, "CaseDate"):
            case_kwargs["CaseDate"] = _safe_text(payload.CaseDate, None)
        if _has_attr(Case, "NationalID"):
            case_kwargs["NationalID"] = _safe_text(payload.OffenderNationalID, None)
        if _has_attr(Case, "year"):
            case_kwargs["year"] = year_clean or None
        if _has_attr(Case, "StopNumber"):
            case_kwargs["StopNumber"] = _safe_text(payload.StopNumber, None)
        if _has_attr(Case, "AdminNumber"):
            case_kwargs["AdminNumber"] = _safe_text(payload.AdminNumber, None)
        if _has_attr(Case, "DecisionNumber"):
            case_kwargs["DecisionNumber"] = _safe_text(payload.DecisionNumber, None)
        if _has_attr(Case, "DecisionDate"):
            case_kwargs["DecisionDate"] = _safe_text(payload.DecisionDate, None)
        if _has_attr(Case, "AreaSize"):
            case_kwargs["AreaSize"] = _safe_float_or_none(payload.AreaSize)
        if _has_attr(Case, "Notes"):
            case_kwargs["Notes"] = _safe_text(payload.Notes, None)
        if _has_attr(Case, "Latitude"):
            case_kwargs["Latitude"] = _safe_text(payload.Latitude, None)
        if _has_attr(Case, "Longitude"):
            case_kwargs["Longitude"] = _safe_text(payload.Longitude, None)
        if _has_attr(Case, "OffenderID"):
            case_kwargs["OffenderID"] = offender_id_val
        if _has_attr(Case, "LocationID"):
            case_kwargs["LocationID"] = location_id_val

        new_case = Case(**case_kwargs)
        db.add(new_case)
        db.flush()
        cid = new_case.CaseID

        if not _has_attr(Case, "OffenderID") and _has_attr(Offender, "CaseID"):
            db.add(Offender(
                CaseID=cid,
                OffenderName=_safe_text(payload.OffenderName, None),
                NationalID=_safe_text(payload.OffenderNationalID, None),
                OffenderAddress=_safe_text(payload.OffenderAddress, None) if _has_attr(Offender, "OffenderAddress") else None,
                Phone=_safe_text(payload.Phone, None) if _has_attr(Offender, "Phone") else None,
                Notes=_safe_text(payload.OffenderNotes, None) if _has_attr(Offender, "Notes") else None,
            ))

        if not _has_attr(Case, "LocationID") and _has_attr(Location, "CaseID"):
            db.add(Location(
                CaseID=cid,
                StreetName=_safe_text(payload.StreetName, None),
                DistrictName=_safe_text(payload.DistrictName, None),
                FullAddress=_safe_text(payload.FullAddress, None) if _has_attr(Location, "FullAddress") else None,
                Landmark=_safe_text(payload.Landmark, None) if _has_attr(Location, "Landmark") else None,
            ))

        exec_kwargs = {"CaseID": cid}
        if _has_attr(ExecutionStatus, "ExecutionStatusName"):
            exec_kwargs["ExecutionStatusName"] = _safe_text(payload.ExecutionStatusName, "تنفيذ جزئي")
        if _has_attr(ExecutionStatus, "procedure"):
            exec_kwargs["procedure"] = _safe_text(payload.procedure, None) or _safe_text(payload.ExecutionStatusName, "تنفيذ جزئي")
        if _has_attr(ExecutionStatus, "ExecutionLevel"):
            exec_kwargs["ExecutionLevel"] = _safe_text(payload.ExecutionLevel, None)
        if _has_attr(ExecutionStatus, "Notes"):
            exec_kwargs["Notes"] = _safe_text(payload.ExecNotes, None)
        db.add(ExecutionStatus(**exec_kwargs))

        if "Judgement" in globals() and _has_attr(Judgement, "CaseID"):
            db.add(Judgement(
                CaseID=cid,
                HasAppeal=payload.HasAppeal,
                CourtDecision=_safe_text(payload.CourtDecision, None),
                CourtDecisionDate=payload.CourtDecisionDate,
                DeputyGovernorDecision=payload.DeputyGovernorDecision,
                DateDeputyGovernorDecision=payload.DateDeputyGovernorDecision,
                FinalJudgement=payload.FinalJudgement,
                FinalJudgementDate=payload.FinalJudgementDate,
                TheFine=payload.TheFine,
                NonExecutionNumber=payload.NonExecutionNumber,
                DepartmentIssueDate=payload.DepartmentIssueDate,
                FelonyNumber=payload.FelonyNumber,
                Notes=_safe_text(payload.JudgementNotes, None),
            ))

        if "Reconciliation" in globals() and _has_attr(Reconciliation, "CaseID"):
            db.add(Reconciliation(
                CaseID=cid,
                ReconciliationNumber=_safe_text(payload.ReconciliationNumber, None),
                ReconciliationDate=_safe_text(payload.ReconciliationDate, None),
                FineAmount=payload.FineAmount,
                Notes=_safe_text(payload.ReconNotes, None),
            ))

        if "ViolationType" in globals() and _has_attr(ViolationType, "CaseID"):
            db.add(ViolationType(
                CaseID=cid,
                ViolationTypeName=_safe_text(payload.ViolationTypeName, None),
                ViolationDescription=_safe_text(payload.ViolationDescription, None),
            ))

        db.commit()
        log_audit(db, user_data.get("sub"), "INSERT", "tbl_cases", f"إضافة محضر جديد ID: {cid}")
        return {"status": "success", "message": "تم إضافة المحضر بنجاح"}
    except Exception as e:
        db.rollback()
        return {"status": "error", "message": str(e)}


@app.put("/api/cases/{case_id}")
def update_case(case_id: int, payload: CaseFullPayload, user_data: dict = Depends(get_current_user_payload), db: Session = Depends(get_db)):
    if user_data.get("role") == "Viewer":
        raise HTTPException(status_code=403, detail="غير مسموح")
    try:
        c = db.query(Case).filter(Case.CaseID == case_id).first()
        if not c:
            raise HTTPException(status_code=404, detail="غير موجود")

        if _has_attr(Case, "SerialNo"):
            c.SerialNo = _safe_int_or_none(payload.SerialNo)
        if _has_attr(Case, "VariableCode"):
            c.VariableCode = _safe_text(payload.VariableCode, None)
        if _has_attr(Case, "VariableDate"):
            c.VariableDate = _safe_text(payload.VariableDate, None)
        if _has_attr(Case, "CaseNumber"):
            c.CaseNumber = _safe_text(payload.CaseNumber, None)
        if _has_attr(Case, "CaseDate"):
            c.CaseDate = _safe_text(payload.CaseDate, None)
        if _has_attr(Case, "NationalID"):
            c.NationalID = _safe_text(payload.OffenderNationalID, None)
        if _has_attr(Case, "year"):
            c.year = _safe_text(payload.year, None)
        if _has_attr(Case, "StopNumber"):
            c.StopNumber = _safe_text(payload.StopNumber, None)
        if _has_attr(Case, "AdminNumber"):
            c.AdminNumber = _safe_text(payload.AdminNumber, None)
        if _has_attr(Case, "DecisionNumber"):
            c.DecisionNumber = _safe_text(payload.DecisionNumber, None)
        if _has_attr(Case, "DecisionDate"):
            c.DecisionDate = _safe_text(payload.DecisionDate, None)
        if _has_attr(Case, "AreaSize"):
            c.AreaSize = _safe_float_or_none(payload.AreaSize)
        if _has_attr(Case, "Notes"):
            c.Notes = _safe_text(payload.Notes, None)
        if _has_attr(Case, "Latitude"):
            c.Latitude = _safe_text(payload.Latitude, None)
        if _has_attr(Case, "Longitude"):
            c.Longitude = _safe_text(payload.Longitude, None)

        off = None
        if _has_attr(Case, "OffenderID") and getattr(c, "OffenderID", None):
            off = db.query(Offender).filter(Offender.OffenderID == c.OffenderID).first()
        elif _has_attr(Offender, "CaseID"):
            off = db.query(Offender).filter(Offender.CaseID == case_id).first()

        if off:
            if _has_attr(Offender, "OffenderName"):
                off.OffenderName = _safe_text(payload.OffenderName, None)
            if _has_attr(Offender, "NationalID"):
                off.NationalID = _safe_text(payload.OffenderNationalID, None)
            if _has_attr(Offender, "OffenderAddress"):
                off.OffenderAddress = _safe_text(payload.OffenderAddress, None)
            if _has_attr(Offender, "Phone"):
                off.Phone = _safe_text(payload.Phone, None)
            if _has_attr(Offender, "Notes"):
                off.Notes = _safe_text(payload.OffenderNotes, None)
        else:
            off_kwargs = {}
            if _has_attr(Offender, "CaseID"):
                off_kwargs["CaseID"] = case_id
            if _has_attr(Offender, "OffenderName"):
                off_kwargs["OffenderName"] = _safe_text(payload.OffenderName, None)
            if _has_attr(Offender, "NationalID"):
                off_kwargs["NationalID"] = _safe_text(payload.OffenderNationalID, None)
            if _has_attr(Offender, "OffenderAddress"):
                off_kwargs["OffenderAddress"] = _safe_text(payload.OffenderAddress, None)
            if _has_attr(Offender, "Phone"):
                off_kwargs["Phone"] = _safe_text(payload.Phone, None)
            if _has_attr(Offender, "Notes"):
                off_kwargs["Notes"] = _safe_text(payload.OffenderNotes, None)
            new_off = Offender(**off_kwargs)
            db.add(new_off)
            db.flush()
            if _has_attr(Case, "OffenderID"):
                c.OffenderID = new_off.OffenderID

        loc = None
        if _has_attr(Case, "LocationID") and getattr(c, "LocationID", None):
            loc = db.query(Location).filter(Location.LocationID == c.LocationID).first()
        elif _has_attr(Location, "CaseID"):
            loc = db.query(Location).filter(Location.CaseID == case_id).first()

        if loc:
            if _has_attr(Location, "StreetName"):
                loc.StreetName = _safe_text(payload.StreetName, None)
            if _has_attr(Location, "DistrictName"):
                loc.DistrictName = _safe_text(payload.DistrictName, None)
            if _has_attr(Location, "FullAddress"):
                loc.FullAddress = _safe_text(payload.FullAddress, None)
            if _has_attr(Location, "Landmark"):
                loc.Landmark = _safe_text(payload.Landmark, None)
        else:
            loc_kwargs = {}
            if _has_attr(Location, "CaseID"):
                loc_kwargs["CaseID"] = case_id
            if _has_attr(Location, "StreetName"):
                loc_kwargs["StreetName"] = _safe_text(payload.StreetName, None)
            if _has_attr(Location, "DistrictName"):
                loc_kwargs["DistrictName"] = _safe_text(payload.DistrictName, None)
            if _has_attr(Location, "FullAddress"):
                loc_kwargs["FullAddress"] = _safe_text(payload.FullAddress, None)
            if _has_attr(Location, "Landmark"):
                loc_kwargs["Landmark"] = _safe_text(payload.Landmark, None)
            new_loc = Location(**loc_kwargs)
            db.add(new_loc)
            db.flush()
            if _has_attr(Case, "LocationID"):
                c.LocationID = new_loc.LocationID

        ex = db.query(ExecutionStatus).filter(ExecutionStatus.CaseID == case_id).first()
        if ex:
            if _has_attr(ExecutionStatus, "ExecutionStatusName"):
                ex.ExecutionStatusName = _safe_text(payload.ExecutionStatusName, "تنفيذ جزئي")
            if _has_attr(ExecutionStatus, "procedure"):
                ex.procedure = _safe_text(payload.procedure, None) or _safe_text(payload.ExecutionStatusName, "تنفيذ جزئي")
            if _has_attr(ExecutionStatus, "ExecutionLevel"):
                ex.ExecutionLevel = _safe_text(payload.ExecutionLevel, None)
            if _has_attr(ExecutionStatus, "Notes"):
                ex.Notes = _safe_text(payload.ExecNotes, None)
        else:
            exec_kwargs = {"CaseID": case_id}
            if _has_attr(ExecutionStatus, "ExecutionStatusName"):
                exec_kwargs["ExecutionStatusName"] = _safe_text(payload.ExecutionStatusName, "تنفيذ جزئي")
            if _has_attr(ExecutionStatus, "procedure"):
                exec_kwargs["procedure"] = _safe_text(payload.procedure, None) or _safe_text(payload.ExecutionStatusName, "تنفيذ جزئي")
            if _has_attr(ExecutionStatus, "ExecutionLevel"):
                exec_kwargs["ExecutionLevel"] = _safe_text(payload.ExecutionLevel, None)
            if _has_attr(ExecutionStatus, "Notes"):
                exec_kwargs["Notes"] = _safe_text(payload.ExecNotes, None)
            db.add(ExecutionStatus(**exec_kwargs))

        if "Judgement" in globals() and _has_attr(Judgement, "CaseID"):
            jd = db.query(Judgement).filter(Judgement.CaseID == case_id).first()
            if jd:
                jd.HasAppeal = payload.HasAppeal
                jd.CourtDecision = _safe_text(payload.CourtDecision, None)
                jd.CourtDecisionDate = payload.CourtDecisionDate
                jd.DeputyGovernorDecision = payload.DeputyGovernorDecision
                jd.DateDeputyGovernorDecision = payload.DateDeputyGovernorDecision
                jd.FinalJudgement = payload.FinalJudgement
                jd.FinalJudgementDate = payload.FinalJudgementDate
                jd.TheFine = payload.TheFine
                jd.NonExecutionNumber = payload.NonExecutionNumber
                jd.DepartmentIssueDate = payload.DepartmentIssueDate
                jd.FelonyNumber = payload.FelonyNumber
                jd.Notes = _safe_text(payload.JudgementNotes, None)
            else:
                db.add(Judgement(
                    CaseID=case_id, HasAppeal=payload.HasAppeal, CourtDecision=_safe_text(payload.CourtDecision, None),
                    CourtDecisionDate=payload.CourtDecisionDate, DeputyGovernorDecision=payload.DeputyGovernorDecision,
                    DateDeputyGovernorDecision=payload.DateDeputyGovernorDecision, FinalJudgement=payload.FinalJudgement,
                    FinalJudgementDate=payload.FinalJudgementDate, TheFine=payload.TheFine,
                    NonExecutionNumber=payload.NonExecutionNumber, DepartmentIssueDate=payload.DepartmentIssueDate,
                    FelonyNumber=payload.FelonyNumber, Notes=_safe_text(payload.JudgementNotes, None),
                ))

        if "Reconciliation" in globals() and _has_attr(Reconciliation, "CaseID"):
            rc = db.query(Reconciliation).filter(Reconciliation.CaseID == case_id).first()
            if rc:
                rc.ReconciliationNumber = _safe_text(payload.ReconciliationNumber, None)
                rc.ReconciliationDate = _safe_text(payload.ReconciliationDate, None)
                rc.FineAmount = payload.FineAmount
                rc.Notes = _safe_text(payload.ReconNotes, None)
            else:
                db.add(Reconciliation(
                    CaseID=case_id, ReconciliationNumber=_safe_text(payload.ReconciliationNumber, None),
                    ReconciliationDate=_safe_text(payload.ReconciliationDate, None), FineAmount=payload.FineAmount,
                    Notes=_safe_text(payload.ReconNotes, None),
                ))

        if "ViolationType" in globals() and _has_attr(ViolationType, "CaseID"):
            vt = db.query(ViolationType).filter(ViolationType.CaseID == case_id).first()
            if vt:
                vt.ViolationTypeName = _safe_text(payload.ViolationTypeName, None)
                vt.ViolationDescription = _safe_text(payload.ViolationDescription, None)
            else:
                db.add(ViolationType(
                    CaseID=case_id, ViolationTypeName=_safe_text(payload.ViolationTypeName, None),
                    ViolationDescription=_safe_text(payload.ViolationDescription, None),
                ))

        db.commit()
        log_audit(db, user_data.get("sub"), "UPDATE", "tbl_cases", f"تعديل المحضر ID: {case_id}")
        return {"status": "success", "message": "تم تعديل بيانات المحضر بنجاح"}
    except Exception as e:
        db.rollback()
        return {"status": "error", "message": str(e)}


@app.delete("/api/cases/{case_id}")
def delete_case(case_id: int, user_data: dict = Depends(get_current_user_payload), db: Session = Depends(get_db)):
    if user_data.get("role") not in ["Admin", "Supervisor"]:
        raise HTTPException(status_code=403, detail="غير مسموح")
    try:
        if _has_attr(ExecutionStatus, "CaseID"):
            db.query(ExecutionStatus).filter(ExecutionStatus.CaseID == case_id).delete()
        if "Judgement" in globals() and _has_attr(Judgement, "CaseID"):
            db.query(Judgement).filter(Judgement.CaseID == case_id).delete()
        if "Reconciliation" in globals() and _has_attr(Reconciliation, "CaseID"):
            db.query(Reconciliation).filter(Reconciliation.CaseID == case_id).delete()
        if "ViolationType" in globals() and _has_attr(ViolationType, "CaseID"):
            db.query(ViolationType).filter(ViolationType.CaseID == case_id).delete()
        if _has_attr(Archive, "CaseID"):
            db.query(Archive).filter(Archive.CaseID == case_id).delete()
        if _has_attr(Offender, "CaseID"):
            db.query(Offender).filter(Offender.CaseID == case_id).delete()
        if _has_attr(Location, "CaseID"):
            db.query(Location).filter(Location.CaseID == case_id).delete()

        db.query(Case).filter(Case.CaseID == case_id).delete()
        db.commit()
        log_audit(db, user_data.get("sub"), "DELETE", "tbl_cases", f"حذف سجل ID: {case_id}")
        return {"status": "success", "message": "تم الحذف بنجاح"}
    except Exception as e:
        db.rollback()
        return {"status": "error", "message": str(e)}


ARCHIVE_BASE_DIR = os.getenv("ARCHIVE_BASE_DIR", os.path.join(BASE_DIR, "archive_violations"))

@app.delete("/api/cases/{case_id}")
def delete_case(case_id: int, user_data: dict = Depends(get_current_user_payload), db: Session = Depends(get_db)):
    if user_data.get("role") not in ["Admin", "Supervisor"]:
        raise HTTPException(status_code=403, detail="غير مسموح")
    try:
        db.query(Case).filter(Case.CaseID == case_id).delete()
        db.query(ExecutionStatus).filter(ExecutionStatus.CaseID == case_id).delete()
        db.query(Offender).filter(Offender.CaseID == case_id).delete()
        db.query(Location).filter(Location.CaseID == case_id).delete()
        db.query(Judgement).filter(Judgement.CaseID == case_id).delete()
        db.query(Reconciliation).filter(Reconciliation.CaseID == case_id).delete()
        db.query(ViolationType).filter(ViolationType.CaseID == case_id).delete()
        db.query(Archive).filter(Archive.CaseID == case_id).delete()
        db.commit()
        log_audit(db, user_data.get("sub"), "DELETE", "tbl_cases", f"حذف سجل ID: {case_id}")
        return {"status": "success", "message": "تم الحذف بنجاح"}
    except Exception as e:
        db.rollback()
        return {"status": "error", "message": str(e)}

ARCHIVE_BASE_DIR = os.getenv("ARCHIVE_BASE_DIR", os.path.join(BASE_DIR, "archive_violations"))

def get_offender_folder_and_prefix(db: Session, case_id: int):
    case = db.query(Case).filter(Case.CaseID == case_id).first()
    offender = db.query(Offender).filter(Offender.CaseID == case_id).first()
    offender_name = (offender.OffenderName if offender and offender.OffenderName else "مخالف_غير_معروف").strip()
    offender_name = "".join([c for c in offender_name if c.isalnum() or c in (' ', '_', '-')]).strip() or "مخالف_غير_معروف"
    file_prefix = "مستند"
    if case:
        if case.VariableCode and case.VariableCode.strip() and case.VariableCode != "-":
            file_prefix = case.VariableCode.strip()
        elif case.CaseNumber and case.CaseNumber.strip() and case.CaseNumber != "-":
            file_prefix = case.CaseNumber.strip()
    return offender_name, file_prefix

def _save_archive_record(db: Session, case_id: int, rel_path: str, filename: str, file_type: str):
    try:
        rec = Archive(
            CaseID=case_id,
            FileType=file_type,
            FileName=filename,
            FilePath=rel_path,
            UploadDate=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        )
        db.add(rec)
        db.commit()
    except Exception:
        try:
            db.rollback()
        except Exception:
            pass

@app.get("/api/cases/{case_id}/archive-files")
def get_case_archive_files(case_id: int, user_data: dict = Depends(get_current_user_payload), db: Session = Depends(get_db)):
    offender_name, file_prefix = get_offender_folder_and_prefix(db, case_id)
    offender_dir = os.path.join(ARCHIVE_BASE_DIR, offender_name)
    found_files = []
    if os.path.exists(offender_dir):
        pattern = os.path.join(offender_dir, f"{file_prefix}_*.*")
        for file_path in glob.glob(pattern):
            rel_path = os.path.join(offender_name, os.path.basename(file_path)).replace("\\", "/")
            if rel_path not in found_files:
                found_files.append(rel_path)

    try:
        db_archives = db.query(Archive).filter(Archive.CaseID == case_id).all()
        for item in db_archives:
            if item.FilePath:
                clean_rel = item.FilePath.replace("\\", "/")
                full_p = os.path.join(ARCHIVE_BASE_DIR, clean_rel)
                if os.path.exists(full_p) and clean_rel not in found_files:
                    found_files.append(clean_rel)
    except Exception:
        pass

    return {"status": "success", "files": found_files}

@app.get("/api/archive/file/{path:path}")
def serve_archive_file(path: str):
    clean_path = os.path.normpath(path).lstrip("\\/").replace("\\", "/")
    file_path = os.path.abspath(os.path.join(ARCHIVE_BASE_DIR, clean_path))
    base_abs = os.path.abspath(ARCHIVE_BASE_DIR)
    if not file_path.startswith(base_abs):
        raise HTTPException(status_code=403, detail="مسار غير مسموح")
    if os.path.exists(file_path) and os.path.isfile(file_path):
        return FileResponse(file_path)
    raise HTTPException(status_code=404, detail="الملف غير موجود على الخادم")

@app.get("/api/archive/open-folder/{case_id}")
def open_offender_folder(case_id: int, user_data: dict = Depends(get_current_user_payload), db: Session = Depends(get_db)):
    # تم حذف أمر تشغيل explorer المحلي ليتوافق مع بيئة السحابة بدون أخطاء
    offender_name, _ = get_offender_folder_and_prefix(db, case_id)
    offender_dir = os.path.join(ARCHIVE_BASE_DIR, offender_name)
    os.makedirs(offender_dir, exist_ok=True)
    return {
        "status": "success",
        "cloud_mode": True,
        "folder": offender_name,
        "message": "النظام يعمل على السحابة؛ يمكنك استعراض المستندات أو رفعها وتحميلها مباشرة من نافذة الأرشيف."
    }

@app.post("/api/cases/{case_id}/archive/upload")
async def upload_archive_file(case_id: int, file: UploadFile = File(...), user_data: dict = Depends(get_current_user_payload), db: Session = Depends(get_db)):
    if user_data.get("role") == "Viewer":
        raise HTTPException(status_code=403, detail="غير مسموح")
    offender_name, file_prefix = get_offender_folder_and_prefix(db, case_id)
    offender_dir = os.path.join(ARCHIVE_BASE_DIR, offender_name)
    os.makedirs(offender_dir, exist_ok=True)
    ext = os.path.splitext(file.filename or "doc.jpg")[1] or ".jpg"
    safe_filename = f"{file_prefix}_{int(datetime.now().timestamp())}{ext}"
    file_path = os.path.join(offender_dir, safe_filename)
    try:
        contents = await file.read()
        with open(file_path, "wb") as f:
            f.write(contents)
        rel_p = os.path.join(offender_name, safe_filename).replace("\\", "/")
        _save_archive_record(db, case_id, rel_p, safe_filename, ext.lstrip(".").lower())
        log_audit(db, user_data.get("sub"), "UPLOAD", "tbl_Archive", f"رفع ملف أرشيف: {rel_p}")
        return {"status": "success", "filename": rel_p}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/cases/{case_id}/archive/scan")
async def scan_and_save_document(
    case_id: int,
    file: Optional[UploadFile] = File(None),
    user_data: dict = Depends(get_current_user_payload),
    db: Session = Depends(get_db)
):
    """
    مسار المسح الضوئي المتوافق مع السحابة:
    يستقبل الصورة الملتقطة من كاميرا/ماسح جهاز المستخدم مباشرة عبر المتصفح (UploadFile)
    ويحفظها في مجلد الأرشيف وقاعدة البيانات، أو يلتقط محلياً فقط إذا لم يُرسل ملف وكان WIA متاحاً.
    """
    if user_data.get("role") == "Viewer":
        raise HTTPException(status_code=403, detail="غير مسموح")
    offender_name, file_prefix = get_offender_folder_and_prefix(db, case_id)
    offender_dir = os.path.join(ARCHIVE_BASE_DIR, offender_name)
    os.makedirs(offender_dir, exist_ok=True)

    if file is not None:
        try:
            ext = os.path.splitext(file.filename or "scan.jpg")[1] or ".jpg"
            safe_filename = f"{file_prefix}_scan_{int(datetime.now().timestamp())}{ext}"
            file_path = os.path.join(offender_dir, safe_filename)
            contents = await file.read()
            with open(file_path, "wb") as f:
                f.write(contents)
            rel_p = os.path.join(offender_name, safe_filename).replace("\\", "/")
            _save_archive_record(db, case_id, rel_p, safe_filename, ext.lstrip(".").lower())
            log_audit(db, user_data.get("sub"), "SCAN", "tbl_Archive", f"مسح ضوئي وحفظ سحابي: {rel_p}")
            return {"status": "success", "message": "تم حفظ المسح الضوئي في الأرشيف السحابي بنجاح", "filename": rel_p}
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"تعذر حفظ ملف المسح الضوئي: {str(e)}")

    # مسار بديل فقط في حال تشغيل الخادم محلياً على ويندوز مع WIA دون إرسال ملف
    if WIA_AVAILABLE and os.name == "nt":
        try:
            pythoncom.CoInitialize()
            wia = win32com.client.Dispatch("WIA.CommonDialog")
            image_file = wia.ShowAcquireImage()
            if image_file:
                safe_filename = f"{file_prefix}_scan_{int(datetime.now().timestamp())}.jpg"
                file_path = os.path.join(offender_dir, safe_filename)
                image_file.SaveFile(file_path)
                rel_p = os.path.join(offender_name, safe_filename).replace("\\", "/")
                _save_archive_record(db, case_id, rel_p, safe_filename, "jpg")
                log_audit(db, user_data.get("sub"), "SCAN", "tbl_Archive", f"مسح ضوئي مباشر وحفظ: {rel_p}")
                return {"status": "success", "message": "تم المسح الضوئي وحفظ المستند في الأرشيف وقاعدة البيانات بنجاح", "filename": rel_p}
            return {"status": "error", "message": "تم إلغاء عملية المسح الضوئي من قبل المستخدم"}
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"تعذر الاتصال بالماسح الضوئي: {str(e)}")
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    raise HTTPException(
        status_code=400,
        detail="في بيئة السحابة يتم التقاط/اختيار صورة المسح الضوئي من جهاز المستخدم ورفعها مباشرة عبر زر المسح الضوئي."
    )

@app.delete("/api/archive/file/{path:path}")
def delete_archive_file(path: str, user_data: dict = Depends(get_current_user_payload), db: Session = Depends(get_db)):
    if user_data.get("role") not in ["Admin", "Supervisor"]:
        raise HTTPException(status_code=403, detail="غير مسموح")
    clean_path = os.path.normpath(path).lstrip("\\/").replace("\\", "/")
    file_path = os.path.abspath(os.path.join(ARCHIVE_BASE_DIR, clean_path))
    base_abs = os.path.abspath(ARCHIVE_BASE_DIR)
    if not file_path.startswith(base_abs):
        raise HTTPException(status_code=403, detail="مسار غير مسموح")
    if os.path.exists(file_path):
        try:
            os.remove(file_path)
            try:
                db.query(Archive).filter(Archive.FilePath == clean_path).delete()
                db.commit()
            except Exception:
                db.rollback()
            return {"status": "success", "message": "تم حذف الملف من الأرشيف بنجاح"}
        except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))
    raise HTTPException(status_code=404, detail="الملف غير موجود")

# --- التصدير المحسّن (Excel & PDF & KML) مع تطبيق كافة فلاتر البحث بدقة ---
@app.get("/api/cases/export-excel")
def export_excel(
    q: Optional[str] = "",
    var_code: Optional[str] = "",
    case_no: Optional[str] = "",
    decision_no: Optional[str] = "",
    national_id: Optional[str] = "",
    lat: Optional[str] = "",
    lon: Optional[str] = "",
    offender: Optional[str] = "",
    address: Optional[str] = "",
    v_type: Optional[str] = "",
    status: Optional[str] = "",
    district: Optional[str] = "",
    vtype: Optional[str] = "",
    year: Optional[int] = None,
    user_data: dict = Depends(get_current_user_payload),
    db: Session = Depends(get_db)
):
    query = build_filtered_cases_query(
        db=db, q=q, var_code=var_code, case_no=case_no, decision_no=decision_no,
        national_id=national_id, lat=lat, lon=lon, offender=offender, address=address,
        v_type=v_type, status=status, district=district, vtype=vtype, year=year
    )
    results = query.yield_per(500).all()

    headers = [
        "م", "معرف السجل", "مسلسل", "كود المتغير", "تاريخ المتغير",
        "رقم المحضر", "تاريخ المحضر", "السنة", "رقم الإيقاف", "الرقم الإداري",
        "رقم القرار", "تاريخ القرار", "المساحة", "ملاحظات المحضر", "خط العرض",
        "خط الطول", "اسم المخالف", "الرقم القومي", "عنوان المخالف", "رقم الهاتف",
        "الشارع", "المنطقة / الحي", "العنوان التفصيلي", "حالة التنفيذ", "الإجراء المتخذ",
        "نوع المخالفة", "وصف المخالفة"
    ]

    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet(title="تقرير المحاضر والمخالفات")
    ws.views.sheetView[0].rightToLeft = True

    header_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    header_font = Font(name="Arial", size=11, bold=True, color="000000")
    thin_border = Border(
        left=Side(style='thin', color='A0AEC0'),
        right=Side(style='thin', color='A0AEC0'),
        top=Side(style='thin', color='A0AEC0'),
        bottom=Side(style='thin', color='A0AEC0')
    )

    header_cells = []
    for h_text in headers:
        c = openpyxl.cell.WriteOnlyCell(ws, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = thin_border
        header_cells.append(c)
    ws.append(header_cells)

    row_font = Font(name="Arial", size=10, color="000000")
    col_widths = {idx + 1: max(len(h) + 4, 13) for idx, h in enumerate(headers)}

    for idx, (r_case, r_off, r_loc, r_exec, r_judg, r_recon, r_vtype) in enumerate(results, 1):
        nat_id_str = str(int(r_off.NationalID)) if r_off and r_off.NationalID and not math.isnan(r_off.NationalID) else ""
        phone_str = str(int(r_off.Phone)) if r_off and r_off.Phone and not math.isnan(r_off.Phone) else ""

        row_values = [
            idx,
            r_case.CaseID if r_case else "",
            r_case.SerialNo if r_case else "",
            r_case.VariableCode if r_case else "",
            r_case.VariableDate if r_case else "",
            r_case.CaseNumber if r_case else "",
            r_case.CaseDate if r_case else "",
            r_case.year if r_case else "",
            r_case.StopNumber if r_case else "",
            r_case.AdminNumber if r_case else "",
            r_case.DecisionNumber if r_case else "",
            r_case.DecisionDate if r_case else "",
            r_case.AreaSize if r_case else "",
            r_case.Notes if r_case else "",
            r_case.Latitude if r_case else "",
            r_case.Longitude if r_case else "",
            r_off.OffenderName if r_off else "",
            nat_id_str,
            r_off.OffenderAddress if r_off else "",
            phone_str,
            r_loc.StreetName if r_loc else "",
            r_loc.DistrictName if r_loc else "",
            r_loc.FullAddress if r_loc else "",
            r_exec.ExecutionStatusName if r_exec else "",
            r_exec.procedure if r_exec else "",
            r_vtype.ViolationTypeName if r_vtype else "",
            r_vtype.ViolationDescription if r_vtype else ""
        ]

        row_cells = []
        for col_idx, val in enumerate(row_values, 1):
            c = openpyxl.cell.WriteOnlyCell(ws, value=val)
            c.font = row_font
            c.border = thin_border
            c.alignment = Alignment(horizontal="center", vertical="center")
            if headers[col_idx - 1] in ["الرقم القومي", "رقم الهاتف"]:
                c.number_format = '@'
            if idx <= 200:
                val_len = len(str(val or '')) + 4
                if val_len > col_widths[col_idx]:
                    col_widths[col_idx] = min(val_len, 45)
            row_cells.append(c)
        ws.append(row_cells)

    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    headers_res = {
        "Content-Disposition": f"attachment; filename=Sovereign_Violations_Report_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    }
    return StreamingResponse(output, headers=headers_res, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.get("/api/cases/export-pdf")
def export_pdf(
    q: Optional[str] = "",
    var_code: Optional[str] = "",
    case_no: Optional[str] = "",
    decision_no: Optional[str] = "",
    national_id: Optional[str] = "",
    lat: Optional[str] = "",
    lon: Optional[str] = "",
    offender: Optional[str] = "",
    address: Optional[str] = "",
    v_type: Optional[str] = "",
    status: Optional[str] = "",
    district: Optional[str] = "",
    vtype: Optional[str] = "",
    year: Optional[int] = None,
    user_data: dict = Depends(get_current_user_payload),
    db: Session = Depends(get_db)
):
    query = build_filtered_cases_query(
        db=db, q=q, var_code=var_code, case_no=case_no, decision_no=decision_no,
        national_id=national_id, lat=lat, lon=lon, offender=offender, address=address,
        v_type=v_type, status=status, district=district, vtype=vtype, year=year
    )
    results = query.yield_per(500).all()

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    story = []

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('ArabicTitle', parent=styles['Heading1'], fontName=ARABIC_FONT_NAME, alignment=1, fontSize=16, leading=20)
    cell_style = ParagraphStyle('ArabicCell', fontName=ARABIC_FONT_NAME, alignment=1, fontSize=8, leading=10)
    header_style = ParagraphStyle('ArabicHeader', fontName=ARABIC_FONT_NAME, alignment=1, fontSize=9, leading=11, textColor=colors.whitesmoke)

    title_text = ar_text("تقرير المحاضر والمخالفات - المنظومة السيادية المؤمنة")
    story.append(Paragraph(title_text, title_style))
    story.append(Spacer(1, 15))

    raw_headers = ["حالة التنفيذ", "السنة", "رقم القرار", "العنوان", "اسم المخالف", "رقم المحضر", "كود المتغير", "#"]
    headers = [Paragraph(ar_text(h), header_style) for h in raw_headers]

    table_data = [headers]

    for idx, (r_case, r_off, r_loc, r_exec, r_judg, r_recon, r_vtype) in enumerate(results, 1):
        address_text = (r_loc.FullAddress if r_loc and r_loc.FullAddress else (r_off.OffenderAddress if r_off else '-'))
        row = [
            Paragraph(ar_text(r_exec.ExecutionStatusName if r_exec else 'تنفيذ جزئي'), cell_style),
            Paragraph(ar_text(str(r_case.year) if r_case and r_case.year else '-'), cell_style),
            Paragraph(ar_text(r_case.DecisionNumber or '-'), cell_style),
            Paragraph(ar_text(address_text), cell_style),
            Paragraph(ar_text(r_off.OffenderName if r_off else 'غير مسجل'), cell_style),
            Paragraph(ar_text(r_case.CaseNumber or '-'), cell_style),
            Paragraph(ar_text(r_case.VariableCode or '-'), cell_style),
            Paragraph(ar_text(str(idx)), cell_style)
        ]
        table_data.append(row)

    t = Table(table_data, colWidths=[90, 50, 80, 220, 150, 85, 85, 30], repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E293B')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('TOPPADDING', (0, 0), (-1, 0), 6),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#F8FAFC')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#94A3B8')),
    ]))

    story.append(t)
    doc.build(story)
    buffer.seek(0)
    filename = f"Sovereign_Violations_Report_{datetime.now().strftime('%Y-%m-%d')}.pdf"
    headers_res = {"Content-Disposition": f"attachment; filename={filename}"}
    return StreamingResponse(buffer, headers=headers_res, media_type="application/pdf")

@app.get("/api/cases/export-kml")
def export_kml(
    q: Optional[str] = "",
    var_code: Optional[str] = "",
    case_no: Optional[str] = "",
    decision_no: Optional[str] = "",
    national_id: Optional[str] = "",
    lat: Optional[str] = "",
    lon: Optional[str] = "",
    offender: Optional[str] = "",
    address: Optional[str] = "",
    v_type: Optional[str] = "",
    status: Optional[str] = "",
    district: Optional[str] = "",
    vtype: Optional[str] = "",
    year: Optional[int] = None,
    user_data: dict = Depends(get_current_user_payload),
    db: Session = Depends(get_db)
):
    query = build_filtered_cases_query(
        db=db, q=q, var_code=var_code, case_no=case_no, decision_no=decision_no,
        national_id=national_id, lat=lat, lon=lon, offender=offender, address=address,
        v_type=v_type, status=status, district=district, vtype=vtype, year=year
    )
    results = query.yield_per(500).all()
    kml_items = []
    for r_case, r_off, r_loc, r_exec, r_judg, r_recon, r_vtype in results:
        if r_case.Latitude and r_case.Longitude:
            try:
                lat_val, lon_val = float(r_case.Latitude), float(r_case.Longitude)
                name = f"مخالفة: {r_off.OffenderName if r_off and r_off.OffenderName else 'غير مسجل'} - محضر: {r_case.CaseNumber or '-'}"
                kml_items.append(f"""<Placemark><name><![CDATA[{name}]]></name><Point><coordinates>{lon_val},{lat_val},0</coordinates></Point></Placemark>""")
            except Exception:
                continue
    kml_content = f"""<?xml version="1.0" encoding="UTF-8"?><kml xmlns="http://www.opengis.net/kml/2.2"><Document><name>مخالفات المنظومة السيادية</name>{''.join(kml_items)}</Document></kml>"""
    response = Response(content=kml_content, media_type="application/vnd.google-earth.kml+xml")
    response.headers["Content-Disposition"] = f"attachment; filename=Sovereign_Violations_{datetime.now().strftime('%Y-%m-%d')}.kml"
    return response

# --- الواجهة الأمامية الشاملة (سحابية + محلية) ---
@app.get("/", response_class=HTMLResponse)
def serve_dashboard():
    return """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>نظام إدارة المحاضر - المنظومة السيادية المؤمنة (JWT & RBAC)</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <script defer src="https://cdn.jsdelivr.net/npm/alpinejs@3.x.x/dist/cdn.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <style>
        body { font-family: 'Cairo', 'Segoe UI', Tahoma, Arial, sans-serif !important; }
        #map { height: 340px; width: 100%; border-radius: 0.75rem; z-index: 10; }
        [x-cloak] { display: none !important; }
    </style>
</head>
<body class="bg-slate-900 text-slate-100 font-sans antialiased select-text" x-data="appData()">

    <!-- شاشة تسجيل الدخول السيادية -->
    <div x-show="!isLoggedIn" class="fixed inset-0 bg-slate-950 z-50 flex items-center justify-center p-4">
        <div class="bg-slate-900 border border-slate-700 w-full max-w-md p-8 rounded-3xl shadow-2xl space-y-6">
            <div class="text-center space-y-2">
                <div class="bg-blue-600/20 text-blue-400 w-16 h-16 mx-auto rounded-2xl flex items-center justify-center border border-blue-500/30 text-2xl"><i class="fa-solid fa-shield-halved"></i></div>
                <h2 class="text-xl font-black text-white">البوابة السيادية المؤمنة</h2>
                <p class="text-xs text-slate-400">أدخل بيانات الاعتماد الرسمية للمتابعة</p>
            </div>
            <div x-show="loginError" class="bg-rose-950/70 border border-rose-500/40 text-rose-200 px-4 py-3 rounded-xl text-xs font-semibold text-center" x-text="loginError"></div>
            <form @submit.prevent="login()" class="space-y-4">
                <div>
                    <label class="block text-xs text-slate-400 mb-1">اسم المستخدم</label>
                    <input type="text" x-model="loginForm.username" required autocomplete="username" placeholder="أدخل اسم المستخدم" class="w-full bg-slate-950 border border-slate-700 rounded-xl px-4 py-3 text-sm text-white outline-none focus:border-blue-500">
                </div>
                <div>
                    <label class="block text-xs text-slate-400 mb-1">كلمة المرور</label>
                    <input type="password" x-model="loginForm.password" required autocomplete="current-password" placeholder="••••••••" class="w-full bg-slate-950 border border-slate-700 rounded-xl px-4 py-3 text-sm text-white outline-none focus:border-blue-500">
                </div>
                <button type="submit" :disabled="isLoggingIn" class="w-full bg-blue-600 hover:bg-blue-500 disabled:opacity-50 text-white py-3 rounded-xl font-bold text-sm shadow-lg transition cursor-pointer">
                    <span x-text="isLoggingIn ? 'جاري التحقق...' : 'مصادقة ودخول آمن'"></span>
                </button>
            </form>
            <div class="text-center text-xs text-slate-500 font-mono">حماية متطورة بنظام الرموز المشفرة (JWT)</div>
        </div>
    </div>

    <!-- Header السيادي -->
    <header class="bg-slate-800 border-b border-slate-700 px-6 py-4 flex flex-col lg:flex-row justify-between items-center gap-4 shadow-xl sticky top-0 z-40">
        <div class="flex items-center gap-3 w-full lg:w-auto justify-between lg:justify-start">
            <div class="flex items-center gap-3">
                <div class="bg-blue-600 p-2.5 rounded-xl shadow-md text-white"><i class="fa-solid fa-building-shield text-xl"></i></div>
                <div>
                    <h1 class="text-base sm:text-lg font-bold tracking-wide">نظام إدارة المحاضر والمخالفات (المنظومة السيادية المؤمنة)</h1>
                    <p class="text-xs text-slate-400">المستخدم النشط: <span class="text-emerald-400 font-bold" x-text="currentUser"></span> (<span class="text-amber-400" x-text="currentRole"></span>)</p>
                </div>
            </div>
        </div>

        <div class="flex items-center gap-2.5 flex-wrap justify-end w-full lg:w-auto">
            <button @click="openModal()" class="bg-blue-600 hover:bg-blue-500 px-4 py-2 rounded-xl text-xs sm:text-sm font-bold transition flex items-center gap-2 shadow cursor-pointer text-white"><i class="fa-solid fa-plus"></i> إضافة محضر</button>

            <button @click="loadData(currentPage); loadStats();" class="bg-emerald-600 hover:bg-emerald-500 px-4 py-2 rounded-xl text-sm font-semibold shadow cursor-pointer flex items-center gap-2 text-white">
                <svg xmlns="http://www.w3.org/2000/svg" class="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
                </svg>
                <span>تحديث البيانات</span>
            </button>

            <template x-if="currentRole === 'Admin'">
                <div class="flex items-center gap-2">
                    <button @click="openBackupModal()" class="bg-emerald-600/20 hover:bg-emerald-600 text-emerald-300 hover:text-white px-3.5 py-2 rounded-xl text-xs sm:text-sm font-semibold transition border border-emerald-500/30 flex items-center gap-1.5 cursor-pointer shadow"><i class="fa-solid fa-database"></i> النسخة الاحتياطية</button>
                    <button @click="openUsersModal()" class="bg-indigo-600/20 hover:bg-indigo-600 text-indigo-300 hover:text-white px-3.5 py-2 rounded-xl text-xs sm:text-sm font-semibold transition border border-indigo-500/30 flex items-center gap-1.5 cursor-pointer shadow"><i class="fa-solid fa-users-gear"></i> إدارة المستخدمين</button>
                    <button @click="openAuditModal()" class="bg-purple-600/20 hover:bg-purple-600 text-purple-300 hover:text-white px-3.5 py-2 rounded-xl text-xs sm:text-sm font-semibold transition border border-purple-500/30 flex items-center gap-1.5 cursor-pointer shadow"><i class="fa-solid fa-user-shield"></i> سجل التدقيق</button>
                </div>
            </template>

            <div class="flex items-center gap-1.5 bg-slate-900/60 p-1 rounded-xl border border-slate-700">
                <button @click="exportToExcel()" :disabled="isExporting" class="bg-emerald-600/20 hover:bg-emerald-600 disabled:opacity-50 text-emerald-300 hover:text-white px-3 py-1.5 rounded-lg text-xs font-semibold transition flex items-center gap-1.5 cursor-pointer"><i class="fa-solid fa-file-excel"></i> إكسل</button>
                <button @click="exportToPDF()" :disabled="isExporting" class="bg-amber-600/20 hover:bg-amber-600 disabled:opacity-50 text-amber-300 hover:text-white px-3 py-1.5 rounded-lg text-xs font-semibold transition flex items-center gap-1.5 cursor-pointer"><i class="fa-solid fa-file-pdf"></i> PDF</button>
                <button @click="printFilteredReport()" :disabled="isExporting" class="bg-cyan-600/20 hover:bg-cyan-600 disabled:opacity-50 text-cyan-300 hover:text-white px-3 py-1.5 rounded-lg text-xs font-semibold transition flex items-center gap-1.5 cursor-pointer"><i class="fa-solid fa-print"></i> طباعة</button>
                <button @click="exportToKML()" :disabled="isExporting" class="bg-indigo-600/20 hover:bg-indigo-600 disabled:opacity-50 text-indigo-300 hover:text-white px-3 py-1.5 rounded-lg text-xs font-semibold transition flex items-center gap-1.5 cursor-pointer"><i class="fa-solid fa-globe"></i> KML</button>
            </div>

            <button @click="logout()" class="bg-rose-600/20 hover:bg-rose-600 text-rose-300 hover:text-white px-3.5 py-2 rounded-xl text-xs sm:text-sm font-semibold transition border border-rose-500/30 flex items-center gap-1.5 cursor-pointer shadow"><i class="fa-solid fa-right-from-bracket"></i> خروج آمن</button>
        </div>
    </header>

    <main class="p-6 max-w-[1700px] mx-auto space-y-6">
        <!-- الإحصائيات -->
        <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
            <div class="bg-slate-800 p-4 rounded-2xl border border-slate-700 shadow-md flex items-center justify-between">
                <div><p class="text-xs text-slate-400">إجمالي السجلات (PostgreSQL)</p><h3 class="text-2xl font-black text-blue-400 mt-1" x-text="stats.total_cases ? stats.total_cases.toLocaleString() : '0'"></h3></div>
                <div class="bg-blue-600/20 text-blue-400 p-3 rounded-xl border border-blue-500/30"><i class="fa-solid fa-database text-xl"></i></div>
            </div>
            <div class="bg-slate-800 p-4 rounded-2xl border border-slate-700 shadow-md flex items-center justify-between">
                <div><p class="text-xs text-slate-400">النتائج المفلترة</p><h3 class="text-2xl font-black text-emerald-400 mt-1" x-text="totalRecords.toLocaleString()"></h3></div>
                <div class="bg-emerald-600/20 text-emerald-400 p-3 rounded-xl border border-emerald-500/30"><i class="fa-solid fa-filter-circle-dollar text-xl"></i></div>
            </div>
            <div class="bg-slate-800 p-4 rounded-2xl border border-slate-700 shadow-md flex items-center justify-between">
                <div><p class="text-xs text-slate-400">الحالة النشطة</p><h3 class="text-base font-bold text-amber-400 mt-1" x-text="statusFilter ? statusFilter : 'الكل'"></h3></div>
                <div class="bg-amber-600/20 text-amber-400 p-3 rounded-xl border border-amber-500/30"><i class="fa-solid fa-bars-progress text-xl"></i></div>
            </div>
            <div class="bg-slate-800 p-4 rounded-2xl border border-slate-700 shadow-md flex items-center justify-between">
                <div><p class="text-xs text-slate-400">السجلات المكودة</p><h3 class="text-2xl font-black text-purple-400 mt-1" x-text="stats.variable_coded_count ? stats.variable_coded_count.toLocaleString() : '0'"></h3></div>
                <div class="bg-purple-600/20 text-purple-400 p-3 rounded-xl border border-purple-500/30"><i class="fa-solid fa-barcode text-xl"></i></div>
            </div>
        </div>

        <!-- الخريطة والرسوم البيانية -->
        <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
            <div class="bg-slate-800 p-5 rounded-2xl border border-slate-700 shadow-xl lg:col-span-1 flex flex-col justify-between">
                <div class="flex justify-between items-center mb-2">
                    <h3 class="text-sm font-bold text-slate-300"><i class="fa-solid fa-chart-pie text-blue-400"></i> حالات التنفيذ</h3>
                    <button x-show="statusFilter" @click="statusFilter=''; loadData(1);" class="text-xs text-rose-400 hover:underline cursor-pointer">إلغاء الفلتر</button>
                </div>
                <div class="relative w-full h-[240px] flex justify-center items-center"><canvas id="statusChart"></canvas></div>
            </div>
            <div class="bg-slate-800 p-5 rounded-2xl border border-slate-700 shadow-xl lg:col-span-2 flex flex-col justify-between">
                <div class="flex justify-between items-center mb-3">
                    <h3 class="text-sm font-bold text-slate-300"><i class="fa-solid fa-satellite text-emerald-400"></i> الخريطة الفضائية الحية</h3>
                    <div class="flex gap-2">
                        <button @click="switchMapLayer('satellite')" :class="{'bg-emerald-600 text-white': currentLayer==='satellite', 'bg-slate-700 text-slate-300': currentLayer!=='satellite'}" class="px-3 py-1 rounded-lg text-xs font-bold transition cursor-pointer">أقمار صناعية 🛰️</button>
                        <button @click="switchMapLayer('streets')" :class="{'bg-blue-600 text-white': currentLayer==='streets', 'bg-slate-700 text-slate-300': currentLayer!=='streets'}" class="px-3 py-1 rounded-lg text-xs font-bold transition cursor-pointer">شوارع 🗺️</button>
                    </div>
                </div>
                <div id="map" class="relative bg-slate-950 border border-slate-700/80 flex items-center justify-center">
                    <div x-show="mapOfflineNotice" class="absolute bottom-2 right-2 z-[500] bg-slate-900/90 border border-amber-500/40 text-amber-300 px-3 py-1.5 rounded-lg text-[11px] pointer-events-none">
                        وضع بدون إنترنت: بلاطات الخريطة تتطلب اتصالاً بالإنترنت، بينما الإحداثيات والبحث والتصدير تعمل محلياً.
                    </div>
                </div>
            </div>
        </div>

        <!-- أدوات البحث المتقدم الشاملة -->
        <div class="bg-slate-800 p-4 rounded-2xl border border-slate-700 space-y-4 shadow-md">
            <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
                <div>
                    <label class="block text-[11px] text-amber-400 mb-1 font-bold">بحث دقيق: كود المتغير</label>
                    <input type="text" x-model="searchVarCode" @input.debounce.350ms="loadData(1)" placeholder="بحث دقيق بكود المتغير..." class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-xs outline-none focus:border-blue-500 text-white font-mono">
                </div>
                <div>
                    <label class="block text-[11px] text-blue-400 mb-1 font-bold">بحث دقيق: رقم المحضر</label>
                    <input type="text" x-model="searchCaseNo" @input.debounce.350ms="loadData(1)" placeholder="بحث دقيق برقم المحضر..." class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-xs outline-none focus:border-blue-500 text-white font-mono">
                </div>
                <div>
                    <label class="block text-[11px] text-purple-400 mb-1 font-bold">بحث دقيق: قرار المحافظ</label>
                    <input type="text" x-model="searchDecisionNo" @input.debounce.350ms="loadData(1)" placeholder="بحث برقم قرار المحافظ..." class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-xs outline-none focus:border-blue-500 text-white font-mono">
                </div>
                <div>
                    <label class="block text-[11px] text-indigo-400 mb-1 font-bold">بحث دقيق: الرقم القومي للمخالف</label>
                    <input type="text" x-model="searchNationalID" @input.debounce.350ms="loadData(1)" placeholder="بحث دقيق بالرقم القومي..." class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-xs outline-none focus:border-blue-500 text-white font-mono">
                </div>
            </div>

            <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-6 gap-3">
                <div>
                    <input type="text" x-model="searchOffender" @input.debounce.350ms="loadData(1)" placeholder="ابحث باسم المخالف..." class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-xs outline-none focus:border-blue-500 text-white">
                </div>
                <div>
                    <input type="text" x-model="searchAddress" @input.debounce.350ms="loadData(1)" placeholder="ابحث بالعنوان أو الشارع..." class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-xs outline-none focus:border-blue-500 text-white">
                </div>
                <div>
                    <input type="text" x-model="searchLat" @input.debounce.350ms="loadData(1)" placeholder="خط العرض (Latitude)..." class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-xs outline-none focus:border-blue-500 text-white font-mono">
                </div>
                <div>
                    <input type="text" x-model="searchLon" @input.debounce.350ms="loadData(1)" placeholder="خط الطول (Longitude)..." class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-xs outline-none focus:border-blue-500 text-white font-mono">
                </div>
                <select x-model="districtFilter" @change="loadData(1)" class="bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-xs outline-none focus:border-blue-500 text-slate-200">
                    <option value="">جميع المناطق والأحياء</option>
                    <template x-for="d in defaultDistricts" :key="d">
                        <option :value="d" x-text="d"></option>
                    </template>
                </select>
                <select x-model="statusFilter" @change="loadData(1)" class="bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-xs outline-none focus:border-blue-500 text-slate-200">
                    <option value="">جميع حالات التنفيذ</option>
                    <option value="تنفيذ جزئي">تنفيذ جزئي</option>
                    <option value="تنفيذ كلي">تنفيذ كلي</option>
                    <option value="موقوف">موقوف</option>
                    <option value="قيد المراجعة">قيد المراجعة</option>
                    <option value="تم التنفيذ">تم التنفيذ</option>
                    <option value="لم يتم التنفيذ">لم يتم التنفيذ</option>
                    <option value="مأهول">مأهول</option>
                    <option value="دراسة امنية">دراسة امنية</option>
                    <option value="طلب تصالح">طلب تصالح</option>
                </select>
            </div>

            <div class="flex flex-wrap gap-3 items-center justify-between pt-2 border-t border-slate-700/60">
                <div class="relative flex-1 min-w-[280px]">
                    <span class="absolute inset-y-0 right-0 flex items-center pr-3 text-slate-400"><i class="fa-solid fa-magnifying-glass"></i></span>
                    <input type="text" x-model="searchQuery" @input.debounce.350ms="loadData(1)" placeholder="بحث عام سريع في كافة الحقول..." class="w-full bg-slate-900 border border-slate-700 rounded-xl pr-10 pl-4 py-2 text-xs outline-none focus:border-blue-500 transition text-white">
                </div>
                <div class="flex items-center gap-2">
                    <select x-model="vtypeFilter" @change="loadData(1)" class="bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-xs outline-none focus:border-blue-500 text-slate-200">
                        <option value="">جميع أنواع المخالفات</option>
                        <template x-for="vt in defaultViolationTypes" :key="vt">
                            <option :value="vt" x-text="vt"></option>
                        </template>
                    </select>
                    <input type="number" x-model.number="yearFilter" @input.debounce.350ms="loadData(1)" placeholder="السنة" class="w-20 bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-xs outline-none focus:border-blue-500 text-slate-200">
                    <button @click="searchVarCode=''; searchCaseNo=''; searchDecisionNo=''; searchNationalID=''; searchLat=''; searchLon=''; searchOffender=''; searchAddress=''; searchQuery=''; statusFilter=''; districtFilter=''; vtypeFilter=''; yearFilter=''; loadData(1);" class="bg-slate-700 hover:bg-slate-600 text-slate-300 px-4 py-2 rounded-xl text-xs transition border border-slate-600 cursor-pointer">مسح جميع الفلاتر</button>
                </div>
            </div>
        </div>

        <!-- جدول العرض الرئيسي -->
        <div class="bg-slate-800 rounded-2xl border border-slate-700 overflow-hidden shadow-xl">
            <div class="overflow-x-auto">
                <table class="w-full text-right border-collapse">
                    <thead>
                        <tr class="bg-slate-800 border-b border-slate-700 text-slate-400 text-xs uppercase">
                            <th class="p-4">#</th>
                            <th class="p-4">كود المتغير</th>
                            <th class="p-4">رقم المحضر</th>
                            <th class="p-4">رقم القرار</th>
                            <th class="p-4">اسم المخالف والرقم القومي</th>
                            <th class="p-4">العنوان (مع الإحداثيات)</th>
                            <th class="p-4">نوع المخالفة</th>
                            <th class="p-4">السنة</th>
                            <th class="p-4">حالة التنفيذ</th>
                            <th class="p-4 text-center">الإجراءات</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-700/50 text-sm">
                        <template x-for="(row, index) in tableData" :key="row.CaseID">
                            <tr class="hover:bg-slate-700/30 transition">
                                <td class="p-4 text-slate-400" x-text="((currentPage - 1) * 50) + index + 1"></td>
                                <td class="p-4 font-bold text-amber-400 select-all" x-text="row.VariableCode"></td>
                                <td class="p-4 font-bold text-blue-400 select-all" x-text="row.CaseNumber"></td>
                                <td class="p-4 font-bold text-amber-300 select-all" x-text="row.DecisionNumber"></td>
                                <td class="p-4 select-all">
                                    <div class="font-semibold text-slate-100" x-text="row.OffenderName"></div>
                                    <div class="text-xs text-indigo-400 font-mono mt-0.5" x-show="row.OffenderNationalID" x-text="'القومي: ' + row.OffenderNationalID"></div>
                                </td>
                                <td class="p-4 text-slate-300 select-all">
                                    <div x-text="row.FullAddress || row.OffenderAddress || '-'"></div>
                                    <div class="text-xs font-mono text-emerald-400 mt-1" x-show="row.Latitude && row.Longitude" x-text="'📌 ' + row.Latitude + ', ' + row.Longitude"></div>
                                </td>
                                <td class="p-4 font-semibold text-purple-300 select-all" x-text="row.ViolationTypeName || '-'"></td>
                                <td class="p-4 text-cyan-400" x-text="row.year"></td>
                                <td class="p-4"><span class="px-3 py-1 rounded-full text-xs font-semibold bg-emerald-500/20 text-emerald-400 border border-emerald-500/30" x-text="row.ExecutionStatusName"></span></td>
                                <td class="p-4 text-center">
                                    <div class="flex justify-center gap-1.5">
                                        <button @click="focusOnMap(row)" class="bg-emerald-600/20 hover:bg-emerald-600 text-emerald-400 hover:text-white px-2.5 py-1 rounded text-xs font-semibold transition border border-emerald-500/30 cursor-pointer" title="عرض بالخريطة"><i class="fa-solid fa-map-pin"></i></button>
                                        <button @click="editCase(row)" class="bg-amber-600/20 hover:bg-amber-600 text-amber-400 hover:text-white px-2 py-1 rounded text-xs font-semibold transition border border-amber-500/30 cursor-pointer" title="تعديل"><i class="fa-solid fa-pen-to-square"></i></button>
                                        <button @click="deleteCase(row.CaseID)" class="bg-rose-600/20 hover:bg-rose-600 text-rose-400 hover:text-white px-2 py-1 rounded text-xs font-semibold transition border border-rose-500/30 cursor-pointer" title="حذف"><i class="fa-solid fa-trash"></i></button>
                                    </div>
                                </td>
                            </tr>
                        </template>
                    </tbody>
                </table>
            </div>

            <div class="bg-slate-800/90 border-t border-slate-700 px-6 py-4 flex justify-between items-center flex-wrap gap-4">
                <div class="text-xs text-slate-400">
                    عرض الصفحة <span class="text-white font-bold" x-text="currentPage"></span> من <span class="text-white font-bold" x-text="totalPages"></span> (إجمالي السجلات: <span class="text-emerald-400 font-bold" x-text="totalRecords"></span>)
                </div>
                <div class="flex gap-2">
                    <button @click="loadData(1)" :disabled="currentPage === 1" class="bg-slate-700 hover:bg-slate-600 px-3.5 py-2 rounded-xl text-xs font-bold disabled:opacity-40 cursor-pointer transition">« الأولى</button>
                    <button @click="loadData(currentPage - 1)" :disabled="currentPage === 1" class="bg-slate-700 hover:bg-slate-600 px-3.5 py-2 rounded-xl text-xs font-bold disabled:opacity-40 cursor-pointer transition">‹ السابق</button>
                    <button @click="loadData(currentPage + 1)" :disabled="currentPage >= totalPages" class="bg-slate-700 hover:bg-slate-600 px-3.5 py-2 rounded-xl text-xs font-bold disabled:opacity-40 cursor-pointer transition">التالي ›</button>
                    <button @click="loadData(totalPages)" :disabled="currentPage >= totalPages" class="bg-slate-700 hover:bg-slate-600 px-3.5 py-2 rounded-xl text-xs font-bold disabled:opacity-40 cursor-pointer transition">الأخيرة »</button>
                </div>
            </div>
        </div>
    </main>

    <!-- Modal المحاضر الشامل بالتبويبات -->
    <div x-show="showModal" x-cloak class="fixed inset-0 bg-black/85 backdrop-blur-sm z-50 flex items-center justify-center p-2 sm:p-4 overflow-hidden">
        <div class="bg-slate-800 border border-slate-700 w-full max-w-7xl h-[92vh] rounded-3xl shadow-2xl overflow-hidden flex flex-col" @click.away="closeModal()">

            <div class="bg-slate-800 border-b border-slate-700 px-6 py-3 flex justify-between items-center shrink-0">
                <h3 class="text-base sm:text-lg font-bold text-white flex items-center gap-2">
                    <i class="fa-solid fa-folder-tree text-blue-500"></i>
                    <span x-text="isEditMode ? 'تعديل المحضر والبيانات الشاملة' : 'إضافة محضر جديد (ترقيم تلقائي SerialNo)'"></span>
                </h3>
                <button @click="closeModal()" class="text-slate-400 hover:text-white text-xl cursor-pointer"><i class="fa-solid fa-xmark"></i></button>
            </div>

            <div class="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-8 bg-slate-900 border-b border-slate-700 p-2 gap-2 text-center shrink-0">
                <button @click="activeTab = 'cases'" :class="{'bg-blue-600 text-white shadow-lg': activeTab==='cases', 'bg-slate-800 text-slate-300 hover:bg-slate-700': activeTab!=='cases'}" class="py-2 px-2 rounded-xl text-xs font-bold transition cursor-pointer truncate">جدول المحاضر</button>
                <button @click="activeTab = 'offenders'" :class="{'bg-blue-600 text-white shadow-lg': activeTab==='offenders', 'bg-slate-800 text-slate-300 hover:bg-slate-700': activeTab!=='offenders'}" class="py-2 px-2 rounded-xl text-xs font-bold transition cursor-pointer truncate">المخالفين</button>
                <button @click="activeTab = 'violationtypes'" :class="{'bg-blue-600 text-white shadow-lg': activeTab==='violationtypes', 'bg-slate-800 text-slate-300 hover:bg-slate-700': activeTab!=='violationtypes'}" class="py-2 px-2 rounded-xl text-xs font-bold transition cursor-pointer truncate">نوع المخالفة</button>
                <button @click="activeTab = 'locations'" :class="{'bg-blue-600 text-white shadow-lg': activeTab==='locations', 'bg-slate-800 text-slate-300 hover:bg-slate-700': activeTab!=='locations'}" class="py-2 px-2 rounded-xl text-xs font-bold transition cursor-pointer truncate">المواقع</button>
                <button @click="activeTab = 'execution'" :class="{'bg-blue-600 text-white shadow-lg': activeTab==='execution', 'bg-slate-800 text-slate-300 hover:bg-slate-700': activeTab!=='execution'}" class="py-2 px-2 rounded-xl text-xs font-bold transition cursor-pointer truncate">التنفيذ</button>
                <button @click="activeTab = 'judgements'" :class="{'bg-blue-600 text-white shadow-lg': activeTab==='judgements', 'bg-slate-800 text-slate-300 hover:bg-slate-700': activeTab!=='judgements'}" class="py-2 px-2 rounded-xl text-xs font-bold transition cursor-pointer truncate">تنفيذ الأحكام</button>
                <button @click="activeTab = 'reconciliation'" :class="{'bg-blue-600 text-white shadow-lg': activeTab==='reconciliation', 'bg-slate-800 text-slate-300 hover:bg-slate-700': activeTab!=='reconciliation'}" class="py-2 px-2 rounded-xl text-xs font-bold transition cursor-pointer truncate">التصالح</button>
                <button @click="activeTab = 'archive'" :class="{'bg-blue-600 text-white shadow-lg': activeTab==='archive', 'bg-slate-800 text-slate-300 hover:bg-slate-700': activeTab!=='archive'}" class="py-2 px-2 rounded-xl text-xs font-bold transition cursor-pointer truncate">الأرشيف السحابي 📂</button>
            </div>

            <div class="p-4 sm:p-6 overflow-y-auto flex-1 space-y-6">
                <div x-show="activeTab === 'cases'" class="space-y-4">
                    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                        <div><label class="block text-xs text-slate-400 mb-1">مسلسل تلقائي (SerialNo)</label><input type="number" x-model.number="form.SerialNo" class="w-full bg-slate-950 border border-emerald-500/50 rounded-xl px-3 py-2.5 text-sm font-bold text-emerald-400 outline-none"></div>
                        <div><label class="block text-xs text-slate-400 mb-1">كود المتغير (VariableCode)</label><input type="text" x-model="form.VariableCode" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm outline-none"></div>
                        <div><label class="block text-xs text-slate-400 mb-1">تاريخ المتغير (VariableDate)</label><input type="text" x-model="form.VariableDate" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm outline-none"></div>
                        <div><label class="block text-xs text-slate-400 mb-1">رقم المحضر (CaseNumber)</label><input type="text" x-model="form.CaseNumber" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm outline-none"></div>
                        <div><label class="block text-xs text-slate-400 mb-1">تاريخ المحضر (CaseDate)</label><input type="text" x-model="form.CaseDate" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm outline-none"></div>
                        <div><label class="block text-xs text-slate-400 mb-1">السنة (year)</label><input type="number" x-model.number="form.year" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm outline-none"></div>
                        <div><label class="block text-xs text-slate-400 mb-1">رقم الإيقاف (StopNumber)</label><input type="text" x-model="form.StopNumber" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm outline-none"></div>
                        <div><label class="block text-xs text-slate-400 mb-1">الرقم الإداري (AdminNumber)</label><input type="text" x-model="form.AdminNumber" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm outline-none"></div>
                        <div><label class="block text-xs text-slate-400 mb-1">رقم القرار (DecisionNumber)</label><input type="text" x-model="form.DecisionNumber" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm outline-none"></div>
                        <div><label class="block text-xs text-slate-400 mb-1">تاريخ القرار (DecisionDate)</label><input type="text" x-model="form.DecisionDate" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm outline-none"></div>
                        <div><label class="block text-xs text-slate-400 mb-1">مساحة الأرض (AreaSize)</label><input type="number" step="any" x-model.number="form.AreaSize" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm outline-none"></div>
                    </div>
                    <div class="bg-slate-900 p-4 rounded-2xl border border-slate-700 space-y-3">
                        <div class="flex justify-between items-center">
                            <span class="text-xs font-bold text-emerald-400"><i class="fa-solid fa-location-crosshairs"></i> الإحداثيات الجغرافية (Latitude & Longitude)</span>
                            <button type="button" @click="captureCurrentLocation()" class="bg-emerald-600 hover:bg-emerald-500 text-white px-3.5 py-1.5 rounded-xl text-xs font-bold transition cursor-pointer">📍 التقاط موقعي الحالي (GPS)</button>
                        </div>
                        <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                            <div><label class="block text-xs text-slate-400 mb-1">خط العرض (Latitude - طولي)</label><input type="text" x-model="form.Latitude" class="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2.5 text-sm font-mono"></div>
                            <div><label class="block text-xs text-slate-400 mb-1">خط الطول (Longitude - عرضي)</label><input type="text" x-model="form.Longitude" class="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2.5 text-sm font-mono"></div>
                        </div>
                    </div>
                    <div><label class="block text-xs text-slate-400 mb-1">ملاحظات المحضر (Notes)</label><textarea x-model="form.Notes" rows="2" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm outline-none"></textarea></div>
                </div>

                <div x-show="activeTab === 'offenders'" class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div><label class="block text-xs text-slate-400 mb-1">اسم المخالف (OffenderName)</label><input type="text" x-model="form.OffenderName" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                    <div><label class="block text-xs text-slate-400 mb-1">الرقم القومي (NationalID)</label><input type="number" step="any" x-model.number="form.OffenderNationalID" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                    <div><label class="block text-xs text-slate-400 mb-1">رقم الهاتف (Phone)</label><input type="number" step="any" x-model.number="form.Phone" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                    <div class="sm:col-span-2"><label class="block text-xs text-slate-400 mb-1">عنوان المخالف (OffenderAddress)</label><input type="text" x-model="form.OffenderAddress" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                    <div class="sm:col-span-2"><label class="block text-xs text-slate-400 mb-1">ملاحظات المخالف (Notes)</label><textarea x-model="form.OffenderNotes" rows="2" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></textarea></div>
                </div>

                <div x-show="activeTab === 'violationtypes'" class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                        <label class="block text-xs text-slate-400 mb-1">اختر نوع المخالفة من القائمة المعيارية</label>
                        <select x-model="form.ViolationTypeName" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm text-white outline-none focus:border-blue-500">
                            <option value="">-- اختر نوع المخالفة --</option>
                            <template x-for="vt in defaultViolationTypes" :key="vt">
                                <option :value="vt" x-text="vt"></option>
                            </template>
                        </select>
                    </div>
                    <div>
                        <label class="block text-xs text-slate-400 mb-1">أو اكتب نوع المخالفة يدوياً</label>
                        <input type="text" x-model="form.ViolationTypeName" placeholder="اكتب نوع المخالفة..." class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm outline-none focus:border-blue-500">
                    </div>
                    <div class="sm:col-span-2">
                        <label class="block text-xs text-slate-400 mb-1">وصف المخالفة التفصيلي (ViolationDescription)</label>
                        <textarea x-model="form.ViolationDescription" rows="3" placeholder="أدخل تفاصيل ووصف المخالفة..." class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm outline-none focus:border-blue-500"></textarea>
                    </div>
                </div>

                <div x-show="activeTab === 'locations'" class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div><label class="block text-xs text-slate-400 mb-1">اسم الشارع (StreetName)</label><input type="text" x-model="form.StreetName" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                    <div>
                        <label class="block text-xs text-slate-400 mb-1">اسم المنطقة / الحي (DistrictName)</label>
                        <select x-model="form.DistrictName" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm text-white outline-none focus:border-blue-500">
                            <option value="">-- اختر المنطقة / الحي --</option>
                            <template x-for="d in defaultDistricts" :key="d">
                                <option :value="d" x-text="d"></option>
                            </template>
                        </select>
                    </div>
                    <div><label class="block text-xs text-slate-400 mb-1">العنوان التفصيلي (FullAddress)</label><input type="text" x-model="form.FullAddress" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                    <div><label class="block text-xs text-slate-400 mb-1">المعلم المميز (Landmark)</label><input type="text" x-model="form.Landmark" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                </div>

                <div x-show="activeTab === 'execution'" class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                        <label class="block text-xs text-slate-400 mb-1">حالة التنفيذ (ExecutionStatusName)</label>
                        <select x-model="form.ExecutionStatusName" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm text-white">
                            <option value="تنفيذ جزئي">تنفيذ جزئي</option>
                            <option value="تنفيذ كلي">تنفيذ كلي</option>
                            <option value="موقوف">موقوف</option>
                            <option value="قيد المراجعة">قيد المراجعة</option>
                            <option value="تم التنفيذ">تم التنفيذ</option>
                            <option value="لم يتم التنفيذ">لم يتم التنفيذ</option>
                            <option value="مأهول">مأهول</option>
                            <option value="دراسة امنية">دراسة امنية</option>
                            <option value="طلب تصالح">طلب تصالح</option>
                        </select>
                    </div>
                    <div><label class="block text-xs text-slate-400 mb-1">الإجراء (procedure)</label><input type="text" x-model="form.procedure" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                    <div><label class="block text-xs text-slate-400 mb-1">مستوى التنفيذ (ExecutionLevel)</label><input type="text" x-model="form.ExecutionLevel" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                    <div class="sm:col-span-2"><label class="block text-xs text-slate-400 mb-1">ملاحظات التنفيذ (Notes)</label><textarea x-model="form.ExecNotes" rows="2" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></textarea></div>
                </div>

                <div x-show="activeTab === 'judgements'" class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                    <div>
                        <label class="block text-xs text-slate-400 mb-1">يوجد استئناف (HasAppeal)</label>
                        <select x-model="form.HasAppeal" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm text-white">
                            <option :value="false">لا</option>
                            <option :value="true">نعم</option>
                        </select>
                    </div>
                    <div><label class="block text-xs text-slate-400 mb-1">قرار المحكمة (CourtDecision)</label><input type="text" x-model="form.CourtDecision" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                    <div><label class="block text-xs text-slate-400 mb-1">تاريخ قرار المحكمة (CourtDecisionDate)</label><input type="number" step="any" x-model.number="form.CourtDecisionDate" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                    <div><label class="block text-xs text-slate-400 mb-1">رقم قرار المحافظ (Deputy Governor Decision)</label><input type="number" x-model.number="form.DeputyGovernorDecision" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                    <div><label class="block text-xs text-slate-400 mb-1">تاريخ قرار المحافظ (Date Deputy Governor's Decision)</label><input type="number" step="any" x-model.number="form.DateDeputyGovernorDecision" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                    <div><label class="block text-xs text-slate-400 mb-1">الحكم النهائي (FinalJudgement)</label><input type="number" step="any" x-model.number="form.FinalJudgement" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                    <div><label class="block text-xs text-slate-400 mb-1">تاريخ الحكم النهائي (FinalJudgementDate)</label><input type="number" step="any" x-model.number="form.FinalJudgementDate" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                    <div><label class="block text-xs text-slate-400 mb-1">الغرامة (The fine)</label><input type="number" step="any" x-model.number="form.TheFine" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                    <div><label class="block text-xs text-slate-400 mb-1">رقم عدم التنفيذ (NonExecutionNumber)</label><input type="number" x-model.number="form.NonExecutionNumber" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                    <div><label class="block text-xs text-slate-400 mb-1">تاريخ صدور الإدارة (Department issue date)</label><input type="number" step="any" x-model.number="form.DepartmentIssueDate" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                    <div><label class="block text-xs text-slate-400 mb-1">رقم الجناية (FelonyNumber)</label><input type="number" step="any" x-model.number="form.FelonyNumber" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                    <div class="sm:col-span-2 lg:col-span-3"><label class="block text-xs text-slate-400 mb-1">ملاحظات تنفيذ الأحكام (Notes)</label><textarea x-model="form.JudgementNotes" rows="2" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></textarea></div>
                </div>

                <div x-show="activeTab === 'reconciliation'" class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div><label class="block text-xs text-slate-400 mb-1">رقم التصالح (ReconciliationNumber)</label><input type="text" x-model="form.ReconciliationNumber" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                    <div><label class="block text-xs text-slate-400 mb-1">تاريخ التصالح (ReconciliationDate)</label><input type="text" x-model="form.ReconciliationDate" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                    <div><label class="block text-xs text-slate-400 mb-1">مبلغ التصالح (FineAmount)</label><input type="number" step="any" x-model.number="form.FineAmount" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></div>
                    <div class="sm:col-span-2"><label class="block text-xs text-slate-400 mb-1">ملاحظات التصالح (Notes)</label><textarea x-model="form.ReconNotes" rows="2" class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm"></textarea></div>
                </div>

                <div x-show="activeTab === 'archive'" class="space-y-4">
                    <div class="bg-slate-900 p-4 sm:p-6 rounded-2xl border border-slate-700 text-center space-y-4">
                        <div class="flex justify-between items-center border-b border-slate-800 pb-3 flex-wrap gap-2">
                            <span class="text-sm font-bold text-amber-400 flex items-center gap-2">
                                <i class="fa-solid fa-box-archive"></i> الأرشيف السحابي للمستندات
                            </span>
                            <span class="bg-slate-800 px-3 py-1 rounded-full text-xs font-mono text-slate-300"
                                  x-text="'المستند ' + (archiveList.length > 0 && !archiveList[0].includes('لا توجد') ? archiveIndex + 1 : 0) + ' من ' + (archiveList[0] && archiveList[0].includes('لا توجد') ? 0 : archiveList.length)">
                            </span>
                        </div>

                        <div class="min-h-[420px] bg-slate-950 rounded-2xl border border-slate-800 flex items-center justify-center relative overflow-auto p-2">
                            <template x-if="archiveList.length > 0 && !archiveList[archiveIndex].includes('لا توجد')">
                                <div class="text-center space-y-3 w-full h-full flex flex-col items-center justify-center">
                                    <template x-if="archiveList[archiveIndex].match(/\.(jpg|jpeg|png|gif|webp)$/i)">
                                        <div class="w-full flex justify-center items-center overflow-auto">
                                            <img :src="('/api/archive/file/' + archiveList[archiveIndex])"
                                                 class="max-h-[450px] w-auto max-w-full rounded-xl border border-slate-700 object-contain shadow-2xl transition-all duration-300">
                                        </div>
                                    </template>
                                    <template x-if="!archiveList[archiveIndex].match(/\.(jpg|jpeg|png|gif|webp)$/i)">
                                        <div class="py-12 space-y-3">
                                            <i class="fa-solid fa-file-pdf text-6xl text-rose-500"></i>
                                            <div class="text-sm text-slate-300 font-bold">ملف مستند مرفق</div>
                                        </div>
                                    </template>

                                    <div class="text-xs text-emerald-400 font-mono truncate max-w-lg dir-ltr text-center" x-text="archiveList[archiveIndex]"></div>

                                    <div class="flex gap-2 flex-wrap justify-center pt-1">
                                        <a :href="('/api/archive/file/' + archiveList[archiveIndex])" target="_blank" class="bg-amber-600 hover:bg-amber-500 text-white px-4 py-2 rounded-xl text-xs font-bold transition shadow flex items-center gap-1.5">
                                            <i class="fa-solid fa-up-right-and-down-left-from-center"></i> عرض بحجم كامل / طباعة
                                        </a>
                                        <button @click="deleteArchiveFile(archiveList[archiveIndex])" class="bg-rose-600 hover:bg-rose-500 text-white px-4 py-2 rounded-xl text-xs font-bold transition shadow cursor-pointer flex items-center gap-1.5">
                                            <i class="fa-solid fa-trash"></i> حذف الملف
                                        </button>
                                    </div>
                                </div>
                            </template>

                            <template x-if="archiveList.length === 0 || archiveList[archiveIndex].includes('لا توجد')">
                                <div class="text-slate-400 text-sm py-16 space-y-3">
                                    <i class="fa-solid fa-folder-open text-5xl text-slate-600 block"></i>
                                    <div>لا توجد مستندات مطابقة لهذا السجل حالياً</div>
                                </div>
                            </template>
                        </div>

                        <div class="flex flex-wrap justify-center gap-2 pt-2">
                            <button @click="prevArchive()" :disabled="archiveList.length <= 1 || archiveList[0].includes('لا توجد')" class="bg-slate-800 hover:bg-slate-700 px-4 py-2 rounded-xl text-xs font-bold disabled:opacity-40 cursor-pointer shadow">
                                <i class="fa-solid fa-chevron-right"></i> السابق
                            </button>
                            <button @click="nextArchive()" :disabled="archiveList.length <= 1 || archiveList[0].includes('لا توجد')" class="bg-slate-800 hover:bg-slate-700 px-4 py-2 rounded-xl text-xs font-bold disabled:opacity-40 cursor-pointer shadow">
                                التالي <i class="fa-solid fa-chevron-left"></i>
                            </button>
                            <label class="bg-blue-600 hover:bg-blue-500 text-white px-4 py-2 rounded-xl text-xs font-bold transition cursor-pointer shadow flex items-center gap-1.5">
                                <i class="fa-solid fa-upload"></i> رفع ملف جديد
                                <input type="file" @change="uploadArchiveFile($event)" class="hidden">
                            </label>
                            <label class="bg-purple-600 hover:bg-purple-500 text-white px-4 py-2 rounded-xl text-xs font-bold transition cursor-pointer shadow flex items-center gap-1.5">
                                <i class="fa-solid fa-print"></i> مسح ضوئي / التقاط مستند (Scanner)
                                <input type="file" accept="image/*,.pdf" capture="environment" @change="scanArchiveDocument($event)" class="hidden">
                            </label>
                        </div>
                    </div>
                </div>
            </div>

            <div class="bg-slate-800 border-t border-slate-700 px-6 py-3 flex justify-between items-center shrink-0">
                <div class="text-xs text-slate-400">جميع العمليات موثقة أمنياً عبر رمز الجلسة السيادية.</div>
                <div class="flex gap-3">
                    <button @click="closeModal()" class="bg-slate-700 hover:bg-slate-600 px-4 py-2 rounded-xl text-sm font-semibold cursor-pointer">إلغاء</button>
                    <button @click="saveCase()" :disabled="isSaving" class="bg-blue-600 hover:bg-blue-500 disabled:opacity-50 px-6 py-2 rounded-xl text-sm font-semibold shadow cursor-pointer flex items-center gap-2">
                        <span x-show="isSaving"><i class="fa-solid fa-spinner fa-spin"></i></span>
                        <span x-text="isSaving ? 'جاري الحفظ...' : 'حفظ وتوثيق أمني'"></span>
                    </button>
                </div>
            </div>

        </div>
    </div>

    <!-- Modal النسخة الاحتياطية -->
    <div x-show="showBackupModal" x-cloak class="fixed inset-0 bg-black/85 backdrop-blur-sm z-50 flex items-center justify-center p-4">
        <div class="bg-slate-800 border border-slate-700 w-full max-w-2xl rounded-3xl shadow-2xl overflow-hidden" @click.away="showBackupModal = false">
            <div class="bg-slate-800 border-b border-slate-700 px-6 py-4 flex justify-between items-center">
                <h3 class="text-base font-bold text-white flex items-center gap-2"><i class="fa-solid fa-database text-emerald-400"></i> إدارة النسخ الاحتياطي لقاعدة البيانات</h3>
                <button @click="showBackupModal = false" class="text-slate-400 hover:text-white text-xl cursor-pointer"><i class="fa-solid fa-xmark"></i></button>
            </div>
            <div class="p-6 space-y-6">
                <div class="bg-slate-900 border border-slate-700 p-5 rounded-2xl flex items-center justify-between">
                    <div>
                        <h4 class="text-sm font-bold text-white">إنشاء نسخة احتياطية جديدة</h4>
                        <p class="text-xs text-slate-400 mt-0.5">حفظ نسخة من بيانات الجداول على الخادم وإتاحتها للتحميل</p>
                    </div>
                    <button @click="generateBackup()" class="bg-emerald-600 hover:bg-emerald-500 text-white px-5 py-2.5 rounded-xl text-xs font-bold shadow transition cursor-pointer">توليد نسخة الآن 💾</button>
                </div>
                <div class="space-y-3">
                    <h4 class="text-xs font-bold text-slate-300 uppercase tracking-wider">النسخ الاحتياطية المتوفرة</h4>
                    <div class="bg-slate-900 border border-slate-700 rounded-2xl max-h-60 overflow-y-auto divide-y divide-slate-800">
                        <template x-for="b in backupsList" :key="b.filename">
                            <div class="p-4 flex items-center justify-between hover:bg-slate-800/50">
                                <div>
                                    <div class="text-sm font-bold text-emerald-400 font-mono" x-text="b.filename"></div>
                                    <div class="text-xs text-slate-400 mt-0.5" x-text="'الحجم: ' + b.size + ' | التوقيت: ' + b.time"></div>
                                </div>
                                <a :href="'/api/admin/backup/download/' + b.filename + '?token=' + authToken" target="_blank" class="bg-blue-600/20 hover:bg-blue-600 text-blue-300 hover:text-white px-3 py-1.5 rounded-xl text-xs font-semibold transition border border-blue-500/30 flex items-center gap-1 cursor-pointer"><i class="fa-solid fa-download"></i> تحميل</a>
                            </div>
                        </template>
                        <template x-if="backupsList.length === 0">
                            <div class="p-6 text-center text-xs text-slate-500">لا توجد نسخ احتياطية مسجلة حالياً</div>
                        </template>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal إدارة المستخدمين -->
    <div x-show="showUsersModal" x-cloak class="fixed inset-0 bg-black/85 backdrop-blur-sm z-50 flex items-center justify-center p-4 overflow-y-auto">
        <div class="bg-slate-800 border border-slate-700 w-full max-w-5xl rounded-3xl shadow-2xl overflow-hidden my-4" @click.away="showUsersModal = false">
            <div class="bg-slate-800 border-b border-slate-700 px-6 py-4 flex justify-between items-center">
                <h3 class="text-base font-bold text-white flex items-center gap-2"><i class="fa-solid fa-users-gear text-indigo-400"></i> لوحة التحكم السيادية لإدارة المستخدمين والصلاحيات (RBAC)</h3>
                <button @click="showUsersModal = false" class="text-slate-400 hover:text-white text-xl cursor-pointer"><i class="fa-solid fa-xmark"></i></button>
            </div>
            <div class="p-6 space-y-6 max-h-[70vh] overflow-y-auto">
                <div class="bg-slate-900 border border-slate-700 p-5 rounded-2xl space-y-4">
                    <h4 class="text-xs font-bold text-indigo-400 uppercase tracking-wider" x-text="userEditMode ? 'تعديل بيانات المستخدم' : 'إضافة مستخدم جديد للنظام'"></h4>
                    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
                        <div><label class="block text-xs text-slate-400 mb-1">اسم المستخدم</label><input type="text" x-model="userForm.Username" :disabled="userEditMode" class="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-sm text-white disabled:opacity-50"></div>
                        <div><label class="block text-xs text-slate-400 mb-1">الاسم الكامل</label><input type="text" x-model="userForm.FullName" class="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-sm text-white"></div>
                        <div><label class="block text-xs text-slate-400 mb-1">كلمة المرور</label><input type="password" x-model="userForm.Password" placeholder="اتركه فارغاً لعدم التغيير" class="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-sm text-white"></div>
                        <div>
                            <label class="block text-xs text-slate-400 mb-1">الصلاحية (Role)</label>
                            <select x-model="userForm.Role" class="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-sm text-white">
                                <option value="Admin">مدير النظام (Admin)</option>
                                <option value="Supervisor">مشرف (Supervisor)</option>
                                <option value="DataEntry">مدخل بيانات (DataEntry)</option>
                                <option value="Viewer">مستعرض فقط (Viewer)</option>
                            </select>
                        </div>
                    </div>
                    <div class="flex justify-between items-center pt-2">
                        <label class="flex items-center gap-2 cursor-pointer text-xs text-slate-300">
                            <input type="checkbox" x-model="userForm.IsActive" class="w-4 h-4 rounded accent-blue-600"> الحساب نشط وفعال
                        </label>
                        <div class="flex gap-2">
                            <button @click="resetUserForm()" x-show="userEditMode" class="bg-slate-700 hover:bg-slate-600 px-4 py-2 rounded-xl text-xs font-bold text-white">إلغاء</button>
                            <button @click="saveUser()" class="bg-indigo-600 hover:bg-indigo-500 px-5 py-2 rounded-xl text-xs font-bold text-white shadow" x-text="userEditMode ? 'حفظ التعديلات' : 'إضافة المستخدم'"></button>
                        </div>
                    </div>
                </div>

                <div class="bg-slate-900 border border-slate-700 rounded-2xl overflow-hidden">
                    <table class="w-full text-right text-xs">
                        <thead>
                            <tr class="bg-slate-800 border-b border-slate-700 text-slate-400">
                                <th class="p-3">#</th>
                                <th class="p-3">اسم المستخدم</th>
                                <th class="p-3">الاسم الكامل</th>
                                <th class="p-3">الصلاحية</th>
                                <th class="p-3">الحالة</th>
                                <th class="p-3">تاريخ الإنشاء</th>
                                <th class="p-3 text-center">الإجراءات</th>
                            </tr>
                        </thead>
                        <tbody class="divide-y divide-slate-800">
                            <template x-for="u in usersList" :key="u.UserID">
                                <tr class="hover:bg-slate-800/50">
                                    <td class="p-3 text-slate-500" x-text="u.UserID"></td>
                                    <td class="p-3 font-bold text-indigo-300" x-text="u.Username"></td>
                                    <td class="p-3 text-slate-200" x-text="u.FullName || '-'"></td>
                                    <td class="p-3"><span class="px-2 py-0.5 rounded text-[10px] font-bold bg-indigo-500/20 text-indigo-400" x-text="u.Role"></span></td>
                                    <td class="p-3"><span class="px-2 py-0.5 rounded text-[10px] font-bold" :class="u.IsActive ? 'bg-emerald-500/20 text-emerald-400' : 'bg-rose-500/20 text-rose-400'" x-text="u.IsActive ? 'نشط' : 'معطل'"></span></td>
                                    <td class="p-3 text-slate-400 font-mono text-[11px]" x-text="u.CreatedAt || '-'"></td>
                                    <td class="p-3 text-center">
                                        <div class="flex justify-center gap-1.5">
                                            <button @click="editUser(u)" class="bg-amber-600/20 hover:bg-amber-600 text-amber-400 hover:text-white px-2 py-1 rounded text-xs transition" title="تعديل"><i class="fa-solid fa-pen-to-square"></i></button>
                                            <button @click="deleteUser(u.UserID)" class="bg-rose-600/20 hover:bg-rose-600 text-rose-400 hover:text-white px-2 py-1 rounded text-xs transition" title="حذف"><i class="fa-solid fa-trash"></i></button>
                                        </div>
                                    </td>
                                </tr>
                            </template>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal سجل التدقيق الأمني -->
    <div x-show="showAuditModal" x-cloak class="fixed inset-0 bg-black/85 backdrop-blur-sm z-50 flex items-center justify-center p-4 overflow-y-auto">
        <div class="bg-slate-800 border border-slate-700 w-full max-w-6xl rounded-3xl shadow-2xl overflow-hidden my-4" @click.away="showAuditModal = false">
            <div class="bg-slate-800 border-b border-slate-700 px-6 py-4 flex justify-between items-center">
                <h3 class="text-base font-bold text-white flex items-center gap-2"><i class="fa-solid fa-user-shield text-purple-400"></i> لوحة تتبع نشاط المستخدمين والتدقيق الأمني (Audit Trail)</h3>
                <button @click="showAuditModal = false" class="text-slate-400 hover:text-white text-xl cursor-pointer"><i class="fa-solid fa-xmark"></i></button>
            </div>
            <div class="p-6 space-y-6 max-h-[75vh] overflow-y-auto">
                <div>
                    <h4 class="text-xs font-bold text-purple-400 uppercase tracking-wider mb-3">إحصائيات نشاط الموظفين وسجلات الإدخال</h4>
                    <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                        <template x-for="(statsObj, uname) in auditSummary" :key="uname">
                            <div class="bg-slate-900 border border-slate-700 p-4 rounded-2xl space-y-3 shadow">
                                <div class="flex justify-between items-center border-b border-slate-800 pb-2">
                                    <span class="font-bold text-emerald-400 text-sm flex items-center gap-1.5"><i class="fa-solid fa-user-tie"></i> <span x-text="uname"></span></span>
                                    <span class="bg-purple-600/20 text-purple-300 px-2.5 py-0.5 rounded-full text-xs font-bold" x-text="'إجمالي السجلات: ' + (statsObj.TOTAL_ADDED_RECORDS || 0)"></span>
                                </div>
                                <div class="grid grid-cols-2 gap-2 text-xs">
                                    <div class="bg-slate-800 p-2 rounded-xl"><span class="text-slate-400 block">إدخال (INSERT):</span><span class="text-blue-400 font-bold text-sm" x-text="statsObj.INSERT || 0"></span></div>
                                    <div class="bg-slate-800 p-2 rounded-xl"><span class="text-slate-400 block">تعديل (UPDATE):</span><span class="text-amber-400 font-bold text-sm" x-text="statsObj.UPDATE || 0"></span></div>
                                    <div class="bg-slate-800 p-2 rounded-xl"><span class="text-slate-400 block">حذف (DELETE):</span><span class="text-rose-400 font-bold text-sm" x-text="statsObj.DELETE || 0"></span></div>
                                    <div class="bg-slate-800 p-2 rounded-xl"><span class="text-slate-400 block">دخول (LOGIN):</span><span class="text-emerald-400 font-bold text-sm" x-text="statsObj.LOGIN || 0"></span></div>
                                </div>
                            </div>
                        </template>
                    </div>
                </div>

                <div class="bg-slate-900 border border-slate-700 rounded-2xl overflow-hidden">
                    <div class="bg-slate-800 px-4 py-3 border-b border-slate-700 text-xs font-bold text-slate-300">سجل العمليات الأخير المفصل</div>
                    <table class="w-full text-right text-xs">
                        <thead>
                            <tr class="bg-slate-800/60 border-b border-slate-700 text-slate-400">
                                <th class="p-3">#</th>
                                <th class="p-3">المستخدم</th>
                                <th class="p-3">نوع العملية</th>
                                <th class="p-3">الجدول</th>
                                <th class="p-3">التفاصيل</th>
                                <th class="p-3">التوقيت</th>
                            </tr>
                        </thead>
                        <tbody class="divide-y divide-slate-800">
                            <template x-for="log in auditLogs" :key="log.LogID">
                                <tr class="hover:bg-slate-800/50">
                                    <td class="p-3 text-slate-500" x-text="log.LogID"></td>
                                    <td class="p-3 font-bold text-emerald-400" x-text="log.Username"></td>
                                    <td class="p-3"><span class="px-2 py-0.5 rounded text-[10px] font-bold bg-blue-500/20 text-blue-400" x-text="log.ActionType"></span></td>
                                    <td class="p-3 text-slate-300 font-mono" x-text="log.TargetTable"></td>
                                    <td class="p-3 text-slate-300" x-text="log.Details"></td>
                                    <td class="p-3 text-slate-400 font-mono text-[11px]" x-text="log.Timestamp"></td>
                                </tr>
                            </template>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <script>
        function appData() {
            return {
                isLoggedIn: false, isLoggingIn: false, loginError: '', currentUser: '', currentRole: '', authToken: '',
                loginForm: { username: '', password: '' },
                tableData: [], stats: {}, currentPage: 1, totalPages: 1, totalRecords: 0,
                searchQuery: '', searchVarCode: '', searchCaseNo: '', searchDecisionNo: '', searchNationalID: '', searchLat: '', searchLon: '', searchOffender: '', searchAddress: '',
                statusFilter: '', districtFilter: '', vtypeFilter: '', yearFilter: '', loading: false, isExporting: false,
                showModal: false, showAuditModal: false, showUsersModal: false, showBackupModal: false, isEditMode: false, currentCaseId: null, activeTab: 'cases',
                isSaving: false,
                map: null, markersGroup: null, chartInstance: null, currentLayer: 'satellite', satelliteLayer: null, streetsLayer: null, mapOfflineNotice: false,
                archiveList: [], archiveIndex: 0,
                auditSummary: {}, auditLogs: [],
                usersList: [], userEditMode: false, currentUserId: null,
                backupsList: [],
                defaultDistricts: [
                    "المرج الغربي", "المرج الشرقية", "المرج البحرية", "المرج القبلية", "كفر ابو صير", "كفر الباشا", "كفر الشرفا", "منطقة زينب الوكيل", "ابو حوشان", "المنية", "العزبه البيضاء", "بركة الحاج", "الاندلس", "العصاره الجديدة", "بركة النصر"
                ],
                defaultViolationTypes: [
                    "بناء بدون ترخيص", "مخالفة ترخيص البناء المنصرف", "تعدي على أملاك الدولة", "تجاوز قيود الارتفاع / زيادة أدوار", "بناء على أرض زراعية", "تغيير استخدام (من سكني إلى تجاري/إداري)", "مخالفة خط التنظيم / البروزات", "عدم الالتزام بالرسومات الهندسية المعتمدة", "استئناف أعمال بعد صدور قرار وقف", "ترميم / تعديل بإنشاءات بدون ترخيص", "هدم بدون ترخيص"
                ],
                userForm: { Username: '', Password: '', FullName: '', Role: 'DataEntry', IsActive: true },
                form: {
                    SerialNo: null, VariableCode: '', VariableDate: '', CaseNumber: '', CaseDate: '', year: new Date().getFullYear(), StopNumber: '', AdminNumber: '', DecisionNumber: '', DecisionDate: '', AreaSize: null, Notes: '', Latitude: '', Longitude: '',
                    OffenderName: '', OffenderNationalID: null, OffenderAddress: '', Phone: null, OffenderNotes: '',
                    StreetName: '', DistrictName: '', FullAddress: '', Landmark: '',
                    ExecutionStatusName: 'تنفيذ جزئي', procedure: '', ExecutionLevel: '', ExecNotes: '',
                    HasAppeal: false, CourtDecision: '', CourtDecisionDate: null, DeputyGovernorDecision: null, DateDeputyGovernorDecision: null, FinalJudgement: null, FinalJudgementDate: null, TheFine: null, NonExecutionNumber: null, DepartmentIssueDate: null, FelonyNumber: null, JudgementNotes: '',
                    ReconciliationNumber: '', ReconciliationDate: '', FineAmount: null, ReconNotes: '',
                    ViolationTypeName: '', ViolationDescription: ''
                },
                buildFilterQueryString() {
                    let params = new URLSearchParams();
                    if (this.searchQuery && this.searchQuery.trim()) params.append('q', this.searchQuery.trim());
                    if (this.searchVarCode && this.searchVarCode.trim()) params.append('var_code', this.searchVarCode.trim());
                    if (this.searchCaseNo && this.searchCaseNo.trim()) params.append('case_no', this.searchCaseNo.trim());
                    if (this.searchDecisionNo && this.searchDecisionNo.trim()) params.append('decision_no', this.searchDecisionNo.trim());
                    if (this.searchNationalID && this.searchNationalID.trim()) params.append('national_id', this.searchNationalID.trim());
                    if (this.searchLat && this.searchLat.trim()) params.append('lat', this.searchLat.trim());
                    if (this.searchLon && this.searchLon.trim()) params.append('lon', this.searchLon.trim());
                    if (this.searchOffender && this.searchOffender.trim()) params.append('offender', this.searchOffender.trim());
                    if (this.searchAddress && this.searchAddress.trim()) params.append('address', this.searchAddress.trim());
                    if (this.statusFilter && this.statusFilter.trim()) params.append('status', this.statusFilter.trim());
                    if (this.districtFilter && this.districtFilter.trim()) params.append('district', this.districtFilter.trim());
                    if (this.vtypeFilter && this.vtypeFilter.trim()) params.append('vtype', this.vtypeFilter.trim());
                    if (this.yearFilter) params.append('year', this.yearFilter);
                    return params.toString();
                },
                async login() {
                    if (this.isLoggingIn) return;
                    this.isLoggingIn = true;
                    this.loginError = '';
                    try {
                        let res = await fetch('/api/auth/login', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({
                                username: (this.loginForm.username || '').trim(),
                                password: this.loginForm.password || ''
                            })
                        });
                        let data = await res.json();
                        if (res.ok && data.access_token) {
                            this.isLoggedIn = true;
                            this.authToken = data.access_token;
                            this.currentUser = data.username;
                            this.currentRole = data.role;
                            this.loginError = '';
                            this.initMap();
                            this.loadStats();
                            this.loadData(1);
                        } else {
                            this.loginError = data.detail || data.message || "فشل المصادقة الأمنية، تحقق من اسم المستخدم وكلمة المرور";
                        }
                    } catch(e) {
                        this.loginError = "تعذر الاتصال بالخادم، تحقق من تشغيل الخدمة أو الاتصال بالشبكة";
                    } finally {
                        this.isLoggingIn = false;
                    }
                },
                logout() {
                    this.isLoggedIn = false;
                    this.currentUser = '';
                    this.currentRole = '';
                    this.authToken = '';
                    this.loginError = '';
                    this.loginForm = { username: '', password: '' };
                },
                async scanArchiveDocument(event) {
                    if (!this.currentCaseId) {
                        alert("يجب حفظ السجل أو فتحه أولاً قبل إجراء المسح الضوئي");
                        if (event && event.target) event.target.value = '';
                        return;
                    }
                    let file = event && event.target && event.target.files ? event.target.files[0] : null;
                    if (!file) return;
                    let formData = new FormData();
                    formData.append("file", file);
                    try {
                        let res = await fetch(`/api/cases/${this.currentCaseId}/archive/scan`, {
                            method: 'POST',
                            headers: { 'Authorization': 'Bearer ' + this.authToken },
                            body: formData
                        });
                        let data = await res.json();
                        if (res.ok && data.status === 'success') {
                            alert("✔ " + data.message);
                            this.loadCaseArchive(this.currentCaseId);
                        } else {
                            alert(data.detail || data.message || "فشل إجراء المسح الضوئي");
                        }
                    } catch(e) {
                        alert("حدث خطأ أثناء رفع وحفظ المسح الضوئي");
                    } finally {
                        if (event && event.target) event.target.value = '';
                    }
                },
                async openBackupModal() {
                    try {
                        let res = await fetch('/api/admin/backups', { headers: { 'Authorization': 'Bearer ' + this.authToken } });
                        let data = await res.json();
                        if(res.ok && data.status === 'success') {
                            this.backupsList = data.backups;
                            this.showBackupModal = true;
                        } else { alert(data.detail || "غير مسموح"); }
                    } catch(e) { alert("تعذر جلب قائمة النسخ الاحتياطية"); }
                },
                async generateBackup() {
                    try {
                        let res = await fetch('/api/admin/backup', {
                            method: 'POST',
                            headers: { 'Authorization': 'Bearer ' + this.authToken }
                        });
                        let data = await res.json();
                        if(res.ok && data.status === 'success') {
                            alert("✔ " + data.message);
                            this.openBackupModal();
                        } else { alert(data.detail || "فشل إنشاء النسخة الاحتياطية"); }
                    } catch(e) { alert("خطأ في الاتصال"); }
                },
                async openUsersModal() {
                    try {
                        let res = await fetch('/api/admin/users', { headers: { 'Authorization': 'Bearer ' + this.authToken } });
                        let data = await res.json();
                        if(res.ok && data.status === 'success') {
                            this.usersList = data.data;
                            this.resetUserForm();
                            this.showUsersModal = true;
                        } else { alert(data.detail || "غير مسموح"); }
                    } catch(e) { alert("تعذر جلب قائمة المستخدمين"); }
                },
                resetUserForm() {
                    this.userEditMode = false;
                    this.currentUserId = null;
                    this.userForm = { Username: '', Password: '', FullName: '', Role: 'DataEntry', IsActive: true };
                },
                editUser(u) {
                    this.userEditMode = true;
                    this.currentUserId = u.UserID;
                    this.userForm = { Username: u.Username, Password: '', FullName: u.FullName, Role: u.Role, IsActive: u.IsActive };
                },
                async saveUser() {
                    try {
                        let url = this.userEditMode ? `/api/admin/users/${this.currentUserId}` : `/api/admin/users`;
                        let method = this.userEditMode ? 'PUT' : 'POST';
                        let res = await fetch(url, {
                            method: method,
                            headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + this.authToken },
                            body: JSON.stringify(this.userForm)
                        });
                        let data = await res.json();
                        if(res.ok && data.status === 'success') {
                            alert("✔ " + data.message);
                            this.openUsersModal();
                        } else { alert(data.detail || data.message || "خطأ في الصلاحيات"); }
                    } catch(e) { alert("خطأ في حفظ بيانات المستخدم"); }
                },
                async deleteUser(uid) {
                    if(confirm("هل أنت متأكد من حذف هذا المستخدم نهائياً؟")) {
                        try {
                            let res = await fetch(`/api/admin/users/${uid}`, {
                                method: 'DELETE',
                                headers: { 'Authorization': 'Bearer ' + this.authToken }
                            });
                            let data = await res.json();
                            if(res.ok && data.status === 'success') {
                                alert("✔ " + data.message);
                                this.openUsersModal();
                            } else { alert(data.detail || "تعذر الحذف"); }
                        } catch(e) { alert("خطأ أثناء الحذف"); }
                    }
                },
                async openAuditModal() {
                    try {
                        let res = await fetch('/api/users/stats', { headers: { 'Authorization': 'Bearer ' + this.authToken } });
                        let data = await res.json();
                        if(res.ok && data.status === 'success') {
                            this.auditSummary = data.summary;
                            this.auditLogs = data.logs;
                            this.showAuditModal = true;
                        } else { alert(data.detail || "غير مسموح"); }
                    } catch(e) { alert("تعذر جلب سجل التدقيق"); }
                },
                initMap() {
                    if(this.map || typeof L === 'undefined') return;
                    try {
                        this.map = L.map('map').setView([26.8206, 30.8025], 6);
                        this.satelliteLayer = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', { maxZoom: 19, attribution: 'Esri' });
                        this.streetsLayer = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 19, attribution: 'OSM' });
                        this.mapOfflineNotice = !navigator.onLine;
                        const handleTileError = () => { this.mapOfflineNotice = true; };
                        this.satelliteLayer.on('tileerror', handleTileError);
                        this.streetsLayer.on('tileerror', handleTileError);
                        window.addEventListener('offline', () => { this.mapOfflineNotice = true; });
                        window.addEventListener('online', () => { this.mapOfflineNotice = false; });
                        this.satelliteLayer.addTo(this.map);
                        this.markersGroup = L.layerGroup().addTo(this.map);
                    } catch(e) {
                        this.mapOfflineNotice = true;
                    }
                },
                switchMapLayer(layerType) {
                    if (!this.map) return;
                    this.currentLayer = layerType;
                    if (layerType === 'satellite') {
                        if (this.streetsLayer && this.map.hasLayer(this.streetsLayer)) this.map.removeLayer(this.streetsLayer);
                        if (this.satelliteLayer) this.satelliteLayer.addTo(this.map);
                    } else {
                        if (this.satelliteLayer && this.map.hasLayer(this.satelliteLayer)) this.map.removeLayer(this.satelliteLayer);
                        if (this.streetsLayer) this.streetsLayer.addTo(this.map);
                    }
                },
                captureCurrentLocation() {
                    if (navigator.geolocation) {
                        navigator.geolocation.getCurrentPosition(
                            (pos) => {
                                this.form.Latitude = pos.coords.latitude.toFixed(6);
                                this.form.Longitude = pos.coords.longitude.toFixed(6);
                                alert("✔ تم جلب الإحداثيات بنجاح!");
                            },
                            (err) => { alert("تعذر التقاط الموقع."); },
                            { enableHighAccuracy: true }
                        );
                    } else { alert("GPS غير مدعوم"); }
                },
                updateMapMarkers() {
                    if (!this.markersGroup || !this.map) return;
                    this.markersGroup.clearLayers();
                    let validPoints = [];
                    this.tableData.forEach(row => {
                        if (row.Latitude && row.Longitude) {
                            let lat = parseFloat(row.Latitude);
                            let lon = parseFloat(row.Longitude);
                            if (!isNaN(lat) && !isNaN(lon)) {
                                let marker = L.marker([lat, lon]).bindPopup(`
                                    <div style="direction: rtl; text-align: right; font-family: 'Cairo', sans-serif; font-size: 12px;">
                                        <b>المخالف:</b> ${row.OffenderName || 'غير مسجل'}<br>
                                        <b>المنطقة:</b> ${row.DistrictName || '-'}<br>
                                        <b>نوع المخالفة:</b> ${row.ViolationTypeName || '-'}<br>
                                        <b>الرقم القومي:</b> ${row.OffenderNationalID || '-'}<br>
                                        <b>محضر:</b> ${row.CaseNumber || '-'}<br>
                                        <b>كود المتغير:</b> ${row.VariableCode}<br>
                                        <b>الحالة:</b> ${row.ExecutionStatusName}<br>
                                        <span style="color: #10b981; font-family: monospace;">📌 ${lat}, ${lon}</span>
                                    </div>
                                `);
                                this.markersGroup.addLayer(marker);
                                validPoints.push([lat, lon]);
                            }
                        }
                    });
                    if (validPoints.length > 0) {
                        let bounds = L.latLngBounds(validPoints);
                        this.map.fitBounds(bounds, { padding: [50, 50], maxZoom: 17 });
                    }
                },
                focusOnMap(row) {
                    if (!this.map || !this.markersGroup) { alert("الخريطة غير متاحة حالياً"); return; }
                    if (row.Latitude && row.Longitude) {
                        let lat = parseFloat(row.Latitude);
                        let lon = parseFloat(row.Longitude);
                        if (!isNaN(lat) && !isNaN(lon)) {
                            this.map.setView([lat, lon], 19);
                            this.markersGroup.eachLayer(layer => {
                                let ll = layer.getLatLng();
                                if (Math.abs(ll.lat - lat) < 0.00001 && Math.abs(ll.lng - lon) < 0.00001) { layer.openPopup(); }
                            });
                        }
                    } else { alert("لا توجد إحداثيات مسجلة"); }
                },
                updateChart(breakdown) {
                    if (typeof Chart === 'undefined') return;
                    const ctx = document.getElementById('statusChart');
                    if (!ctx) return;
                    let labels = Object.keys(breakdown || {});
                    let values = Object.values(breakdown || {});
                    if (this.chartInstance) this.chartInstance.destroy();
                    let self = this;
                    this.chartInstance = new Chart(ctx, {
                        type: 'doughnut',
                        data: { labels: labels, datasets: [{ data: values, backgroundColor: ['#10b981', '#f43f5e', '#f59e0b', '#3b82f6', '#8b5cf6', '#64748b', '#06b6d4'], borderWidth: 1 }] },
                        options: {
                            responsive: true, maintainAspectRatio: false,
                            onClick: (event, elements) => { if (elements.length > 0) { self.statusFilter = labels[elements[0].index]; self.loadData(1); } },
                            plugins: { legend: { position: 'bottom', labels: { color: '#cbd5e1', font: { family: 'Cairo', size: 11 } } } }
                        }
                    });
                },
                async loadStats() {
                    try {
                        let res = await fetch('/api/stats', { headers: { 'Authorization': 'Bearer ' + this.authToken } });
                        let data = await res.json();
                        if (res.ok && data.status === 'success') { this.stats = data; this.updateChart(data.status_breakdown); }
                    } catch (e) { console.error(e); }
                },
                async loadData(page = 1) {
                    if (this.loading) return;
                    this.loading = true;
                    this.currentPage = page;
                    try {
                        let qs = this.buildFilterQueryString();
                        let joiner = qs ? '&' : '';
                        let url = `/api/cases/search?${qs}${joiner}page=${this.currentPage}&limit=50`;
                        let res = await fetch(url, { headers: { 'Authorization': 'Bearer ' + this.authToken } });
                        let data = await res.json();
                        if (res.ok && data.status === 'success') {
                            this.tableData = data.data;
                            this.totalRecords = data.total;
                            this.totalPages = data.total_pages;
                            this.updateMapMarkers();
                        }
                    } catch (e) { console.error(e); }
                    finally { this.loading = false; setTimeout(() => { if(this.map) this.map.invalidateSize(); }, 200); }
                },
                async loadCaseArchive(caseId) {
                    try {
                        let res = await fetch(`/api/cases/${caseId}/archive-files`, {
                            headers: { 'Authorization': 'Bearer ' + this.authToken }
                        });
                        let data = await res.json();
                        if(res.ok && data.status === 'success') {
                            this.archiveList = data.files.length > 0 ? data.files : ['لا توجد مستندات مطابقة لهذا السجل'];
                            this.archiveIndex = 0;
                        }
                    } catch(e) {
                        this.archiveList = ['خطأ في جلب الأرشيف'];
                    }
                },
                async uploadArchiveFile(event) {
                    let file = event.target.files[0];
                    if (!file || !this.currentCaseId) { alert("يجب حفظ السجل أو فتحه أولاً قبل رفع المستندات"); return; }
                    let formData = new FormData();
                    formData.append("file", file);
                    try {
                        let res = await fetch(`/api/cases/${this.currentCaseId}/archive/upload`, {
                            method: 'POST',
                            headers: { 'Authorization': 'Bearer ' + this.authToken },
                            body: formData
                        });
                        let data = await res.json();
                        if(res.ok && data.status === 'success') {
                            alert("✔ تم رفع وتخزين المستند بنجاح في الأرشيف السحابي!");
                            this.loadCaseArchive(this.currentCaseId);
                        } else { alert(data.detail || "فشل رفع الملف"); }
                    } catch(e) { alert("خطأ في الاتصال"); }
                    finally { if (event && event.target) event.target.value = ''; }
                },
                async deleteArchiveFile(filename) {
                    if(confirm(`هل أنت متأكد من حذف المستند (${filename}) نهائياً؟`)) {
                        try {
                            let res = await fetch(`/api/archive/file/${filename}`, {
                                method: 'DELETE',
                                headers: { 'Authorization': 'Bearer ' + this.authToken }
                            });
                            let data = await res.json();
                            if(res.ok && data.status === 'success') {
                                alert("✔ " + data.message);
                                if(this.currentCaseId) this.loadCaseArchive(this.currentCaseId);
                            } else { alert(data.detail || "تعذر الحذف"); }
                        } catch(e) { alert("خطأ في الاتصال"); }
                    }
                },
                async openModal() {
                    this.isEditMode = false; this.currentCaseId = null; this.activeTab = 'cases'; this.isSaving = false;
                    try {
                        let res = await fetch('/api/next-serial', { headers: { 'Authorization': 'Bearer ' + this.authToken } });
                        let data = await res.json();
                        this.form.SerialNo = data.next_serial;
                    } catch(e) { this.form.SerialNo = 1; }

                    this.form.VariableCode = ''; this.form.VariableDate = ''; this.form.CaseNumber = ''; this.form.CaseDate = ''; this.form.year = new Date().getFullYear(); this.form.StopNumber = ''; this.form.AdminNumber = ''; this.form.DecisionNumber = ''; this.form.DecisionDate = ''; this.form.AreaSize = null; this.form.Notes = ''; this.form.Latitude = ''; this.form.Longitude = '';
                    this.form.OffenderName = ''; this.form.OffenderNationalID = null; this.form.OffenderAddress = ''; this.form.Phone = null; this.form.OffenderNotes = '';
                    this.form.StreetName = ''; this.form.DistrictName = ''; this.form.FullAddress = ''; this.form.Landmark = '';
                    this.form.ExecutionStatusName = 'تنفيذ جزئي'; this.form.procedure = ''; this.form.ExecutionLevel = ''; this.form.ExecNotes = '';
                    this.form.HasAppeal = false; this.form.CourtDecision = ''; this.form.CourtDecisionDate = null; this.form.DeputyGovernorDecision = null; this.form.DateDeputyGovernorDecision = null; this.form.FinalJudgement = null; this.form.FinalJudgementDate = null; this.form.TheFine = null; this.form.NonExecutionNumber = null; this.form.DepartmentIssueDate = null; this.form.FelonyNumber = null; this.form.JudgementNotes = '';
                    this.form.ReconciliationNumber = ''; this.form.ReconciliationDate = ''; this.form.FineAmount = null; this.form.ReconNotes = '';
                    this.form.ViolationTypeName = ''; this.form.ViolationDescription = '';
                    this.archiveList = ['لا توجد مستندات مطابقة لهذا السجل'];
                    this.showModal = true;
                },
                editCase(row) {
                    this.isEditMode = true; this.currentCaseId = row.CaseID; this.activeTab = 'cases'; this.isSaving = false;
                    this.form = {
                        SerialNo: row.SerialNo === '-' ? null : row.SerialNo,
                        VariableCode: row.VariableCode === '-' ? '' : row.VariableCode,
                        VariableDate: row.VariableDate || '',
                        CaseNumber: row.CaseNumber === '-' ? '' : row.CaseNumber,
                        CaseDate: row.CaseDate || '',
                        year: row.year === '-' ? null : row.year,
                        StopNumber: row.StopNumber || '',
                        AdminNumber: row.AdminNumber || '',
                        DecisionNumber: row.DecisionNumber === '-' ? '' : row.DecisionNumber,
                        DecisionDate: row.DecisionDate || '',
                        AreaSize: row.AreaSize || null,
                        Notes: row.Notes || '',
                        Latitude: row.Latitude || '',
                        Longitude: row.Longitude || '',
                        OffenderName: row.OffenderName || '',
                        OffenderNationalID: row.OffenderNationalID || null,
                        OffenderAddress: row.OffenderAddress || '',
                        Phone: row.Phone || null,
                        OffenderNotes: row.OffenderNotes || '',
                        StreetName: row.StreetName || '',
                        DistrictName: row.DistrictName || '',
                        FullAddress: row.FullAddress || '',
                        Landmark: row.Landmark || '',
                        ExecutionStatusName: row.ExecutionStatusName || 'تنفيذ جزئي',
                        procedure: row.procedure || '',
                        ExecutionLevel: row.ExecutionLevel || '',
                        ExecNotes: row.ExecNotes || '',
                        HasAppeal: row.HasAppeal || false,
                        CourtDecision: row.CourtDecision || '',
                        CourtDecisionDate: row.CourtDecisionDate || null,
                        DeputyGovernorDecision: row.DeputyGovernorDecision || null,
                        DateDeputyGovernorDecision: row.DateDeputyGovernorDecision || null,
                        FinalJudgement: row.FinalJudgement || null,
                        FinalJudgementDate: row.FinalJudgementDate || null,
                        TheFine: row.TheFine || null,
                        NonExecutionNumber: row.NonExecutionNumber || null,
                        DepartmentIssueDate: row.DepartmentIssueDate || null,
                        FelonyNumber: row.FelonyNumber || null,
                        JudgementNotes: row.JudgementNotes || '',
                        ReconciliationNumber: row.ReconciliationNumber || '',
                        ReconciliationDate: row.ReconciliationDate || '',
                        FineAmount: row.FineAmount || null,
                        ReconNotes: row.ReconNotes || '',
                        ViolationTypeName: row.ViolationTypeName || '',
                        ViolationDescription: row.ViolationDescription || ''
                    };
                    this.loadCaseArchive(row.CaseID);
                    this.showModal = true;
                },
                closeModal() { this.showModal = false; this.isSaving = false; },
                prevArchive() { if (this.archiveIndex > 0) this.archiveIndex--; },
                nextArchive() { if (this.archiveIndex < this.archiveList.length - 1) this.archiveIndex++; },
                async saveCase() {
                    if (this.isSaving) return;
                    this.isSaving = true;
                    try {
                        let url = this.isEditMode ? `/api/cases/${this.currentCaseId}` : `/api/cases`;
                        let res = await fetch(url, {
                            method: this.isEditMode ? 'PUT' : 'POST',
                            headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + this.authToken },
                            body: JSON.stringify(this.form)
                        });
                        let result = await res.json();
                        if (res.ok && result.status === 'success') {
                            alert("✔ " + result.message);
                            this.closeModal();
                            this.loadData(this.currentPage);
                            this.loadStats();
                        } else {
                            alert(result.detail || result.message || "خطأ في حفظ المحضر");
                            this.isSaving = false;
                        }
                    } catch (e) {
                        alert("خطأ في الاتصال بالخادم");
                        this.isSaving = false;
                    }
                },
                async deleteCase(caseId) {
                    if (confirm("هل أنت متأكد من الحذف وتوثيق العملية في سجل التدقيق؟")) {
                        try {
                            let res = await fetch(`/api/cases/${caseId}`, {
                                method: 'DELETE',
                                headers: { 'Authorization': 'Bearer ' + this.authToken }
                            });
                            let result = await res.json();
                            if (res.ok && result.status === 'success') {
                                this.loadData(this.currentPage); this.loadStats(); alert(result.message);
                            } else { alert(result.detail || "تعذر الحذف"); }
                        } catch(e) { alert("خطأ في الاتصال"); }
                    }
                },
                async downloadFileSecurely(endpoint, filename) {
                    if (this.isExporting) return;
                    this.isExporting = true;
                    try {
                        let qs = this.buildFilterQueryString();
                        let url = qs ? `${endpoint}?${qs}` : endpoint;
                        let res = await fetch(url, {
                            headers: { 'Authorization': 'Bearer ' + this.authToken }
                        });
                        if (!res.ok) {
                            let errData = null;
                            try { errData = await res.json(); } catch(_) {}
                            alert((errData && errData.detail) || "فشل عملية التصدير، تأكد من الصلاحيات أو البيانات");
                            return;
                        }
                        let blob = await res.blob();
                        let downloadUrl = window.URL.createObjectURL(blob);
                        let link = document.createElement('a');
                        link.href = downloadUrl;
                        link.download = filename;
                        document.body.appendChild(link);
                        link.click();
                        document.body.removeChild(link);
                        setTimeout(() => window.URL.revokeObjectURL(downloadUrl), 1500);
                    } catch(e) {
                        alert("حدث خطأ أثناء تنزيل الملف");
                    } finally {
                        this.isExporting = false;
                    }
                },
                exportToExcel() {
                    let dStr = new Date().toISOString().slice(0, 10);
                    this.downloadFileSecurely('/api/cases/export-excel', `Sovereign_Violations_Report_${dStr}.xlsx`);
                },
                exportToPDF() {
                    let dStr = new Date().toISOString().slice(0, 10);
                    this.downloadFileSecurely('/api/cases/export-pdf', `Sovereign_Violations_Report_${dStr}.pdf`);
                },
                async printFilteredReport() {
                    if (this.isExporting) return;
                    this.isExporting = true;
                    try {
                        let qs = this.buildFilterQueryString();
                        let joiner = qs ? '&' : '';
                        let url = `/api/cases/search?${qs}${joiner}page=1&limit=500`;
                        let res = await fetch(url, {
                            headers: { 'Authorization': 'Bearer ' + this.authToken }
                        });
                        let data = await res.json();
                        if (!res.ok || data.status !== 'success') {
                            alert(data.detail || "تعذر تجهيز بيانات الطباعة");
                            return;
                        }
                        let rows = data.data || [];
                        if (rows.length === 0) {
                            alert("لا توجد سجلات مطابقة للفلاتر الحالية لطباعتها");
                            return;
                        }
                        let printWin = window.open('', '_blank', 'width=1200,height=800');
                        if (!printWin) {
                            alert("يرجى السماح بفتح النوافذ المنبثقة (Pop-ups) في المتصفح لإتمام الطباعة");
                            return;
                        }
                        let filterSummary = [
                            this.searchVarCode ? `كود المتغير: ${this.searchVarCode}` : '',
                            this.searchCaseNo ? `رقم المحضر: ${this.searchCaseNo}` : '',
                            this.searchDecisionNo ? `قرار المحافظ: ${this.searchDecisionNo}` : '',
                            this.searchNationalID ? `الرقم القومي: ${this.searchNationalID}` : '',
                            this.searchOffender ? `المخالف: ${this.searchOffender}` : '',
                            this.searchAddress ? `العنوان: ${this.searchAddress}` : '',
                            this.districtFilter ? `المنطقة: ${this.districtFilter}` : '',
                            this.statusFilter ? `الحالة: ${this.statusFilter}` : '',
                            this.vtypeFilter ? `نوع المخالفة: ${this.vtypeFilter}` : '',
                            this.yearFilter ? `السنة: ${this.yearFilter}` : '',
                            this.searchQuery ? `بحث عام: ${this.searchQuery}` : ''
                        ].filter(Boolean).join(' | ') || 'جميع السجلات';

                        let trHtml = rows.map((r, i) => `
                            <tr>
                                <td>${i + 1}</td>
                                <td>${r.VariableCode || '-'}</td>
                                <td>${r.CaseNumber || '-'}</td>
                                <td>${r.DecisionNumber || '-'}</td>
                                <td>${r.OffenderName || 'غير مسجل'}<br><small>${r.OffenderNationalID || ''}</small></td>
                                <td>${r.FullAddress || r.OffenderAddress || '-'}<br><small>${(r.Latitude && r.Longitude) ? (r.Latitude + ', ' + r.Longitude) : ''}</small></td>
                                <td>${r.ViolationTypeName || '-'}</td>
                                <td>${r.year || '-'}</td>
                                <td>${r.ExecutionStatusName || '-'}</td>
                            </tr>
                        `).join('');

                        printWin.document.write(`
                            <!DOCTYPE html>
                            <html lang="ar" dir="rtl">
                            <head>
                                <meta charset="UTF-8">
                                <title>طباعة تقرير المحاضر والمخالفات</title>
                                <style>
                                    body { font-family: 'Cairo', Arial, sans-serif; direction: rtl; margin: 20px; color: #111; }
                                    h2 { margin: 0 0 8px 0; font-size: 18px; text-align: center; }
                                    .meta { font-size: 12px; margin-bottom: 14px; padding: 8px 12px; background: #f1f5f9; border: 1px solid #cbd5e1; border-radius: 6px; }
                                    table { width: 100%; border-collapse: collapse; font-size: 11px; }
                                    th, td { border: 1px solid #64748b; padding: 6px 8px; text-align: center; vertical-align: middle; }
                                    th { background: #1e293b; color: #fff; }
                                    tr:nth-child(even) { background: #f8fafc; }
                                    @media print { @page { size: A4 landscape; margin: 12mm; } }
                                </style>
                            </head>
                            <body>
                                <h2>تقرير المحاضر والمخالفات - المنظومة السيادية المؤمنة</h2>
                                <div class="meta">
                                    <strong>الفلاتر المطبقة:</strong> ${filterSummary}  | 
                                    <strong>إجمالي النتائج المفلترة:</strong> ${data.total}  | 
                                    <strong>المعروض في الطباعة:</strong> ${rows.length} سجل
                                </div>
                                <table>
                                    <thead>
                                        <tr>
                                            <th>#</th>
                                            <th>كود المتغير</th>
                                            <th>رقم المحضر</th>
                                            <th>رقم القرار</th>
                                            <th>اسم المخالف والرقم القومي</th>
                                            <th>العنوان (مع الإحداثيات)</th>
                                            <th>نوع المخالفة</th>
                                            <th>السنة</th>
                                            <th>حالة التنفيذ</th>
                                        </tr>
                                    </thead>
                                    <tbody>${trHtml}</tbody>
                                </table>
                            </body>
                            </html>
                        `);
                        printWin.document.close();
                        printWin.focus();
                        setTimeout(() => { printWin.print(); }, 300);
                    } catch (e) {
                        alert("حدث خطأ أثناء تجهيز الطباعة");
                    } finally {
                        this.isExporting = false;
                    }
                },
                exportToKML() {
                    let dStr = new Date().toISOString().slice(0, 10);
                    this.downloadFileSecurely('/api/cases/export-kml', `Sovereign_Violations_${dStr}.kml`);
                }
            }
        }
    </script>
</body>
</html>
"""

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", "8000"))
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=False)
