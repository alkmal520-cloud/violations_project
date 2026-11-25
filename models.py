import os
from sqlalchemy import create_engine, Column, Integer, String, ForeignKey, Text
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship

# إعدادات الاتصال بخادم PostgreSQL السحابي
POSTGRES_USER = os.getenv("DB_USER", "postgres")
POSTGRES_PASSWORD = os.getenv("DB_PASSWORD", "roo71900")
POSTGRES_HOST = os.getenv("DB_HOST", "localhost")
POSTGRES_PORT = os.getenv("DB_PORT", "5432")
POSTGRES_DB = os.getenv("DB_NAME", "violations_db")

DATABASE_URL = f"postgresql+psycopg2://{POSTGRES_USER}:{POSTGRES_PASSWORD}@{POSTGRES_HOST}:{POSTGRES_PORT}/{POSTGRES_DB}"

engine = create_engine(DATABASE_URL, pool_pre_ping=True)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()


class Case(Base):
    __tablename__ = "tbl_cases"

    CaseID = Column(Integer, primary_key=True, index=True)
    VariableCode = Column(String, index=True, nullable=True)
    CaseNumber = Column(String, index=True, nullable=True)
    CaseDate = Column(String, nullable=True)
    NationalID = Column(String, nullable=True)
    DecisionNumber = Column(String, nullable=True)
    year = Column(String, nullable=True)

    OffenderID = Column(Integer, ForeignKey("tbl_offenders.OffenderID"), nullable=True)
    LocationID = Column(Integer, ForeignKey("tbl_locations.LocationID"), nullable=True)

    offender = relationship("Offender", back_populates="cases")
    location = relationship("Location", back_populates="cases")
    execution_status = relationship("ExecutionStatus", back_populates="case")
    archive_files = relationship("Archive", back_populates="case")


class Offender(Base):
    __tablename__ = "tbl_offenders"

    OffenderID = Column(Integer, primary_key=True, index=True)
    OffenderName = Column(String, index=True, nullable=True)
    NationalID = Column(String, nullable=True)

    cases = relationship("Case", back_populates="offender")


class Location(Base):
    __tablename__ = "tbl_locations"

    LocationID = Column(Integer, primary_key=True, index=True)
    StreetName = Column(String, nullable=True)
    DistrictName = Column(String, nullable=True)

    cases = relationship("Case", back_populates="location")


class ExecutionStatus(Base):
    __tablename__ = "tbl_executionstatus"

    StatusID = Column(Integer, primary_key=True, index=True)
    CaseID = Column(Integer, ForeignKey("tbl_cases.CaseID"))
    procedure = Column(String, nullable=True)

    case = relationship("Case", back_populates="execution_status")


class Archive(Base):
    __tablename__ = "tbl_archive"

    ArchiveID = Column(Integer, primary_key=True, index=True)
    CaseID = Column(Integer, ForeignKey("tbl_cases.CaseID"))
    FileType = Column(String, nullable=True)
    FileName = Column(String, nullable=True)
    FilePath = Column(String, nullable=True)
    UploadDate = Column(String, nullable=True)

    case = relationship("Case", back_populates="archive_files")