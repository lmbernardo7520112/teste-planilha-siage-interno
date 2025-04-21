from sqlalchemy import Column, Integer, String, ForeignKey, Float, Boolean
from sqlalchemy.orm import relationship
from app.database import Base

class Turma(Base):
    __tablename__ = "turmas"
    
    id = Column(Integer, primary_key=True, index=True)
    nome_turma = Column(String, unique=True, index=True)
    
    alunos = relationship("Aluno", back_populates="turma")

class Aluno(Base):
    __tablename__ = "alunos"
    
    id = Column(Integer, primary_key=True, index=True)
    numero = Column(String, index=True)
    nome = Column(String, index=True)
    turma_id = Column(Integer, ForeignKey("turmas.id"))
    
    turma = relationship("Turma", back_populates="alunos")
    notas = relationship("Nota", back_populates="aluno")

class Disciplina(Base):
    __tablename__ = "disciplinas"
    
    id = Column(Integer, primary_key=True, index=True)
    codigo = Column(String, unique=True, index=True)  # Ex.: "BIO", "MAT"
    nome = Column(String, unique=True)  # Ex.: "Biologia", "Matemática"
    
    notas = relationship("Nota", back_populates="disciplina")

class Nota(Base):
    __tablename__ = "notas"
    
    id = Column(Integer, primary_key=True, index=True)
    aluno_id = Column(Integer, ForeignKey("alunos.id"))
    disciplina_id = Column(Integer, ForeignKey("disciplinas.id"))
    q1 = Column(Float, default=0.0)  # 1º BIM
    q2 = Column(Float, default=0.0)  # 2º BIM
    q3 = Column(Float, default=0.0)  # 3º BIM
    q4 = Column(Float, default=0.0)  # 4º BIM
    pf = Column(Float, default=0.0)  # Prova Final
    sf = Column(Float, default=0.0)  # Situação Final
    
    aluno = relationship("Aluno", back_populates="notas")
    disciplina = relationship("Disciplina", back_populates="notas")