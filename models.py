from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


class Producto(db.Model):
    __tablename__ = "productos"

    id = db.Column(db.Integer, primary_key=True)
    PRODUCTO = db.Column(db.String(250), nullable=False)
    CODIGO = db.Column(db.String(100), nullable=False, unique=True)
    UBICACION = db.Column(db.String(150), nullable=False)

    def __repr__(self):
        return f"<Producto {self.CODIGO}>" 
