import datetime
from dataclasses import dataclass

@dataclass
class Ticket:
    Ticket_Id: str
    Tipo: str
    Resumen: str
    Prioridad: str
    Estado: str
    Fecha_de_creacion: datetime.datetime
    Fecha_de_cambio: datetime.datetime
    Fecha_de_resolucion: datetime.datetime
    Resolucion: str
    Solicitante: str
    Asignado: str
    Descripcion: str
    Fecha_estimada: datetime.datetime
    Worksheet: str
