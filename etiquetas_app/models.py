from django.db import models
from django.db import models
from django.contrib.auth.models import User

class ArchivoGenerado(models.Model):
    usuario = models.ForeignKey(User, on_delete=models.CASCADE, null=True, blank=True)
    excel_original = models.CharField(max_length=255)
    zip_original = models.CharField(max_length=255)
    documento_generado = models.FileField(upload_to='output')
    fecha_creacion = models.DateTimeField(auto_now_add=True)
    
    def __str__(self):
        return f"Documento generado el {self.fecha_creacion}"