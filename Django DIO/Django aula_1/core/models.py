from django.db import models
from django.db.models.fields import TextField
from django.contrib.auth.models import User

# Create your models here.
class Evento(models.Model):
    titulo = models.CharField(max_length=100)
    descricao = models.TextField(blank=True, null=True)
    data_evento = models.DateTimeField(verbose_name='Data do Evento')
    date_criacao = models.DateTimeField(auto_now=True)
    usuario = models.ForeignKey(User, on_delete=models.CASCADE)

    class meta:
        db_table = 'evento'
    def __str__(self):
        return self.titulo