from django.db import models
# Create your models here.
class Excel(models.Model):
    xlsx = models.FileField(upload_to='files/xlsx/')
