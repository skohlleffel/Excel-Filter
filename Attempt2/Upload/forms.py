from django import forms
from .models import Excel

class FileForm(forms.ModelForm):
    class Meta:
        model = Excel
        fields = ('xlsx', )