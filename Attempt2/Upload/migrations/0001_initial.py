# Generated by Django 3.0.1 on 2020-02-02 21:19

from django.db import migrations, models


class Migration(migrations.Migration):

    initial = True

    dependencies = [
    ]

    operations = [
        migrations.CreateModel(
            name='Excel',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('xlsx', models.FileField(upload_to='files/xlsx/')),
            ],
        ),
    ]
