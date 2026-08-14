from django.db import migrations


def seed_pharmacy(apps, schema_editor):
    Pharmacy = apps.get_model('reconciliation', 'Pharmacy')
    Pharmacy.objects.get_or_create(name='صيدلية راما رسلان')


def unseed_pharmacy(apps, schema_editor):
    Pharmacy = apps.get_model('reconciliation', 'Pharmacy')
    Pharmacy.objects.filter(name='صيدلية راما رسلان').delete()


class Migration(migrations.Migration):

    dependencies = [
        ('reconciliation', '0001_initial'),
    ]

    operations = [
        migrations.RunPython(seed_pharmacy, unseed_pharmacy),
    ]
