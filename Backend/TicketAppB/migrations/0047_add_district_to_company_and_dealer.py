from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('TicketAppB', '0046_add_stage_fk_to_transactiondata'),
    ]

    operations = [
        migrations.AddField(
            model_name='company',
            name='district',
            field=models.CharField(blank=True, max_length=100, null=True),
        ),
        migrations.AddField(
            model_name='dealer',
            name='district',
            field=models.CharField(blank=True, max_length=100, null=True),
        ),
    ]
