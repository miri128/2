from django.db import models


class ItemProcess(models.Model):
    department = models.CharField("部門",max_length=2)
    item_name = models.CharField("社内品名",max_length=12)
    process_code = models.CharField("工程ｺｰﾄﾞ",max_length=6)
    process_name = models.CharField("工程名",max_length=20)
    total_qty = models.PositiveIntegerField("処理数")
    good_qty = models.PositiveIntegerField("良品数")
    board_total_qty = models.PositiveIntegerField("基板数")
    board_good_qty = models.PositiveIntegerField("基板良品数")
    created_at = models.CharField("作業日時",max_length=20)

    def __str__(self):
        return f"{self.item_name} - {self.process_name} ({self.department})"


