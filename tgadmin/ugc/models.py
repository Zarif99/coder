from django.db import models

# Create your models here.
class Profile(models.Model):
    external_id = models.PositiveIntegerField(
        verbose_name="ID User",
        unique=True,
    )
    name = models.TextField(
        verbose_name="Username",
    )
    def __str__(self):
        return f'#{self.external_id}{self.name}'

    class Meta:
        verbose_name="Account"
        verbose_name_plural ="Accounts"

class Message(models.Model):
    profile = models.ForeignKey(
        to = 'ugc.Profile',
        verbose_name='Account',
        on_delete=models.PROTECT,
    )
    text = models.TextField(verbose_name="Text")

    created_at = models.DateTimeField(
        verbose_name="Receipt time",
        auto_now_add= True,
    )
    def __str__(self):
        return f'Message{self.pk} from {self.profile}'
    class Meta:
        verbose_name="Message"
        verbose_name_plural="Messages"