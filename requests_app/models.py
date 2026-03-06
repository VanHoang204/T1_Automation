from django.db import models


class FloorRequest(models.Model):
    name = models.CharField(max_length=255)
    email = models.EmailField()
    badge_id = models.CharField(max_length=50)
    badge_type = models.CharField(max_length=100, blank=True)
    floor = models.CharField(max_length=100)
    project = models.CharField(max_length=255, blank=True)
    mail_subject = models.CharField(max_length=255, blank=True)
    mail_entry_id = models.CharField(max_length=255, unique=True)
    request_time = models.DateTimeField()
    created_at = models.DateTimeField(auto_now_add=True)
    is_read = models.BooleanField(default=False)

    class Meta:
        ordering = ["-created_at"]
