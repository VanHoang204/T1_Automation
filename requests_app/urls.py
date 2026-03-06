from django.urls import path
from . import views

urlpatterns = [
    path("", views.home, name="home"),
    path("processed", views.processed_list, name="processed_list"),
    path("request/<int:request_id>", views.request_detail, name="request_detail"),
    path("add", views.add_request, name="add_request"),
    path("edit/<int:request_id>", views.edit_request, name="edit_request"),
    path("delete/<int:request_id>", views.delete_request, name="delete_request"),
    path("mark-processed/<int:request_id>", views.mark_processed, name="mark_processed"),
]
