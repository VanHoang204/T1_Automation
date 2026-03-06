from django.http import Http404, HttpRequest, HttpResponse
from django.db import models
from django.shortcuts import render, redirect
from django.urls import reverse
from django.utils import timezone

from .forms import RequestForm
from .models import FloorRequest


def home(request: HttpRequest) -> HttpResponse:
    records = FloorRequest.objects.all()

    query = request.GET.get("q", "").strip()
    floor_filter = request.GET.get("floor", "").strip()

    if query:
        q_lower = query.lower()
        records = records.filter(
            models.Q(name__icontains=q_lower) | models.Q(email__icontains=q_lower)
        )

    if floor_filter:
        records = records.filter(floor=floor_filter)

    floors = (
        FloorRequest.objects.exclude(floor="")
        .values_list("floor", flat=True)
        .distinct()
        .order_by("floor")
    )

    context = {
        "requests": records,
        "query": query,
        "floor_filter": floor_filter,
        "floors": floors,
    }
    return render(request, "requests_app/home.html", context)


def request_detail(request: HttpRequest, request_id: int) -> HttpResponse:
    try:
        record = FloorRequest.objects.get(pk=request_id)
    except FloorRequest.DoesNotExist:
        raise Http404("Request not found")

    if not record.is_read:
        record.is_read = True
        record.save(update_fields=["is_read"])

    context = {"request_obj": record}
    return render(request, "requests_app/detail.html", context)


def add_request(request: HttpRequest) -> HttpResponse:
    if request.method == "POST":
        form = RequestForm(request.POST)
        if form.is_valid():
            obj = FloorRequest.objects.create(
                name=form.cleaned_data["name"],
                email=form.cleaned_data["email"],
                badge_id=form.cleaned_data["badge_id"],
                badge_type=form.cleaned_data["badge_type"],
                floor=form.cleaned_data["floor"],
                project=form.cleaned_data["project"],
                mail_subject="",
                mail_entry_id=f"manual-{FloorRequest.objects.count() + 1}",
                request_time=timezone.now(),
            )
            return redirect(reverse("request_detail", args=[obj.pk]))
    else:
        form = RequestForm()

    context = {
        "form": form,
        "title": "Add Request",
        "submit_label": "Create",
    }
    return render(request, "requests_app/form.html", context)


def edit_request(request: HttpRequest, request_id: int) -> HttpResponse:
    try:
        record = FloorRequest.objects.get(pk=request_id)
    except FloorRequest.DoesNotExist:
        raise Http404("Request not found")

    if request.method == "POST":
        form = RequestForm(request.POST)
        if form.is_valid():
            record.name = form.cleaned_data["name"]
            record.email = form.cleaned_data["email"]
            record.badge_id = form.cleaned_data["badge_id"]
            record.badge_type = form.cleaned_data["badge_type"]
            record.floor = form.cleaned_data["floor"]
            record.project = form.cleaned_data["project"]
            record.save()
            return redirect(reverse("request_detail", args=[record.pk]))
    else:
        initial = {
            "name": record.customer_name,
            "email": record.email,
            "badge_id": record.badge_id,
            "badge_type": record.badge_type,
            "floor": record.floor,
            "project": record.project,
            "status": record.status,
        }
        form = RequestForm(initial=initial)

    context = {
        "form": form,
        "title": "Edit Request",
        "submit_label": "Save",
    }
    return render(request, "requests_app/form.html", context)


def delete_request(request: HttpRequest, request_id: int) -> HttpResponse:
    if request.method == "POST":
        FloorRequest.objects.filter(pk=request_id).delete()
        return redirect(reverse("home"))
    return redirect(reverse("home"))
