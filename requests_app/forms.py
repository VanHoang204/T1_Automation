from django import forms


class RequestForm(forms.Form):
    name = forms.CharField(max_length=255, required=True)
    email = forms.EmailField(required=True)
    badge_id = forms.CharField(max_length=50, required=True)
    badge_type = forms.CharField(max_length=100, required=False)
    floor = forms.CharField(max_length=100, required=True)
    project = forms.CharField(max_length=255, required=False)
    status = forms.CharField(max_length=50, required=True, initial="Approved")

