from django import forms


class UploadFileForm(forms.Form):
    file = forms.FileField(
        label='Excelファイル又はテキストファイルを選択',
        widget=forms.ClearableFileInput(attrs={"accept":".xlsx,.txt"})
    )