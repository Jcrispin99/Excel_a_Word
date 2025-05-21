from django import forms

class UploadForm(forms.Form):
    excel_file = forms.FileField(
        label='Archivo Excel',
        help_text='Sube tu archivo Excel con los datos de los productos.',
        widget=forms.FileInput(attrs={'class': 'form-control'})
    )
    images_zip = forms.FileField(
        label='Archivo ZIP con imágenes',
        help_text='Sube un archivo ZIP que contenga todas las imágenes de los productos.',
        widget=forms.FileInput(attrs={'class': 'form-control'})
    )
    
    def clean_excel_file(self):
        file = self.cleaned_data.get('excel_file')
        if file:
            if not file.name.endswith('.xlsx'):
                raise forms.ValidationError('El archivo debe ser un Excel (.xlsx)')
        return file
    
    def clean_images_zip(self):
        file = self.cleaned_data.get('images_zip')
        if file:
            if not file.name.endswith('.zip'):
                raise forms.ValidationError('El archivo debe ser un ZIP (.zip)')
        return file