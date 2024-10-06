from django import forms

class ExcelUploadForm(forms.Form):
    excel_file = forms.FileField(label='Select Excel File', help_text='Upload the Excel file containing ASISTENCIA and PROYECCION NOMINA sheets.')
