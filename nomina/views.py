import pandas as pd
from datetime import datetime, timedelta
from django.shortcuts import render
from django.http import HttpResponse, FileResponse
from django.contrib import messages
from .forms import ExcelUploadForm
import io

def index(request):
    if request.method == 'POST':
        form = ExcelUploadForm(request.POST, request.FILES)
        if form.is_valid():
            excel_file = request.FILES['excel_file']
            
            try:
                # Load the Excel file
                xls = pd.ExcelFile(excel_file)
                
                # Read the sheets ASISTENCIA and PROYECCION NOMINA into DataFrames
                asistencia_df = pd.read_excel(xls, 'ASISTENCIA')
                proyeccion_nomina_df = pd.read_excel(xls, 'proyeccion nomina')
                
                # Create a list of dates for the current month from the first to the last day
                today = datetime.now()
                current_month_start = datetime(today.year, today.month, 1)
                next_month = today.month % 12 + 1
                current_month_end = datetime(today.year if next_month > 1 else today.year + 1, next_month, 1) - timedelta(days=1)
                dates = pd.date_range(current_month_start, current_month_end)
                
                # Add the new columns with the dates to the DataFrames
                for date in dates:
                    asistencia_df[date.strftime('%Y-%m-%d')] = 0
                    proyeccion_nomina_df[date.strftime('%Y-%m-%d')] = 0
                
                # Insert the date columns after specific columns in each DataFrame
                asistencia_df = pd.concat([asistencia_df.iloc[:, :11], asistencia_df.iloc[:, 11:]], axis=1)
                proyeccion_nomina_df = pd.concat([proyeccion_nomina_df.iloc[:, :12], proyeccion_nomina_df.iloc[:, 12:]], axis=1)
                
                # Update the asistencia_df: Copy CANT RACIONES to the date columns where FECHA matches
                for index, row in asistencia_df.iterrows():
                    fecha = row['FECHA']
                    cant_raciones = row['CANT RACIONES']
                    fecha_str = fecha.strftime('%Y-%m-%d') if pd.notna(fecha) else None
                    if fecha_str in asistencia_df.columns:
                        asistencia_df.at[index, fecha_str] = cant_raciones
                
                # Update the proyeccion_nomina_df: Copy CANT RACIONES to the date columns within FECHAI and FECHA F range
                for index, row in proyeccion_nomina_df.iterrows():
                    fecha_i = row[9]
                    fecha_f = row[10]
                    cant_raciones = row[11]
                    if pd.notna(fecha_i) and pd.notna(fecha_f):
                        date_range = pd.date_range(fecha_i, fecha_f)
                        for date in date_range:
                            date_str = date.strftime('%Y-%m-%d')
                            if date_str in proyeccion_nomina_df.columns:
                                if date.weekday() < 5:  # Only set CANT RACIONES on weekdays
                                    proyeccion_nomina_df.at[index, date_str] = cant_raciones
                                else:  # Set value to 0 for weekends
                                    proyeccion_nomina_df.at[index, date_str] = 0
                
                # Filter and reorder columns for asistencia_df and proyeccion_nomina_df
                columns_to_keep = ['NOMBRE COLABORADOR', 'MODALIDAD', 'INSITUCION EDUCATIVA']
                columns_to_keep.extend([date.strftime('%Y-%m-%d') for date in dates if date.month == today.month])
                
                asistencia_filtered = asistencia_df[columns_to_keep]
                proyeccion_nomina_filtered = proyeccion_nomina_df[columns_to_keep]
                
                # Merge asistencia_filtered and proyeccion_nomina_filtered
                merged_df = pd.concat([asistencia_filtered, proyeccion_nomina_filtered])
                
                # Group by 'NOMBRE COLABORADOR', 'MODALIDAD', 'INSITUCION EDUCATIVA' and aggregate the date columns
                grouped_df = merged_df.groupby(['NOMBRE COLABORADOR', 'MODALIDAD', 'INSITUCION EDUCATIVA'], as_index=False).sum()
                
                # Save the resulting DataFrame to an Excel file
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    grouped_df.to_excel(writer, index=False, sheet_name='Processed Data')
                
                output.seek(0)
                
                # Create a FileResponse with the Excel file
                response = FileResponse(output, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                response['Content-Disposition'] = 'attachment; filename=processed_nomina.xlsx'
                
                return response
            
            except Exception as e:
                messages.error(request, f"An error occurred: {str(e)}")
    else:
        form = ExcelUploadForm()
    
    return render(request, 'nomina/index.html', {'form': form})
