import os
import uuid
import zipfile
import shutil
import tempfile
from django.shortcuts import render, redirect
from django.conf import settings
from django.http import HttpResponse, FileResponse
from django.contrib import messages
from .forms import UploadForm
from .models import ArchivoGenerado
from .utils import generate_word_document

def index(request):
    form = UploadForm()
    return render(request, 'index.html', {'form': form})

def procesar_archivos(request):
    if request.method == 'POST':
        form = UploadForm(request.POST, request.FILES)
        if form.is_valid():
            # Generar nombres únicos para los archivos
            unique_id = str(uuid.uuid4())
            excel_filename = f"{unique_id}_{request.FILES['excel_file'].name}"
            zip_filename = f"{unique_id}_{request.FILES['images_zip'].name}"
            
            # Guardar archivos subidos
            excel_path = os.path.join(settings.UPLOAD_DIR, excel_filename)
            zip_path = os.path.join(settings.UPLOAD_DIR, zip_filename)
            
            with open(excel_path, 'wb+') as destination:
                for chunk in request.FILES['excel_file'].chunks():
                    destination.write(chunk)
                    
            with open(zip_path, 'wb+') as destination:
                for chunk in request.FILES['images_zip'].chunks():
                    destination.write(chunk)
            
            # Crear directorio temporal para extraer imágenes
            temp_dir = os.path.join(settings.TEMP_DIR, unique_id)
            os.makedirs(temp_dir, exist_ok=True)
            
            # Extraer archivos ZIP
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            # Generar documento Word
            output_filename = f"{unique_id}_documento.docx"
            output_path = os.path.join(settings.OUTPUT_DIR, output_filename)
            
            try:
                # Llamar a la función que genera el documento
                generate_word_document(excel_path, temp_dir, output_path)
                
                # Guardar referencia en la base de datos
                if request.user.is_authenticated:
                    ArchivoGenerado.objects.create(
                        usuario=request.user,
                        excel_original=excel_filename,
                        zip_original=zip_filename,
                        documento_generado=os.path.relpath(output_path, settings.MEDIA_ROOT)
                    )
                
                # Guardar ruta del archivo en la sesión para descarga
                request.session['documento_generado'] = output_path
                
                # Limpiar archivos temporales
                shutil.rmtree(temp_dir)
                
                return redirect('descargar')
            
            except Exception as e:
                messages.error(request, f"Error al procesar los archivos: {str(e)}")
                # Limpiar archivos en caso de error
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                return redirect('index')
        else:
            # Si el formulario no es válido, mostrar errores
            return render(request, 'index.html', {'form': form})
    
    return redirect('index')

def descargar(request):
    if 'documento_generado' not in request.session:
        messages.error(request, "No hay documento disponible para descargar")
        return redirect('index')
    
    return render(request, 'download.html')

def obtener_documento(request):
    if 'documento_generado' not in request.session:
        messages.error(request, "No hay documento disponible para descargar")
        return redirect('index')
    
    file_path = request.session['documento_generado']
    if os.path.exists(file_path):
        response = FileResponse(open(file_path, 'rb'))
        response['Content-Disposition'] = f'attachment; filename="documento.docx"'
        return response
    else:
        messages.error(request, "El archivo no existe")
        return redirect('index')
