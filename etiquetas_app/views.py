import os
import uuid
import zipfile
import shutil
# tempfile no es estrictamente necesario si TEMP_DIR se define en settings
from django.shortcuts import render, redirect
from django.conf import settings
from django.http import FileResponse # HttpResponse no se usa directamente en procesar_archivos
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
            unique_id = str(uuid.uuid4())
            excel_file_uploaded = request.FILES['excel_file']
            zip_file_uploaded = request.FILES['images_zip']

            excel_filename = f"{unique_id}_{excel_file_uploaded.name}"
            zip_filename = f"{unique_id}_{zip_file_uploaded.name}"

            # Definir rutas base
            base_uploads_excel_path = os.path.join(settings.MEDIA_ROOT, 'uploads', 'excel')
            base_uploads_zip_path = os.path.join(settings.MEDIA_ROOT, 'uploads', 'zip')
            base_output_path = os.path.join(settings.MEDIA_ROOT, 'output')

            excel_path = os.path.join(base_uploads_excel_path, excel_filename)
            zip_path = os.path.join(base_uploads_zip_path, zip_filename)
            output_filename = f"{unique_id}_documento.docx"
            output_path = os.path.join(base_output_path, output_filename)

            temp_dir = None  # Inicializar temp_dir para el bloque finally
            archivos_a_limpiar = []  # Lista para seguimiento de archivos a eliminar

            try:
                # Asegurar que los directorios de carga y salida existan
                os.makedirs(base_uploads_excel_path, exist_ok=True)
                os.makedirs(base_uploads_zip_path, exist_ok=True)
                os.makedirs(base_output_path, exist_ok=True)

                # Guardar archivos subidos
                with open(excel_path, 'wb+') as destination:
                    for chunk in excel_file_uploaded.chunks():
                        destination.write(chunk)
                archivos_a_limpiar.append(excel_path)  # Añadir a la lista de limpieza
                
                with open(zip_path, 'wb+') as destination:
                    for chunk in zip_file_uploaded.chunks():
                        destination.write(chunk)
                archivos_a_limpiar.append(zip_path)  # Añadir a la lista de limpieza

                # Verificar la configuración de TEMP_DIR
                if not hasattr(settings, 'TEMP_DIR') or not settings.TEMP_DIR:
                    messages.error(request, "Error de configuración: TEMP_DIR no está definido en settings.py.")
                    # Limpiar archivos subidos si no podemos proceder
                    for archivo in archivos_a_limpiar:
                        if os.path.exists(archivo): 
                            os.remove(archivo)
                    return redirect('index')
                
                # Crear directorio temporal para la extracción de imágenes
                temp_dir = os.path.join(settings.TEMP_DIR, unique_id)
                os.makedirs(temp_dir, exist_ok=True)
                
                # Extraer archivos ZIP
                try:
                    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                        zip_ref.extractall(temp_dir)
                except zipfile.BadZipFile:
                    messages.error(request, f"El archivo ZIP '{zip_file_uploaded.name}' está corrupto o no es un ZIP válido.")
                    raise # Re-lanzar para ser capturado por el try-except principal y limpiar temp_dir
                except Exception as e_zip:
                    messages.error(request, f"Error al descomprimir el archivo ZIP '{zip_file_uploaded.name}': {str(e_zip)}")
                    raise # Re-lanzar

                # Generar documento Word
                generate_word_document(excel_path, temp_dir, output_path)
                
                # Guardar referencia en la base de datos si el usuario está autenticado
                if request.user.is_authenticated:
                    ArchivoGenerado.objects.create(
                        usuario=request.user,
                        excel_original=os.path.join('uploads', 'excel', excel_filename),
                        zip_original=os.path.join('uploads', 'zip', zip_filename),
                        documento_generado=os.path.join('output', output_filename)
                    )
                else:
                    # Si el usuario no está autenticado, podemos eliminar los archivos originales
                    # ya que no hay referencia en la base de datos
                    for archivo in archivos_a_limpiar:
                        if os.path.exists(archivo):
                            os.remove(archivo)
                
                request.session['documento_generado'] = output_path
                messages.success(request, "Archivos procesados y documento generado exitosamente.")
                return redirect('descargar')

            except Exception as e:
                messages.error(request, f"Error general al procesar los archivos: {str(e)}")
                return redirect('index')
            
            finally:
                # Limpiar el directorio temporal si fue creado
                if temp_dir and os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                    
                # Opcionalmente, limpiar archivos originales si no se guardaron en la base de datos
                # Esto se puede habilitar si se desea una limpieza agresiva
                # if not request.user.is_authenticated:
                #     for archivo in archivos_a_limpiar:
                #         if os.path.exists(archivo):
                #             os.remove(archivo)

        else:
            # Si el formulario no es válido, mostrar errores en la página de subida
            # Los mensajes de error del formulario se mostrarán automáticamente por la plantilla
            return render(request, 'index.html', {'form': form})
    
    # Si no es POST, redirigir a la página principal
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
