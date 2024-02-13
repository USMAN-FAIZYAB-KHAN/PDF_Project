from django.shortcuts import render
from django.shortcuts import render
from django.views.decorators.http import require_POST
from django.http import FileResponse
from .utils import *
import os

def index(request):
    return render(request, 'Pdf_to_Excel/index.html')

@require_POST
def upload_pdf(request):
    files = request.FILES.getlist('pdf_files')

    if not files:
        return render(request, 'Pdf_to_Excel/index.html', {'error': 'No files were uploaded'})

    file_list = []
    for file in files:
        if file.name.endswith('.pdf'):
            file_content = file.read()
            file_path = os.path.join('media', file.name)
            with open(file_path, 'wb') as f:
                f.write(file_content)
            file_list.append(file_path)
        else:
            return render(request, 'Pdf_to_Excel/index.html', {'error': 'Only PDF files are allowed'})
    
    cwp = get_categories_with_products()
    save_path = os.path.join('media', 'Point of Sale Analysis.xlsx')
    save_data_to_excel(file_list, save_path, cwp)

    for file_path in file_list:
        os.remove(file_path)    

    return FileResponse(open(save_path, 'rb'), as_attachment=True, filename='Point of Sale Analysis.xlsx')