from django.shortcuts import render, redirect
from django.conf import settings
from django.core.files.storage import FileSystemStorage
from django.http import HttpResponse

import subprocess
import os 
import time

from webapp.models import Document
from webapp.forms import DocumentForm
from webapp.Converter import Converter
from wsgiref.util import FileWrapper


def model_form_upload(request):
    if request.method == 'POST':
        form = DocumentForm(request.POST, request.FILES)

        if form.is_valid():

            path = os.path.join(os.getcwd(), 'webapp/documents')

            # delete old file in system
            subprocess.Popen(['rm', '-r', path])

            form.save()

            doc = request.FILES['document']

            filename = doc.name.replace(" ", "_")
            output_filename =  filename[:-5] + '.docx'

            time_to_wait = 10
            time_counter = 0

            while not os.path.exists(path):
                time.sleep(1)
                time_counter += 1
                if time_counter > time_to_wait:break

            Converter().convert(os.path.join(path, filename))

            time_to_wait = 10
            time_counter = 0

            while not os.path.exists(os.path.join(path, output_filename)):
                time.sleep(1)
                time_counter += 1
                if time_counter > time_to_wait:break

            with open(os.path.join(path, output_filename), "rb") as f:
                wrapper = FileWrapper(f)
                response = HttpResponse(wrapper, content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                response['Content-Disposition'] = 'attachment; filename=' + output_filename

                return response

        else:
            HttpResponse('File upload failed.')
            # redirect('model_form_upload')
    else:
        form = DocumentForm()
        
    return render(request, 'model_form_upload.html', {
        'form': form
    })