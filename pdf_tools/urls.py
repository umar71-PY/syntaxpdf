from django.urls import path
from django.views.generic import TemplateView
from . import views

app_name = "pdf_tools"

urlpatterns = [
    # Home Page ("") aur "/editor/" dono par editor.html khulega!
    path("", TemplateView.as_view(template_name="editor.html"), name="home"),
    path("editor/", TemplateView.as_view(template_name="editor.html"), name="editor"),

    path("merge/", views.merge_pdf, name="merge"),
    path("pdf-to-jpg/", views.pdf_to_jpg, name="pdf_to_jpg"),
    path("protect-pdf/", views.protect_pdf, name="protect_pdf"),
    path("compress-pdf/", views.compress_pdf, name="compress_pdf"),
    path("pdf-to-text/", views.pdf_to_text, name="pdf_to_text"),
    path("pdf-to-word/", views.pdf_to_word, name="pdf_to_word"),
    path("pdf-to-speech/", views.pdf_to_speech, name="pdf_to_speech"),
    path("split-pdf/", views.split_pdf, name="split_pdf"),
    path("unlock-pdf/", views.unlock_pdf, name="unlock_pdf"),
    path("pdf-to-excel/", views.pdf_to_excel, name="pdf_to_excel"),
    path("excel-to-pdf/", views.excel_to_pdf, name="excel_to_pdf"),
    path("csv-to-pdf/", views.csv_to_pdf, name="csv_to_pdf"),
    path("word-to-pdf/", views.word_to_pdf, name="word_to_pdf"),
    path("html-to-pdf/", views.html_to_pdf, name="html_to_pdf"),
    path("image-to-pdf/", views.image_to_pdf, name="image_to_pdf"),
    path("crop-pdf/", views.crop_pdf, name="crop_pdf"),
    path("edit-pdf/", views.edit_pdf, name="edit_pdf"),
    path("ocr-pdf/", views.ocr_pdf, name="ocr_pdf"),
    path("watermark-pdf/", views.watermark_pdf, name="watermark_pdf"),
    path("sign-pdf/", views.sign_pdf, name="sign_pdf"),
    path("powerpoint-to-pdf/", views.powerpoint_to_pdf, name="powerpoint_to_pdf"),
    path("add-page-numbers/", views.add_page_numbers, name="add_page_numbers"),
]