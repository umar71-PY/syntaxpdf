import base64
import csv
import io
import json
import logging
import math
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from functools import wraps
from html.parser import HTMLParser
from pathlib import Path
from xml.sax.saxutils import escape as _xml_escape

from django.conf import settings
from django.http import FileResponse, JsonResponse
from django.utils.text import get_valid_filename
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_POST

from pypdf import PdfReader, PdfWriter
from pypdf.errors import DependencyError, PyPdfError
from pypdf.generic import RectangleObject
from pdf2image import convert_from_path, pdfinfo_from_path
from pdf2image.exceptions import (
    PDFInfoNotInstalledError,
    PDFPageCountError,
    PDFPopplerTimeoutError,
    PDFSyntaxError,
)

logger = logging.getLogger(__name__)

# ======================================================================
# Global Limits (Increased for Large Gazette PDFs & Multi-File Uploads)
# ======================================================================
MAX_FILES = getattr(settings, "PDF_MERGE_MAX_FILES", 50)
MAX_FILE_SIZE = getattr(settings, "PDF_MERGE_MAX_FILE_SIZE", 200 * 1024 * 1024)

JPG_MAX_FILE_SIZE = getattr(settings, "PDF_TO_JPG_MAX_FILE_SIZE", 150 * 1024 * 1024)
JPG_MAX_PAGES = getattr(settings, "PDF_TO_JPG_MAX_PAGES", 300)
JPG_DEFAULT_DPI = getattr(settings, "PDF_TO_JPG_DEFAULT_DPI", 150)
JPG_MAX_DPI = getattr(settings, "PDF_TO_JPG_MAX_DPI", 300)
JPG_QUALITY = getattr(settings, "PDF_TO_JPG_JPEG_QUALITY", 88)
POPPLER_TIMEOUT = getattr(settings, "PDF_TO_JPG_TIMEOUT", 120)

PROTECT_MAX_FILE_SIZE = getattr(settings, "PDF_PROTECT_MAX_FILE_SIZE", 150 * 1024 * 1024)
PROTECT_MAX_PASSWORD_LEN = getattr(settings, "PDF_PROTECT_MAX_PASSWORD_LENGTH", 128)
COMPRESS_MAX_FILE_SIZE = getattr(settings, "PDF_COMPRESS_MAX_FILE_SIZE", 200 * 1024 * 1024)
COMPRESS_TIMEOUT = getattr(settings, "PDF_COMPRESS_TIMEOUT", 180)
GS_BINARY = getattr(settings, "GHOSTSCRIPT_BINARY", "gs")

COMPRESS_LEVELS = {
    "low": "/printer",
    "medium": "/ebook",
    "high": "/screen",
}

TEXT_MAX_FILE_SIZE = getattr(settings, "PDF_TEXT_MAX_FILE_SIZE", 150 * 1024 * 1024)
TEXT_MAX_PAGES = getattr(settings, "PDF_TEXT_MAX_PAGES", 2000)
WORD_MAX_FILE_SIZE = getattr(settings, "PDF_WORD_MAX_FILE_SIZE", 100 * 1024 * 1024)
WORD_MAX_PAGES = getattr(settings, "PDF_WORD_MAX_PAGES", 300)
SPEECH_MAX_FILE_SIZE = getattr(settings, "PDF_SPEECH_MAX_FILE_SIZE", 100 * 1024 * 1024)
SPEECH_MAX_CHARS = getattr(settings, "PDF_SPEECH_MAX_CHARS", 50000)
SPEECH_TIMEOUT = getattr(settings, "PDF_SPEECH_TIMEOUT", 60)

SPLIT_MAX_FILE_SIZE = getattr(settings, "PDF_SPLIT_MAX_FILE_SIZE", 200 * 1024 * 1024)
SPLIT_MAX_PAGES = getattr(settings, "PDF_SPLIT_MAX_PAGES", 10000)
SPLIT_MAX_RANGES = getattr(settings, "PDF_SPLIT_MAX_RANGES", 500)
SPLIT_MAX_PAGES_STRING = getattr(settings, "PDF_SPLIT_MAX_PAGES_STRING", 2000)
UNLOCK_MAX_FILE_SIZE = getattr(settings, "PDF_UNLOCK_MAX_FILE_SIZE", 200 * 1024 * 1024)
UNLOCK_MAX_PASSWORD_LEN = getattr(settings, "PDF_UNLOCK_MAX_PASSWORD_LENGTH", 128)

EXCEL_MAX_FILE_SIZE = getattr(settings, "PDF_EXCEL_MAX_FILE_SIZE", 100 * 1024 * 1024)
EXCEL_MAX_PAGES = getattr(settings, "PDF_EXCEL_MAX_PAGES", 500)
EXCEL_MAX_CELLS = getattr(settings, "PDF_EXCEL_MAX_CELLS", 300_000)
XLSX_MAX_FILE_SIZE = getattr(settings, "EXCEL_PDF_MAX_FILE_SIZE", 50 * 1024 * 1024)
XLSX_MAX_SHEETS = getattr(settings, "EXCEL_PDF_MAX_SHEETS", 30)
XLSX_MAX_CELLS = getattr(settings, "EXCEL_PDF_MAX_CELLS", 100_000)
XLSX_MAX_CELL_CHARS = getattr(settings, "EXCEL_PDF_MAX_CELL_CHARS", 80)

DOCX_MAX_FILE_SIZE = getattr(settings, "WORD_PDF_MAX_FILE_SIZE", 50 * 1024 * 1024)
PPTX_MAX_FILE_SIZE = getattr(settings, "PPTX_PDF_MAX_FILE_SIZE", 100 * 1024 * 1024)
PPTX_MAX_SLIDES = getattr(settings, "PPTX_PDF_MAX_SLIDES", 300)
CSV_MAX_FILE_SIZE = getattr(settings, "CSV_PDF_MAX_FILE_SIZE", 50 * 1024 * 1024)
HTML_MAX_FILE_SIZE = getattr(settings, "HTML_PDF_MAX_FILE_SIZE", 20 * 1024 * 1024)

OCR_MAX_FILE_SIZE = getattr(settings, "PDF_OCR_MAX_FILE_SIZE", 100 * 1024 * 1024)
OCR_MAX_PAGES = getattr(settings, "PDF_OCR_MAX_PAGES", 100)
OCR_TIMEOUT = getattr(settings, "PDF_OCR_TIMEOUT", 180)
WATERMARK_MAX_FILE_SIZE = getattr(settings, "PDF_WATERMARK_MAX_FILE_SIZE", 150 * 1024 * 1024)
SIGN_MAX_FILE_SIZE = getattr(settings, "PDF_SIGN_MAX_FILE_SIZE", 150 * 1024 * 1024)
PAGENUM_MAX_FILE_SIZE = getattr(settings, "PDF_PAGENUM_MAX_FILE_SIZE", 200 * 1024 * 1024)

PDF_FONT_PATH = getattr(settings, "PDF_FONT_PATH", None)
IMAGE_MAX_FILE_SIZE = getattr(settings, "IMAGE_PDF_MAX_FILE_SIZE", 50 * 1024 * 1024)
IMAGE_MAX_FILES = getattr(settings, "IMAGE_PDF_MAX_FILES", 50)
IMAGE_MAX_PIXELS = getattr(settings, "IMAGE_PDF_MAX_PIXELS", 80_000_000)
CROP_MAX_FILE_SIZE = getattr(settings, "PDF_CROP_MAX_FILE_SIZE", 200 * 1024 * 1024)
CROP_MAX_PAGES = getattr(settings, "PDF_CROP_MAX_PAGES", 10000)

XLSX_EXTENSIONS = {".xlsx", ".xlsm"}
DOCX_EXTENSIONS = {".docx"}
PPTX_EXTENSIONS = {".pptx", ".ppt"}
CSV_EXTENSIONS = {".csv", ".tsv", ".txt"}
HTML_EXTENSIONS = {".html", ".htm"}
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif", ".tif", ".tiff"}
EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

NO_TEXT_MESSAGE = (
    "No extractable text found. The PDF is probably scanned images and needs OCR first."
)

_PAGE_TOKEN_RE = re.compile(r"^(\d{1,6})(?:-(\d{1,6}))?$")


# ======================================================================
# Optional Clerk Auth Guard (Enforced if CLERK_REQUIRE_AUTH=True in settings)
# ======================================================================
def clerk_authenticated(view_func):
    @wraps(view_func)
    def _wrapped(request, *args, **kwargs):
        require_auth = getattr(settings, "CLERK_REQUIRE_AUTH", False)
        auth_header = request.headers.get("Authorization", "")
        if require_auth:
            if not auth_header.startswith("Bearer ") or len(auth_header.split(" ", 1)[1].strip()) < 20:
                return JsonResponse({"error": "Authentication required. Please sign in."}, status=401)
        response = view_func(request, *args, **kwargs)
        response["Access-Control-Expose-Headers"] = "Content-Disposition, X-Original-Size, X-Compressed-Size"
        return response
    return _wrapped


class PDFMergeError(Exception):
    """Raised for user-facing validation/processing errors."""
    def __init__(self, message, status=400):
        super().__init__(message)
        self.message = message
        self.status = status


class SelfCleaningFileResponse(FileResponse):
    """FileResponse that deletes its temp directory once the response is closed."""
    def __init__(self, *args, cleanup_dir=None, **kwargs):
        self._cleanup_dir = cleanup_dir
        super().__init__(*args, **kwargs)

    def close(self):
        try:
            super().close()
        finally:
            if self._cleanup_dir:
                shutil.rmtree(self._cleanup_dir, ignore_errors=True)
                self._cleanup_dir = None


# ======================================================================
# Core Validation & File Helpers
# ======================================================================
def _save_pdf_upload(upload, dest_path, max_size=MAX_FILE_SIZE):
    if not upload.name.lower().endswith(".pdf"):
        raise PDFMergeError(f"'{upload.name}' is not a .pdf file.", status=415)
    if upload.size == 0:
        raise PDFMergeError(f"'{upload.name}' is empty.")
    if upload.size > max_size:
        raise PDFMergeError(
            f"'{upload.name}' exceeds the {max_size // (1024 * 1024)} MB limit.", status=413
        )

    header = upload.read(1024)
    upload.seek(0)
    if b"%PDF-" not in header[:1024]:
        raise PDFMergeError(f"'{upload.name}' is not a valid PDF file.", status=415)

    with open(dest_path, "wb") as out:
        for chunk in upload.chunks():
            out.write(chunk)


def _parse_dpi(raw_value):
    if not raw_value:
        return JPG_DEFAULT_DPI
    try:
        dpi = int(raw_value)
    except ValueError:
        raise PDFMergeError("'dpi' must be an integer.")
    if not 72 <= dpi <= JPG_MAX_DPI:
        raise PDFMergeError(f"'dpi' must be between 72 and {JPG_MAX_DPI}.")
    return dpi


def _parse_page_ranges(raw):
    raw = (raw or "").strip()
    if not raw:
        raise PDFMergeError("'pages' is required, for example: 1,3,5-7.")
    if len(raw) > SPLIT_MAX_PAGES_STRING:
        raise PDFMergeError("'pages' value is too long.")

    normalized = re.sub(r"[.;\s]+", ",", raw.strip())
    tokens = [t.strip() for t in normalized.split(",") if t.strip()]

    if not tokens:
        raise PDFMergeError("Provide at least one page number, e.g. 1,3,5-7.")
    if len(tokens) > SPLIT_MAX_RANGES:
        raise PDFMergeError(f"Too many page ranges; the maximum is {SPLIT_MAX_RANGES}.")

    ranges = []
    for token in tokens:
        match = _PAGE_TOKEN_RE.match(token)
        if not match:
            raise PDFMergeError(
                f"Invalid page selection '{token}'. Use numbers or ranges like 1,3,5-7."
            )
        start = int(match.group(1))
        end = int(match.group(2)) if match.group(2) else start
        if start < 1 or end < 1:
            raise PDFMergeError("Page numbers start at 1.")
        if end < start:
            start, end = end, start
        ranges.append((start, end))
    return ranges


def _expand_page_ranges(ranges, total_pages):
    indexes, seen = [], set()
    for start, end in ranges:
        if end > total_pages:
            bad = end if start <= total_pages else start
            raise PDFMergeError(
                f"Page {bad} is out of range; the PDF has {total_pages} page(s)."
            )
        for page_no in range(start, end + 1):
            if page_no not in seen:
                seen.add(page_no)
                indexes.append(page_no - 1)
    return indexes


def _run_single_pdf_tool(request, *, prefix, max_size, processor):
    upload = request.FILES.get("file")
    if upload is None:
        return JsonResponse({"error": "No file uploaded. Use the 'file' field."}, status=400)

    work_dir = tempfile.mkdtemp(prefix=prefix, dir=getattr(settings, "PDF_TEMP_DIR", None))
    response = None

    try:
        input_path = os.path.join(work_dir, "input.pdf")
        _save_pdf_upload(upload, input_path, max_size)

        output_path, extension, content_type = processor(request, input_path, work_dir)

        if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
            raise PDFMergeError("Conversion produced no output. The PDF may be corrupted.")

        base_name = get_valid_filename(Path(upload.name).stem) or "document"
        response = SelfCleaningFileResponse(
            open(output_path, "rb"),
            as_attachment=True,
            filename=f"{base_name}{extension}",
            content_type=content_type,
            cleanup_dir=work_dir,
        )
        return response

    except PDFMergeError as exc:
        return JsonResponse({"error": exc.message}, status=exc.status)
    except Exception:
        logger.exception("Unexpected error in %s tool", prefix.rstrip("_"))
        return JsonResponse({"error": "Internal error while processing the PDF."}, status=500)
    finally:
        if response is None:
            shutil.rmtree(work_dir, ignore_errors=True)


def _check_pdf_readable(pdf_path, max_pages):
    try:
        reader = PdfReader(pdf_path)
        if reader.is_encrypted:
            raise PDFMergeError("The PDF is password protected. Unlock it first.")
        page_count = len(reader.pages)
    except PDFMergeError:
        raise
    except (PyPdfError, ValueError, KeyError, OSError, RecursionError) as exc:
        logger.warning("Unreadable PDF: %s", exc)
        raise PDFMergeError("The PDF is corrupted or unreadable.")

    if page_count == 0:
        raise PDFMergeError("The PDF contains no pages.")
    if page_count > max_pages:
        raise PDFMergeError(
            f"PDF has {page_count} pages; the maximum is {max_pages}.", status=413
        )
    return reader


def _extract_pdf_text(pdf_path, max_pages, page_indexes=None):
    reader = _check_pdf_readable(pdf_path, max_pages)
    target_indexes = page_indexes if page_indexes is not None else range(len(reader.pages))

    pages = []
    for idx in target_indexes:
        try:
            pages.append((reader.pages[idx].extract_text() or "").strip())
        except (PyPdfError, ValueError, KeyError, OSError, RecursionError) as exc:
            logger.warning("Text extraction failed on page %s: %s", idx + 1, exc)
            raise PDFMergeError(f"Page {idx + 1} is corrupted and its text cannot be read.")

    if not any(pages):
        raise PDFMergeError(NO_TEXT_MESSAGE, status=422)
    return pages


# ======================================================================
# TOOL 1: MERGE PDF
# ======================================================================
@csrf_exempt
@require_POST
@clerk_authenticated
def merge_pdf(request):
    uploads = request.FILES.getlist("files")
    if not uploads:
        return JsonResponse({"error": "No files uploaded. Use the 'files' field."}, status=400)
    if len(uploads) < 2:
        return JsonResponse({"error": "Upload at least two PDF files to merge."}, status=400)
    if len(uploads) > MAX_FILES:
        return JsonResponse({"error": f"Maximum {MAX_FILES} files allowed."}, status=400)

    work_dir = tempfile.mkdtemp(prefix="merge_", dir=getattr(settings, "PDF_TEMP_DIR", None))
    response = None

    try:
        saved_paths = []
        for index, upload in enumerate(uploads):
            path = os.path.join(work_dir, f"{index:03d}.pdf")
            _save_pdf_upload(upload, path, MAX_FILE_SIZE)
            saved_paths.append((upload.name, path))

        writer = PdfWriter()
        try:
            for original_name, path in saved_paths:
                try:
                    reader = PdfReader(path)
                    if reader.is_encrypted:
                        raise PDFMergeError(f"'{original_name}' is password protected. Unlock it first.")
                    if len(reader.pages) == 0:
                        raise PDFMergeError(f"'{original_name}' contains no pages.")
                    writer.append(reader)
                except PDFMergeError:
                    raise
                except (PyPdfError, ValueError, KeyError, OSError, RecursionError) as exc:
                    logger.warning("Corrupted PDF %s: %s", original_name, exc)
                    raise PDFMergeError(f"'{original_name}' is corrupted or unreadable.")

            merged_path = os.path.join(work_dir, "merged.pdf")
            with open(merged_path, "wb") as merged_file:
                writer.write(merged_file)
        finally:
            writer.close()

        response = SelfCleaningFileResponse(
            open(merged_path, "rb"),
            as_attachment=True,
            filename="merged.pdf",
            content_type="application/pdf",
            cleanup_dir=work_dir,
        )
        return response

    except PDFMergeError as exc:
        return JsonResponse({"error": exc.message}, status=exc.status)
    except Exception:
        logger.exception("Unexpected error while merging PDFs")
        return JsonResponse({"error": "Internal error while merging PDFs."}, status=500)
    finally:
        if response is None:
            shutil.rmtree(work_dir, ignore_errors=True)


# ======================================================================
# TOOL 2: PDF TO JPG
# ======================================================================
@csrf_exempt
@require_POST
@clerk_authenticated
def pdf_to_jpg(request):
    upload = request.FILES.get("file")
    if upload is None:
        return JsonResponse({"error": "No file uploaded. Use the 'file' field."}, status=400)

    work_dir = tempfile.mkdtemp(prefix="pdf2jpg_", dir=getattr(settings, "PDF_TEMP_DIR", None))
    response = None

    try:
        dpi = _parse_dpi(request.POST.get("dpi"))
        pdf_path = os.path.join(work_dir, "input.pdf")
        _save_pdf_upload(upload, pdf_path, JPG_MAX_FILE_SIZE)

        base_name = get_valid_filename(Path(upload.name).stem) or "document"
        raw_pages = (request.POST.get("pages") or "").strip()

        try:
            fitz = _import_fitz()
            doc = fitz.open(pdf_path)
            if doc.needs_pass:
                doc.close()
                raise PDFMergeError("The PDF is password protected. Unlock it first.")
            page_count = doc.page_count
            if page_count < 1:
                doc.close()
                raise PDFMergeError("The PDF contains no pages.")
            if page_count > JPG_MAX_PAGES:
                doc.close()
                raise PDFMergeError(f"PDF has {page_count} pages; maximum is {JPG_MAX_PAGES}.", status=413)

            target_indexes = (
                _expand_page_ranges(_parse_page_ranges(raw_pages), page_count)
                if raw_pages else list(range(page_count))
            )

            zoom = dpi / 72.0
            mat = fitz.Matrix(zoom, zoom)

            if len(target_indexes) == 1:
                single_jpg = os.path.join(work_dir, "output.jpg")
                pix = doc[target_indexes[0]].get_pixmap(matrix=mat, alpha=False)
                pix.pil_save(single_jpg, format="JPEG", quality=JPG_QUALITY, optimize=True)
                doc.close()
                response = SelfCleaningFileResponse(
                    open(single_jpg, "rb"),
                    as_attachment=True,
                    filename=f"{base_name}.jpg",
                    content_type="image/jpeg",
                    cleanup_dir=work_dir,
                )
                return response

            zip_path = os.path.join(work_dir, "images.zip")
            page_tmp = os.path.join(work_dir, "page.jpg")
            with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_STORED) as zf:
                for idx in target_indexes:
                    pix = doc[idx].get_pixmap(matrix=mat, alpha=False)
                    pix.pil_save(page_tmp, format="JPEG", quality=JPG_QUALITY, optimize=True)
                    zf.write(page_tmp, arcname=f"{base_name}_page_{idx + 1:03d}.jpg")
                    os.remove(page_tmp)
            doc.close()

            response = SelfCleaningFileResponse(
                open(zip_path, "rb"),
                as_attachment=True,
                filename=f"{base_name}_images.zip",
                content_type="application/zip",
                cleanup_dir=work_dir,
            )
            return response
        except PDFMergeError:
            raise
        except Exception:
            pass

        try:
            info = pdfinfo_from_path(pdf_path, timeout=POPPLER_TIMEOUT)
            page_count = int(info["Pages"])
        except PDFInfoNotInstalledError:
            logger.error("Poppler is not installed on this server.")
            raise PDFMergeError("Server is missing PDF conversion tools.", status=500)
        except PDFPopplerTimeoutError:
            raise PDFMergeError("PDF analysis timed out.", status=408)
        except (PDFPageCountError, PDFSyntaxError, KeyError, ValueError):
            raise PDFMergeError("The PDF is corrupted, unreadable, or password protected.")

        if page_count < 1:
            raise PDFMergeError("The PDF contains no pages.")
        if page_count > JPG_MAX_PAGES:
            raise PDFMergeError(f"PDF has {page_count} pages; the maximum is {JPG_MAX_PAGES}.", status=413)

        target_indexes = (
            _expand_page_ranges(_parse_page_ranges(raw_pages), page_count)
            if raw_pages else list(range(page_count))
        )

        if len(target_indexes) == 1:
            page_no = target_indexes[0] + 1
            images = convert_from_path(pdf_path, dpi=dpi, first_page=page_no, last_page=page_no, timeout=POPPLER_TIMEOUT)
            single_jpg = os.path.join(work_dir, "output.jpg")
            image = images[0].convert("RGB")
            image.save(single_jpg, "JPEG", quality=JPG_QUALITY, optimize=True)
            image.close()
            response = SelfCleaningFileResponse(
                open(single_jpg, "rb"),
                as_attachment=True,
                filename=f"{base_name}.jpg",
                content_type="image/jpeg",
                cleanup_dir=work_dir,
            )
            return response

        zip_path = os.path.join(work_dir, "images.zip")
        page_tmp = os.path.join(work_dir, "page.jpg")

        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_STORED) as zf:
            for idx in target_indexes:
                page_no = idx + 1
                images = convert_from_path(
                    pdf_path, dpi=dpi, first_page=page_no, last_page=page_no, timeout=POPPLER_TIMEOUT,
                )
                image = images[0].convert("RGB")
                try:
                    image.save(page_tmp, "JPEG", quality=JPG_QUALITY, optimize=True)
                finally:
                    image.close()
                    for img in images:
                        img.close()

                zf.write(page_tmp, arcname=f"{base_name}_page_{page_no:03d}.jpg")
                os.remove(page_tmp)

        response = SelfCleaningFileResponse(
            open(zip_path, "rb"),
            as_attachment=True,
            filename=f"{base_name}_images.zip",
            content_type="application/zip",
            cleanup_dir=work_dir,
        )
        return response

    except PDFMergeError as exc:
        return JsonResponse({"error": exc.message}, status=exc.status)
    except Exception:
        logger.exception("Unexpected error during PDF to JPG conversion")
        return JsonResponse({"error": "Internal error while converting the PDF."}, status=500)
    finally:
        if response is None:
            shutil.rmtree(work_dir, ignore_errors=True)


# ======================================================================
# TOOL 3: PROTECT PDF
# ======================================================================
@csrf_exempt
@require_POST
@clerk_authenticated
def protect_pdf(request):
    upload = request.FILES.get("file")
    password = request.POST.get("password", "")

    if upload is None:
        return JsonResponse({"error": "No file uploaded. Use the 'file' field."}, status=400)
    if not password:
        return JsonResponse({"error": "A 'password' is required."}, status=400)
    if len(password) > PROTECT_MAX_PASSWORD_LEN:
        return JsonResponse(
            {"error": f"Password must be at most {PROTECT_MAX_PASSWORD_LEN} characters."},
            status=400,
        )

    work_dir = tempfile.mkdtemp(prefix="protect_", dir=getattr(settings, "PDF_TEMP_DIR", None))
    response = None

    try:
        input_path = os.path.join(work_dir, "input.pdf")
        _save_pdf_upload(upload, input_path, PROTECT_MAX_FILE_SIZE)

        try:
            reader = PdfReader(input_path)
            if reader.is_encrypted:
                raise PDFMergeError("This PDF is already password protected.")
            if len(reader.pages) == 0:
                raise PDFMergeError("The PDF contains no pages.")
            writer = PdfWriter(clone_from=reader)
        except PDFMergeError:
            raise
        except (PyPdfError, ValueError, KeyError, OSError, RecursionError) as exc:
            logger.warning("Unreadable PDF for protection: %s", exc)
            raise PDFMergeError("The PDF is corrupted or unreadable.")

        output_path = os.path.join(work_dir, "protected.pdf")
        try:
            writer.encrypt(user_password=password, algorithm="AES-256")
            with open(output_path, "wb") as out:
                writer.write(out)
        except ImportError:
            logger.error("'cryptography' package is missing.")
            raise PDFMergeError("Server encryption support is not configured.", status=500)
        except (PyPdfError, ValueError, OSError) as exc:
            logger.warning("Encryption failed: %s", exc)
            raise PDFMergeError("Failed to encrypt the PDF.", status=500)
        finally:
            writer.close()

        base_name = get_valid_filename(Path(upload.name).stem) or "document"
        response = SelfCleaningFileResponse(
            open(output_path, "rb"),
            as_attachment=True,
            filename=f"{base_name}_protected.pdf",
            content_type="application/pdf",
            cleanup_dir=work_dir,
        )
        return response

    except PDFMergeError as exc:
        return JsonResponse({"error": exc.message}, status=exc.status)
    except Exception:
        logger.exception("Unexpected error while protecting PDF")
        return JsonResponse({"error": "Internal error while protecting the PDF."}, status=500)
    finally:
        if response is None:
            shutil.rmtree(work_dir, ignore_errors=True)


# ======================================================================
# TOOL 4: COMPRESS PDF
# ======================================================================
@csrf_exempt
@require_POST
@clerk_authenticated
def compress_pdf(request):
    upload = request.FILES.get("file")
    if upload is None:
        return JsonResponse({"error": "No file uploaded. Use the 'file' field."}, status=400)

    level = request.POST.get("level", "medium").lower()
    if level not in COMPRESS_LEVELS:
        return JsonResponse(
            {"error": f"'level' must be one of: {', '.join(COMPRESS_LEVELS)}."}, status=400
        )

    work_dir = tempfile.mkdtemp(prefix="compress_", dir=getattr(settings, "PDF_TEMP_DIR", None))
    response = None

    try:
        input_path = os.path.join(work_dir, "input.pdf")
        output_path = os.path.join(work_dir, "output.pdf")
        _save_pdf_upload(upload, input_path, COMPRESS_MAX_FILE_SIZE)

        try:
            if PdfReader(input_path).is_encrypted:
                raise PDFMergeError("Password-protected PDFs cannot be compressed. Unlock it first.")
        except PDFMergeError:
            raise
        except Exception:
            pass

        cmd = [
            GS_BINARY,
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            f"-dPDFSETTINGS={COMPRESS_LEVELS[level]}",
            "-dNOPAUSE",
            "-dBATCH",
            "-dQUIET",
            "-dSAFER",
            f"-sOutputFile={output_path}",
            input_path,
        ]

        try:
            result = subprocess.run(
                cmd,
                cwd=work_dir,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=COMPRESS_TIMEOUT,
                check=False,
            )
            gs_ok = result.returncode == 0 and os.path.exists(output_path) and os.path.getsize(output_path) > 0
        except FileNotFoundError:
            gs_ok = False
        except subprocess.TimeoutExpired:
            raise PDFMergeError("Compression timed out. Try a smaller file.", status=408)

        if not gs_ok:
            try:
                fitz = _import_fitz()
                doc = fitz.open(input_path)
                doc.save(output_path, garbage=4, deflate=True, clean=True)
                doc.close()
            except Exception:
                raise PDFMergeError("The PDF is corrupted or could not be compressed.")

        original_size = os.path.getsize(input_path)
        compressed_size = os.path.getsize(output_path)
        served_path = output_path if compressed_size < original_size else input_path
        served_size = min(compressed_size, original_size)

        base_name = get_valid_filename(Path(upload.name).stem) or "document"
        response = SelfCleaningFileResponse(
            open(served_path, "rb"),
            as_attachment=True,
            filename=f"{base_name}_compressed.pdf",
            content_type="application/pdf",
            cleanup_dir=work_dir,
        )
        response["X-Original-Size"] = str(original_size)
        response["X-Compressed-Size"] = str(served_size)
        return response

    except PDFMergeError as exc:
        return JsonResponse({"error": exc.message}, status=exc.status)
    except Exception:
        logger.exception("Unexpected error while compressing PDF")
        return JsonResponse({"error": "Internal error while compressing the PDF."}, status=500)
    finally:
        if response is None:
            shutil.rmtree(work_dir, ignore_errors=True)


# ======================================================================
# TOOL 5: PDF TO TEXT
# ======================================================================
def _process_pdf_to_text(request, input_path, work_dir):
    pages = _extract_pdf_text(input_path, TEXT_MAX_PAGES)
    output_path = os.path.join(work_dir, "output.txt")
    with open(output_path, "w", encoding="utf-8") as out:
        out.write("\n\n".join(pages))
    return output_path, ".txt", "text/plain; charset=utf-8"


@csrf_exempt
@require_POST
@clerk_authenticated
def pdf_to_text(request):
    return _run_single_pdf_tool(
        request, prefix="pdf2txt_", max_size=TEXT_MAX_FILE_SIZE, processor=_process_pdf_to_text
    )


# ======================================================================
# TOOL 6: PDF TO WORD (pdf2docx)
# ======================================================================
def _process_pdf_to_word(request, input_path, work_dir):
    _check_pdf_readable(input_path, WORD_MAX_PAGES)
    try:
        from pdf2docx import Converter
    except ImportError:
        logger.error("pdf2docx is not installed.")
        raise PDFMergeError("Server is missing PDF to Word conversion tools.", status=500)

    output_path = os.path.join(work_dir, "output.docx")
    converter = None
    try:
        converter = Converter(input_path)
        converter.convert(output_path, start=0, end=None)
    except Exception as exc:
        logger.warning("pdf2docx conversion failed: %s", exc)
        raise PDFMergeError("The PDF could not be converted. It may be corrupted.")
    finally:
        if converter is not None:
            try:
                converter.close()
            except Exception:
                pass

    return (
        output_path,
        ".docx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@csrf_exempt
@require_POST
@clerk_authenticated
def pdf_to_word(request):
    return _run_single_pdf_tool(
        request, prefix="pdf2word_", max_size=WORD_MAX_FILE_SIZE, processor=_process_pdf_to_word
    )


# ======================================================================
# TOOL 7: PDF TO SPEECH
# ======================================================================
def _process_pdf_to_speech(request, input_path, work_dir):
    try:
        from gtts import gTTS
        from gtts.lang import tts_langs
        from gtts.tts import gTTSError
    except ImportError:
        logger.error("gTTS is not installed.")
        raise PDFMergeError("Server is missing text-to-speech tools.", status=500)

    lang = request.POST.get("lang", "en").strip().lower()
    if lang not in tts_langs():
        raise PDFMergeError(f"Unsupported language '{lang}'.")

    raw_pages = (request.POST.get("pages") or "").strip()
    reader = _check_pdf_readable(input_path, TEXT_MAX_PAGES)
    target_indexes = (
        _expand_page_ranges(_parse_page_ranges(raw_pages), len(reader.pages))
        if raw_pages else range(len(reader.pages))
    )

    pages = _extract_pdf_text(input_path, TEXT_MAX_PAGES, page_indexes=target_indexes)
    text = " ".join(" ".join(pages).split())
    if not text:
        raise PDFMergeError(NO_TEXT_MESSAGE, status=422)
    if len(text) > SPEECH_MAX_CHARS:
        text = text[:SPEECH_MAX_CHARS]

    output_path = os.path.join(work_dir, "output.mp3")
    try:
        gTTS(text=text, lang=lang, timeout=SPEECH_TIMEOUT).save(output_path)
    except AssertionError:
        raise PDFMergeError(NO_TEXT_MESSAGE, status=422)
    except gTTSError as exc:
        logger.warning("gTTS request failed: %s", exc)
        raise PDFMergeError("The speech service is unavailable. Try again later.", status=502)
    except (ValueError, OSError) as exc:
        logger.warning("Speech generation failed: %s", exc)
        raise PDFMergeError("Failed to generate audio from this PDF.", status=500)

    return output_path, ".mp3", "audio/mpeg"


@csrf_exempt
@require_POST
@clerk_authenticated
def pdf_to_speech(request):
    return _run_single_pdf_tool(
        request, prefix="pdf2speech_", max_size=SPEECH_MAX_FILE_SIZE, processor=_process_pdf_to_speech
    )


# ======================================================================
# TOOL 8: SPLIT PDF
# ======================================================================
def _process_split_pdf(request, input_path, work_dir):
    ranges = _parse_page_ranges(request.POST.get("pages"))
    reader = _check_pdf_readable(input_path, SPLIT_MAX_PAGES)
    indexes = _expand_page_ranges(ranges, len(reader.pages))

    output_path = os.path.join(work_dir, "output.pdf")
    writer = PdfWriter()
    try:
        for index in indexes:
            writer.add_page(reader.pages[index])
        with open(output_path, "wb") as out:
            writer.write(out)
    except (PyPdfError, ValueError, KeyError, IndexError, OSError, RecursionError) as exc:
        logger.warning("Page extraction failed: %s", exc)
        raise PDFMergeError("Some selected pages are corrupted and cannot be extracted.")
    finally:
        writer.close()

    return output_path, "_split.pdf", "application/pdf"


@csrf_exempt
@require_POST
@clerk_authenticated
def split_pdf(request):
    return _run_single_pdf_tool(
        request, prefix="split_", max_size=SPLIT_MAX_FILE_SIZE, processor=_process_split_pdf
    )


# ======================================================================
# TOOL 9: UNLOCK PDF
# ======================================================================
def _process_unlock_pdf(request, input_path, work_dir):
    password = request.POST.get("password")
    if password is None:
        raise PDFMergeError("A 'password' field is required.")
    if len(password) > UNLOCK_MAX_PASSWORD_LEN:
        raise PDFMergeError(f"Password must be at most {UNLOCK_MAX_PASSWORD_LEN} characters.")

    try:
        reader = PdfReader(input_path)
        if not reader.is_encrypted:
            raise PDFMergeError("This PDF is not password protected.")
        if not reader.decrypt(password):
            raise PDFMergeError("Incorrect password.")
        if len(reader.pages) == 0:
            raise PDFMergeError("The PDF contains no pages.")
        writer = PdfWriter(clone_from=reader)
    except PDFMergeError:
        raise
    except DependencyError:
        logger.error("'cryptography' package is missing.")
        raise PDFMergeError("Server decryption support is not configured.", status=500)
    except (PyPdfError, NotImplementedError, ValueError, KeyError, OSError, RecursionError) as exc:
        logger.warning("Unreadable encrypted PDF: %s", exc)
        raise PDFMergeError("The PDF is corrupted or uses an unsupported encryption method.")

    output_path = os.path.join(work_dir, "output.pdf")
    try:
        with open(output_path, "wb") as out:
            writer.write(out)
    except (PyPdfError, ValueError, KeyError, OSError, RecursionError) as exc:
        logger.warning("Writing unlocked PDF failed: %s", exc)
        raise PDFMergeError("The PDF is corrupted and cannot be unlocked.")
    finally:
        writer.close()

    return output_path, "_unlocked.pdf", "application/pdf"


@csrf_exempt
@require_POST
@clerk_authenticated
def unlock_pdf(request):
    return _run_single_pdf_tool(
        request, prefix="unlock_", max_size=UNLOCK_MAX_FILE_SIZE, processor=_process_unlock_pdf
    )


# ======================================================================
# Non-PDF Upload Helpers
# ======================================================================
def _is_zip_header(header):
    return header.startswith(b"PK\x03\x04")


def _is_office_header(header):
    # Supports both ZIP-based Office (.pptx/.docx/.xlsx) and legacy OLE2 (.ppt/.doc/.xls)
    return header.startswith(b"PK\x03\x04") or header.startswith(b"\xd0\xcf\x11\xe0")


def _is_image_header(header):
    return (
        header.startswith(b"\xff\xd8\xff")
        or header.startswith(b"\x89PNG\r\n\x1a\n")
        or header.startswith(b"GIF8")
        or header.startswith(b"BM")
        or header.startswith(b"II*\x00") or header.startswith(b"MM\x00*")
        or (header[:4] == b"RIFF" and header[8:12] == b"WEBP")
    )


def _is_text_header(header):
    return b"\x00" not in header


def _save_typed_upload(upload, work_dir, *, max_size, allowed_exts, magic_check, label, index=0):
    ext = Path(upload.name).suffix.lower()
    if ext not in allowed_exts:
        raise PDFMergeError(
            f"Only {label} files are accepted ({', '.join(sorted(allowed_exts))}).", status=415
        )
    if upload.size == 0:
        raise PDFMergeError(f"'{upload.name}' is empty.")
    if upload.size > max_size:
        raise PDFMergeError(
            f"'{upload.name}' exceeds the {max_size // (1024 * 1024)} MB limit.", status=413
        )

    header = upload.read(512)
    upload.seek(0)
    if not magic_check(header):
        raise PDFMergeError(f"'{upload.name}' is not a valid {label} file.", status=415)

    dest_path = os.path.join(work_dir, f"input_{index:03d}{ext}")
    with open(dest_path, "wb") as out:
        for chunk in upload.chunks():
            out.write(chunk)
    return dest_path


def _run_single_file_tool(request, *, prefix, max_size, allowed_exts, magic_check, label, processor):
    upload = request.FILES.get("file")
    if upload is None:
        return JsonResponse({"error": "No file uploaded. Use the 'file' field."}, status=400)

    work_dir = tempfile.mkdtemp(prefix=prefix, dir=getattr(settings, "PDF_TEMP_DIR", None))
    response = None

    try:
        input_path = _save_typed_upload(
            upload, work_dir, max_size=max_size, allowed_exts=allowed_exts,
            magic_check=magic_check, label=label,
        )

        output_path, extension, content_type = processor(request, input_path, work_dir)

        if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
            raise PDFMergeError("Conversion produced no output. The file may be corrupted.")

        base_name = get_valid_filename(Path(upload.name).stem) or "document"
        response = SelfCleaningFileResponse(
            open(output_path, "rb"),
            as_attachment=True,
            filename=f"{base_name}{extension}",
            content_type=content_type,
            cleanup_dir=work_dir,
        )
        return response

    except PDFMergeError as exc:
        return JsonResponse({"error": exc.message}, status=exc.status)
    except Exception:
        logger.exception("Unexpected error in %s tool", prefix.rstrip("_"))
        return JsonResponse({"error": "Internal error while processing the file."}, status=500)
    finally:
        if response is None:
            shutil.rmtree(work_dir, ignore_errors=True)


# ======================================================================
# TOOL 10: PDF TO EXCEL
# ======================================================================
_NUMERIC_RE = re.compile(r"^-?\d+(\.\d+)?$")


def _excel_cell(value, illegal_re):
    if value is None:
        return None
    text = illegal_re.sub("", str(value)).strip()
    if not text:
        return None
    if _NUMERIC_RE.match(text):
        try:
            return int(text) if "." not in text else float(text)
        except ValueError:
            pass
    return text[:32767]


def _process_pdf_to_excel(request, input_path, work_dir):
    _check_pdf_readable(input_path, EXCEL_MAX_PAGES)
    try:
        import pdfplumber
        from openpyxl import Workbook
        from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    except ImportError:
        logger.error("pdfplumber/openpyxl not installed.")
        raise PDFMergeError("Server is missing PDF to Excel conversion tools.", status=500)

    workbook = Workbook()
    workbook.remove(workbook.active)
    cell_count = 0

    try:
        with pdfplumber.open(input_path) as pdf:
            for page_no, page in enumerate(pdf.pages, start=1):
                tables = page.extract_tables() or []
                wrote_table = False

                for table_no, table in enumerate(tables, start=1):
                    rows = [r for r in table if r and any(c not in (None, "") for c in r)]
                    if not rows:
                        continue
                    sheet = workbook.create_sheet(title=f"P{page_no}_T{table_no}"[:31])
                    for row in rows:
                        sheet.append([_excel_cell(c, ILLEGAL_CHARACTERS_RE) for c in row])
                        cell_count += len(row)
                        if cell_count > EXCEL_MAX_CELLS:
                            raise PDFMergeError(
                                f"The PDF contains more than {EXCEL_MAX_CELLS} cells.", status=413
                            )
                    wrote_table = True

                if not wrote_table:
                    text = (page.extract_text() or "").strip()
                    if text:
                        sheet = workbook.create_sheet(title=f"P{page_no}_Text"[:31])
                        for line in text.splitlines():
                            cell = _excel_cell(line, ILLEGAL_CHARACTERS_RE)
                            if cell is not None:
                                sheet.append([cell])
                                cell_count += 1
    except PDFMergeError:
        raise
    except Exception as exc:
        logger.warning("pdfplumber failed: %s", exc)
        raise PDFMergeError("The PDF is corrupted or its content cannot be read.")

    if not workbook.sheetnames:
        raise PDFMergeError(NO_TEXT_MESSAGE, status=422)

    output_path = os.path.join(work_dir, "output.xlsx")
    workbook.save(output_path)
    return output_path, ".xlsx", EXCEL_MIME


@csrf_exempt
@require_POST
@clerk_authenticated
def pdf_to_excel(request):
    return _run_single_pdf_tool(
        request, prefix="pdf2xlsx_", max_size=EXCEL_MAX_FILE_SIZE, processor=_process_pdf_to_excel
    )


# ======================================================================
# TOOL 11: EXCEL TO PDF & CSV TO PDF
# ======================================================================
def _xlsx_cell_text(value):
    if value is None:
        return ""
    if hasattr(value, "isoformat"):
        text = value.isoformat(sep=" ") if hasattr(value, "hour") else value.isoformat()
    elif isinstance(value, float) and value.is_integer():
        text = str(int(value))
    else:
        text = str(value)
    text = " ".join(text.split())
    if len(text) > XLSX_MAX_CELL_CHARS:
        text = text[: XLSX_MAX_CELL_CHARS - 1] + "…"
    return text


def _register_pdf_font():
    if PDF_FONT_PATH and os.path.exists(PDF_FONT_PATH):
        try:
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.ttfonts import TTFont
            if "SnapPDFSans" not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont("SnapPDFSans", PDF_FONT_PATH))
            return "SnapPDFSans", "SnapPDFSans"
        except Exception as exc:
            logger.warning("Could not register PDF font %s: %s", PDF_FONT_PATH, exc)
    return "Helvetica", "Helvetica-Bold"


def _render_sheets_to_pdf(sheets, output_path, doc_title="Spreadsheet"):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
    )

    font, font_bold = _register_pdf_font()
    page_size = landscape(A4)
    margin = 12 * mm
    avail_width = page_size[0] - 2 * margin
    styles = getSampleStyleSheet()
    title_style = styles["Heading2"]
    title_style.fontName = font_bold

    story = []
    for sheet_no, (title, rows) in enumerate(sheets):
        cols = len(rows[0])
        font_size = 8
        widths = [
            max(6, min(40, max(len(r[c]) for r in rows))) * font_size * 0.55 + 8
            for c in range(cols)
        ]
        total = sum(widths)
        if total > avail_width:
            scale = avail_width / total
            widths = [w * scale for w in widths]
            font_size = max(5, font_size * scale)

        table = Table(rows, colWidths=widths, repeatRows=1)
        table.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), font),
            ("FONTNAME", (0, 0), (-1, 0), font_bold),
            ("FONTSIZE", (0, 0), (-1, -1), font_size),
            ("LEADING", (0, 0), (-1, -1), font_size * 1.2),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EFEFEF")),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#BBBBBB")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 2.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
        ]))

        if sheet_no:
            story.append(PageBreak())
        if title:
            story.append(Paragraph(_xml_escape(title), title_style))
            story.append(Spacer(1, 4 * mm))
        story.append(table)

    try:
        doc = SimpleDocTemplate(
            output_path, pagesize=page_size,
            leftMargin=margin, rightMargin=margin, topMargin=margin, bottomMargin=margin,
            title=doc_title, author="SnapPDF",
        )
        doc.build(story)
    except Exception as exc:
        logger.warning("reportlab build failed: %s", exc)
        raise PDFMergeError("The table data could not be laid out as a PDF.", status=500)


def _process_excel_to_pdf(request, input_path, work_dir):
    try:
        from openpyxl import load_workbook
        from openpyxl.utils.exceptions import InvalidFileException
    except ImportError:
        raise PDFMergeError("Server is missing Excel to PDF conversion tools.", status=500)

    try:
        workbook = load_workbook(input_path, read_only=True, data_only=True)
    except (InvalidFileException, zipfile.BadZipFile, KeyError, ValueError, OSError):
        raise PDFMergeError("The Excel file is corrupted or unreadable.")

    if len(workbook.sheetnames) > XLSX_MAX_SHEETS:
        raise PDFMergeError(
            f"Workbook has {len(workbook.sheetnames)} sheets; the maximum is {XLSX_MAX_SHEETS}.",
            status=413,
        )

    sheets, total_cells = [], 0
    try:
        for sheet in workbook.worksheets:
            rows, used_cols = [], 0
            for raw_row in sheet.iter_rows(values_only=True):
                row = [_xlsx_cell_text(v) for v in raw_row]
                last = max((i + 1 for i, v in enumerate(row) if v), default=0)
                rows.append(row)
                used_cols = max(used_cols, last)
                total_cells += len(row)
                if total_cells > XLSX_MAX_CELLS:
                    raise PDFMergeError(
                        f"Workbook has more than {XLSX_MAX_CELLS} cells; split it first.", status=413
                    )
            while rows and not any(rows[-1]):
                rows.pop()
            rows = [(r + [""] * used_cols)[:used_cols] for r in rows]
            if rows and used_cols:
                sheets.append((sheet.title, rows))
    finally:
        workbook.close()

    if not sheets:
        raise PDFMergeError("The spreadsheet contains no data.", status=422)

    output_path = os.path.join(work_dir, "output.pdf")
    _render_sheets_to_pdf(sheets, output_path, doc_title=Path(input_path).stem)
    return output_path, ".pdf", "application/pdf"


@csrf_exempt
@require_POST
@clerk_authenticated
def excel_to_pdf(request):
    return _run_single_file_tool(
        request, prefix="xlsx2pdf_", max_size=XLSX_MAX_FILE_SIZE,
        allowed_exts=XLSX_EXTENSIONS, magic_check=_is_zip_header, label="Excel",
        processor=_process_excel_to_pdf,
    )


# ======================================================================
# TOOL: CSV TO PDF
# ======================================================================
def _process_csv_to_pdf(request, input_path, work_dir):
    with open(input_path, "rb") as raw_f:
        raw_bytes = raw_f.read()

    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            text = raw_bytes.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise PDFMergeError("Could not decode the CSV file text encoding.")

    sample = text[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel

    reader = csv.reader(io.StringIO(text), dialect)
    rows, used_cols, total_cells = [], 0, 0
    for raw_row in reader:
        row = [_xlsx_cell_text(v) for v in raw_row]
        if not any(row):
            continue
        last = max((i + 1 for i, v in enumerate(row) if v), default=0)
        rows.append(row)
        used_cols = max(used_cols, last)
        total_cells += len(row)
        if total_cells > XLSX_MAX_CELLS:
            raise PDFMergeError(f"CSV exceeds {XLSX_MAX_CELLS} cells limit.", status=413)

    if not rows or used_cols == 0:
        raise PDFMergeError("The CSV file contains no data.", status=422)

    rows = [(r + [""] * used_cols)[:used_cols] for r in rows]
    output_path = os.path.join(work_dir, "output.pdf")
    _render_sheets_to_pdf([("CSV Table", rows)], output_path, doc_title="CSV Export")
    return output_path, ".pdf", "application/pdf"


@csrf_exempt
@require_POST
@clerk_authenticated
def csv_to_pdf(request):
    return _run_single_file_tool(
        request, prefix="csv2pdf_", max_size=CSV_MAX_FILE_SIZE,
        allowed_exts=CSV_EXTENSIONS, magic_check=_is_text_header, label="CSV",
        processor=_process_csv_to_pdf,
    )


# ======================================================================
# TOOL: WORD (.DOCX) TO PDF
# ======================================================================
def _process_word_to_pdf(request, input_path, work_dir):
    output_path = os.path.join(work_dir, "output.pdf")

    for binary in ("libreoffice", "soffice"):
        if shutil.which(binary):
            try:
                subprocess.run(
                    [binary, "--headless", "--convert-to", "pdf", "--outdir", work_dir, input_path],
                    cwd=work_dir, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=90, check=False,
                )
                generated = os.path.join(work_dir, Path(input_path).stem + ".pdf")
                if os.path.exists(generated) and os.path.getsize(generated) > 0:
                    os.replace(generated, output_path)
                    return output_path, ".pdf", "application/pdf"
            except Exception as exc:
                logger.warning("LibreOffice conversion failed, falling back to python-docx: %s", exc)

    try:
        import docx
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.lib.units import mm
        from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    except ImportError:
        raise PDFMergeError("Server is missing Word (.docx) conversion libraries.", status=500)

    try:
        document = docx.Document(input_path)
    except Exception:
        raise PDFMergeError("The Word (.docx) file is corrupted or unreadable.")

    font, font_bold = _register_pdf_font()
    styles = getSampleStyleSheet()
    normal_style = ParagraphStyle(
        "DocNormal", parent=styles["Normal"], fontName=font, fontSize=11, leading=15, spaceAfter=6
    )
    h1_style = ParagraphStyle(
        "DocH1", parent=styles["Heading1"], fontName=font_bold, fontSize=18, leading=22, spaceAfter=10
    )
    h2_style = ParagraphStyle(
        "DocH2", parent=styles["Heading2"], fontName=font_bold, fontSize=14, leading=18, spaceAfter=8
    )

    story = []
    for para in document.paragraphs:
        text = para.text.strip()
        if not text:
            story.append(Spacer(1, 3 * mm))
            continue
        style_name = (para.style.name or "").lower() if para.style else ""
        escaped = _xml_escape(text)
        if "heading 1" in style_name or "title" in style_name:
            story.append(Paragraph(escaped, h1_style))
        elif "heading" in style_name:
            story.append(Paragraph(escaped, h2_style))
        else:
            story.append(Paragraph(escaped, normal_style))

    for tbl in document.tables:
        t_rows = []
        for row in tbl.rows:
            t_rows.append([_xlsx_cell_text(cell.text) for cell in row.cells])
        if t_rows:
            rl_table = Table(t_rows, repeatRows=1)
            rl_table.setStyle(TableStyle([
                ("FONTNAME", (0, 0), (-1, -1), font),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F3F4F6")),
                ("PADDING", (0, 0), (-1, -1), 4),
            ]))
            story.append(Spacer(1, 3 * mm))
            story.append(rl_table)
            story.append(Spacer(1, 4 * mm))

    if not story:
        raise PDFMergeError("The Word document contains no readable text or tables.", status=422)

    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm, topMargin=18 * mm, bottomMargin=18 * mm,
    )
    doc.build(story)
    return output_path, ".pdf", "application/pdf"


@csrf_exempt
@require_POST
@clerk_authenticated
def word_to_pdf(request):
    return _run_single_file_tool(
        request, prefix="docx2pdf_", max_size=DOCX_MAX_FILE_SIZE,
        allowed_exts=DOCX_EXTENSIONS, magic_check=_is_zip_header, label="Word (.docx)",
        processor=_process_word_to_pdf,
    )


# ======================================================================
# TOOL: HTML TO PDF
# ======================================================================
class _SimpleHTMLParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.blocks = []
        self._curr_tag = "p"
        self._buf = []
        self._skip = False

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag in ("script", "style", "head", "noscript"):
            self._skip = True
        elif tag in ("h1", "h2", "h3", "p", "li", "div", "tr", "br"):
            self._flush()
            self._curr_tag = tag

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in ("script", "style", "head", "noscript"):
            self._skip = False
        elif tag in ("h1", "h2", "h3", "p", "li", "div", "tr"):
            self._flush()
            self._curr_tag = "p"

    def handle_data(self, data):
        if not self._skip:
            self._buf.append(data)

    def _flush(self):
        text = " ".join("".join(self._buf).split()).strip()
        self._buf.clear()
        if text:
            self.blocks.append((self._curr_tag, text))


def _render_html_string_to_pdf(html_content, output_path):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import Paragraph, SimpleDocTemplate

    parser = _SimpleHTMLParser()
    parser.feed(html_content)
    parser._flush()

    if not parser.blocks:
        raise PDFMergeError("No readable text content found in the HTML.", status=422)

    font, font_bold = _register_pdf_font()
    styles = getSampleStyleSheet()
    body_style = ParagraphStyle("HTMLBody", parent=styles["Normal"], fontName=font, fontSize=10.5, leading=15, spaceAfter=6)
    h1_style = ParagraphStyle("HTMLH1", parent=styles["Heading1"], fontName=font_bold, fontSize=18, leading=22, spaceAfter=10)
    h2_style = ParagraphStyle("HTMLH2", parent=styles["Heading2"], fontName=font_bold, fontSize=14, leading=18, spaceAfter=8)
    li_style = ParagraphStyle("HTMLLi", parent=body_style, leftIndent=14, spaceAfter=4)

    story = []
    for tag, text in parser.blocks:
        esc = _xml_escape(text)
        if tag == "h1":
            story.append(Paragraph(esc, h1_style))
        elif tag in ("h2", "h3"):
            story.append(Paragraph(esc, h2_style))
        elif tag == "li":
            story.append(Paragraph(f"• {esc}", li_style))
        else:
            story.append(Paragraph(esc, body_style))

    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm, topMargin=18 * mm, bottomMargin=18 * mm,
    )
    doc.build(story)


@csrf_exempt
@require_POST
@clerk_authenticated
def html_to_pdf(request):
    upload = request.FILES.get("file")
    raw_html = (request.POST.get("html_code") or "").strip()

    if upload is None and not raw_html:
        return JsonResponse({"error": "Upload an .html file or paste HTML code."}, status=400)

    work_dir = tempfile.mkdtemp(prefix="html2pdf_", dir=getattr(settings, "PDF_TEMP_DIR", None))
    response = None

    try:
        if upload is not None:
            input_path = _save_typed_upload(
                upload, work_dir, max_size=HTML_MAX_FILE_SIZE,
                allowed_exts=HTML_EXTENSIONS, magic_check=_is_text_header, label="HTML",
            )
            with open(input_path, "rb") as f:
                html_str = f.read().decode("utf-8", errors="replace")
            base_name = get_valid_filename(Path(upload.name).stem) or "webpage"
        else:
            html_str = raw_html
            base_name = "webpage"

        output_path = os.path.join(work_dir, "output.pdf")
        _render_html_string_to_pdf(html_str, output_path)

        response = SelfCleaningFileResponse(
            open(output_path, "rb"),
            as_attachment=True,
            filename=f"{base_name}.pdf",
            content_type="application/pdf",
            cleanup_dir=work_dir,
        )
        return response
    except PDFMergeError as exc:
        return JsonResponse({"error": exc.message}, status=exc.status)
    except Exception:
        logger.exception("Unexpected error in HTML to PDF")
        return JsonResponse({"error": "Internal error while converting HTML to PDF."}, status=500)
    finally:
        if response is None:
            shutil.rmtree(work_dir, ignore_errors=True)


# ======================================================================
# TOOL 12: MULTI-IMAGE / JPG TO PDF
# ======================================================================
@csrf_exempt
@require_POST
@clerk_authenticated
def image_to_pdf(request):
    uploads = request.FILES.getlist("files") or request.FILES.getlist("file")
    if not uploads:
        return JsonResponse({"error": "No images uploaded."}, status=400)
    if len(uploads) > IMAGE_MAX_FILES:
        return JsonResponse({"error": f"Maximum {IMAGE_MAX_FILES} images allowed at once."}, status=400)

    page_size_opt = (request.POST.get("page_size") or "fit").strip().lower()
    orientation_opt = (request.POST.get("orientation") or "auto").strip().lower()
    margin_opt = (request.POST.get("margin") or "none").strip().lower()

    margin_pts = {"none": 0, "small": 24, "big": 54}.get(margin_opt, 0)

    work_dir = tempfile.mkdtemp(prefix="img2pdf_", dir=getattr(settings, "PDF_TEMP_DIR", None))
    response = None

    try:
        from PIL import Image, ImageOps
    except ImportError:
        shutil.rmtree(work_dir, ignore_errors=True)
        return JsonResponse({"error": "Server is missing image conversion tools."}, status=500)

    try:
        saved_paths = []
        for idx, upload in enumerate(uploads):
            p = _save_typed_upload(
                upload, work_dir, max_size=IMAGE_MAX_FILE_SIZE,
                allowed_exts=IMAGE_EXTENSIONS, magic_check=_is_image_header, label="image", index=idx,
            )
            saved_paths.append(p)

        processed_pages = []
        for p in saved_paths:
            with Image.open(p) as img:
                w, h = img.size
                if w * h > IMAGE_MAX_PIXELS:
                    raise PDFMergeError(f"Image is {w}x{h}px; maximum is {IMAGE_MAX_PIXELS:,} pixels.", status=413)
                try:
                    img = ImageOps.exif_transpose(img) or img
                except Exception:
                    pass

                has_alpha = img.mode in ("RGBA", "LA", "PA") or (img.mode == "P" and "transparency" in img.info)
                if has_alpha:
                    rgba = img.convert("RGBA")
                    flat = Image.new("RGB", rgba.size, (255, 255, 255))
                    flat.paste(rgba, mask=rgba.getchannel("A"))
                    rgb_img = flat
                elif img.mode == "RGB":
                    rgb_img = img.copy()
                else:
                    rgb_img = img.convert("RGB")

                if page_size_opt == "a4":
                    a4_w, a4_h = 1240, 1754
                    if orientation_opt == "landscape" or (orientation_opt == "auto" and rgb_img.width > rgb_img.height):
                        a4_w, a4_h = a4_h, a4_w
                    pad = int(margin_pts * (150.0 / 72.0))
                    avail_w = max(100, a4_w - 2 * pad)
                    avail_h = max(100, a4_h - 2 * pad)
                    rgb_img.thumbnail((avail_w, avail_h), Image.Resampling.LANCZOS)
                    canvas = Image.new("RGB", (a4_w, a4_h), (255, 255, 255))
                    canvas.paste(rgb_img, ((a4_w - rgb_img.width) // 2, (a4_h - rgb_img.height) // 2))
                    rgb_img.close()
                    processed_pages.append(canvas)
                else:
                    if margin_pts > 0:
                        pad = int(margin_pts * 2)
                        canvas = Image.new("RGB", (rgb_img.width + 2 * pad, rgb_img.height + 2 * pad), (255, 255, 255))
                        canvas.paste(rgb_img, (pad, pad))
                        rgb_img.close()
                        processed_pages.append(canvas)
                    else:
                        processed_pages.append(rgb_img)

        output_path = os.path.join(work_dir, "output.pdf")
        first, rest = processed_pages[0], processed_pages[1:]
        first.save(output_path, "PDF", resolution=150.0, save_all=True, append_images=rest)
        for im in processed_pages:
            im.close()

        base_name = get_valid_filename(Path(uploads[0].name).stem) or "images"
        suffix = "_merged.pdf" if len(uploads) > 1 else ".pdf"
        response = SelfCleaningFileResponse(
            open(output_path, "rb"),
            as_attachment=True,
            filename=f"{base_name}{suffix}",
            content_type="application/pdf",
            cleanup_dir=work_dir,
        )
        return response

    except PDFMergeError as exc:
        return JsonResponse({"error": exc.message}, status=exc.status)
    except Exception:
        logger.exception("Unexpected error in image_to_pdf")
        return JsonResponse({"error": "Internal error while converting images to PDF."}, status=500)
    finally:
        if response is None:
            shutil.rmtree(work_dir, ignore_errors=True)


# ======================================================================
# TOOL 13: CROP PDF
# ======================================================================
_CROP_SIDES = ("left", "top", "right", "bottom")
_CROP_UNITS = {"pt": 1.0, "mm": 72 / 25.4, "cm": 72 / 2.54, "in": 72.0}


def _parse_crop_margins(post):
    unit = (post.get("unit") or "pt").strip().lower()
    if unit not in _CROP_UNITS:
        raise PDFMergeError(f"'unit' must be one of: {', '.join(_CROP_UNITS)}.")
    factor = _CROP_UNITS[unit]

    missing = [side for side in _CROP_SIDES if post.get(side, "").strip() == ""]
    if missing:
        raise PDFMergeError(
            f"Missing crop value(s): {', '.join(missing)}. Provide left, top, right and bottom."
        )

    margins = {}
    for side in _CROP_SIDES:
        try:
            value = float(post[side])
        except (TypeError, ValueError):
            raise PDFMergeError(f"'{side}' must be a number.")
        if not math.isfinite(value) or value < 0:
            raise PDFMergeError(f"'{side}' must be a non-negative number.")
        margins[side] = value * factor

    if not any(margins.values()):
        raise PDFMergeError("Drag the crop box or set at least one margin greater than zero.")
    return margins


def _margins_for_rotation(m, rotation):
    l, t, r, b = m["left"], m["top"], m["right"], m["bottom"]
    if rotation == 90:
        return t, l, b, r
    if rotation == 180:
        return r, t, l, b
    if rotation == 270:
        return b, r, t, l
    return l, b, r, t


def _process_crop_pdf(request, input_path, work_dir):
    margins = _parse_crop_margins(request.POST)
    raw_pages = (request.POST.get("pages") or "").strip()
    ranges = _parse_page_ranges(raw_pages) if raw_pages else None

    reader = _check_pdf_readable(input_path, CROP_MAX_PAGES)
    total_pages = len(reader.pages)
    indexes = _expand_page_ranges(ranges, total_pages) if ranges else range(total_pages)

    output_path = os.path.join(work_dir, "output.pdf")
    try:
        writer = PdfWriter(clone_from=reader)
    except (PyPdfError, ValueError, KeyError, OSError, RecursionError) as exc:
        logger.warning("Cloning PDF for crop failed: %s", exc)
        raise PDFMergeError("The PDF is corrupted or unreadable.")

    try:
        for index in indexes:
            page = writer.pages[index]
            try:
                rotation = int(page.rotation) % 360
            except Exception:
                rotation = 0

            ml, mb, mr, mt = _margins_for_rotation(margins, rotation)
            box = page.cropbox
            left = float(box.left) + ml
            bottom = float(box.bottom) + mb
            right = float(box.right) - mr
            top = float(box.top) - mt

            if right - left < 1 or top - bottom < 1:
                raise PDFMergeError(
                    f"Crop margins remove all of page {index + 1} "
                    f"({float(box.width):.0f} x {float(box.height):.0f} pt)."
                )

            new_box = RectangleObject((left, bottom, right, top))
            page.mediabox = new_box
            page.cropbox = new_box
            if "/TrimBox" in page:
                page.trimbox = new_box
            if "/BleedBox" in page:
                page.bleedbox = new_box
            if "/ArtBox" in page:
                page.artbox = new_box

        with open(output_path, "wb") as out:
            writer.write(out)
    except PDFMergeError:
        raise
    except (PyPdfError, ValueError, KeyError, IndexError, OSError, RecursionError) as exc:
        logger.warning("Crop failed: %s", exc)
        raise PDFMergeError("Some pages are corrupted and cannot be cropped.")
    finally:
        writer.close()

    return output_path, "_cropped.pdf", "application/pdf"


@csrf_exempt
@require_POST
@clerk_authenticated
def crop_pdf(request):
    return _run_single_pdf_tool(
        request, prefix="crop_", max_size=CROP_MAX_FILE_SIZE, processor=_process_crop_pdf
    )


# ======================================================================
# TOOL 14: EDIT PDF (With Sans / Serif / Mono & Bold / Italic Auto-Match)
# ======================================================================
EDIT_MAX_FILE_SIZE = getattr(settings, "PDF_EDIT_MAX_FILE_SIZE", 150 * 1024 * 1024)
EDIT_MAX_PAGES = getattr(settings, "PDF_EDIT_MAX_PAGES", 2000)
EDIT_MAX_ELEMENTS = getattr(settings, "PDF_EDIT_MAX_ELEMENTS", 500)
EDIT_MAX_TEXT_CHARS = getattr(settings, "PDF_EDIT_MAX_TEXT_CHARS", 5000)
EDIT_MAX_JSON_BYTES = getattr(settings, "PDF_EDIT_MAX_JSON_BYTES", 5 * 1024 * 1024)
EDIT_FONT_PATH = getattr(settings, "PDF_FONT_PATH", None)
EDIT_MIN_FONT_SIZE, EDIT_MAX_FONT_SIZE = 4.0, 200.0
EDIT_LINE_HEIGHT = 1.2

_HEX_COLOR_RE = re.compile(r"^#?([0-9a-fA-F]{6})$")
_CONTROL_CHARS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")


def _import_fitz():
    try:
        import pymupdf as fitz
        return fitz
    except ImportError:
        try:
            import fitz
            return fitz
        except ImportError:
            logger.error("PyMuPDF is not installed.")
            raise PDFMergeError("Server is missing PDF editing tools.", status=500)


def _edit_font():
    if EDIT_FONT_PATH and os.path.exists(EDIT_FONT_PATH):
        return "SnapPDFSans", EDIT_FONT_PATH
    return "helv", None


def _hex_to_rgb(value, label):
    match = _HEX_COLOR_RE.match(str(value).strip())
    if not match:
        raise PDFMergeError(f"{label}: invalid color '{value}'. Use #RRGGBB.")
    h = match.group(1)
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def _edit_number(obj, key, label, *, minimum=None, maximum=None, required=True):
    value = obj.get(key)
    if value is None:
        if required:
            raise PDFMergeError(f"{label}: '{key}' is required.")
        return None
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise PDFMergeError(f"{label}: '{key}' must be a number.")
    value = float(value)
    if not math.isfinite(value):
        raise PDFMergeError(f"{label}: '{key}' must be a finite number.")
    if minimum is not None and value < minimum:
        raise PDFMergeError(f"{label}: '{key}' must be at least {minimum}.")
    if maximum is not None and value > maximum:
        raise PDFMergeError(f"{label}: '{key}' must be at most {maximum}.")
    return value


def _parse_edit_payload(raw):
    if not raw or not raw.strip():
        raise PDFMergeError("'edits' JSON is required.")
    if len(raw) > EDIT_MAX_JSON_BYTES:
        raise PDFMergeError("'edits' payload is too large.", status=413)
    try:
        data = json.loads(raw)
    except ValueError:
        raise PDFMergeError("'edits' is not valid JSON.")
    if not isinstance(data, dict) or not isinstance(data.get("elements"), list):
        raise PDFMergeError("'edits' must be an object with an 'elements' list.")

    raw_elements = data["elements"]
    if not raw_elements:
        raise PDFMergeError("No edits to apply. Add text or a whiteout box first.")
    if len(raw_elements) > EDIT_MAX_ELEMENTS:
        raise PDFMergeError(f"Too many edits; the maximum is {EDIT_MAX_ELEMENTS}.", status=413)

    elements = []
    for index, el in enumerate(raw_elements, start=1):
        label = f"Element {index}"
        if not isinstance(el, dict):
            raise PDFMergeError(f"{label}: must be an object.")
        kind = str(el.get("type", "")).lower()
        if kind not in ("text", "whiteout"):
            raise PDFMergeError(f"{label}: 'type' must be 'text' or 'whiteout'.")

        item = {
            "type": kind,
            "page": int(_edit_number(el, "page", label, minimum=1, maximum=EDIT_MAX_PAGES)),
            "x": _edit_number(el, "x", label, minimum=-100_000, maximum=100_000),
            "y": _edit_number(el, "y", label, minimum=-100_000, maximum=100_000),
            "w": _edit_number(el, "w", label, minimum=0, maximum=100_000),
            "h": _edit_number(el, "h", label, minimum=0, maximum=100_000),
            "page_w": _edit_number(el, "pageWidth", label, minimum=1, maximum=100_000, required=False),
            "page_h": _edit_number(el, "pageHeight", label, minimum=1, maximum=100_000, required=False),
        }

        if kind == "text":
            text = el.get("text", "")
            if not isinstance(text, str):
                raise PDFMergeError(f"{label}: 'text' must be a string.")
            text = _CONTROL_CHARS_RE.sub("", text.replace("\r\n", "\n").replace("\r", "\n"))
            if not text.strip():
                raise PDFMergeError(f"{label}: 'text' is empty.")
            if len(text) > EDIT_MAX_TEXT_CHARS:
                raise PDFMergeError(
                    f"{label}: text exceeds {EDIT_MAX_TEXT_CHARS} characters.", status=413
                )
            font_size = _edit_number(
                el, "fontSize", label,
                minimum=EDIT_MIN_FONT_SIZE, maximum=EDIT_MAX_FONT_SIZE, required=False,
            )
            align = str(el.get("align", "left")).lower()
            if align not in ("left", "center", "right"):
                raise PDFMergeError(f"{label}: 'align' must be left, center or right.")

            font_family = str(el.get("fontFamily", "sans")).lower()
            if font_family not in ("sans", "serif", "mono"):
                font_family = "sans"

            item.update(
                text=text,
                font_size=font_size if font_size is not None else 14.0,
                font_family=font_family,
                color=_hex_to_rgb(el.get("color", "#000000"), label),
                bold=bool(el.get("bold", False)),
                italic=bool(el.get("italic", False)),
                align=align,
            )
        else:
            if item["w"] <= 0 or item["h"] <= 0:
                raise PDFMergeError(f"{label}: a whiteout box needs a positive width and height.")
            item["fill"] = _hex_to_rgb(el.get("fill", "#ffffff"), label)

        elements.append(item)

    return elements, bool(data.get("redact", False))


def _element_rect(fitz, el, page_rect, label):
    sx = page_rect.width / el["page_w"] if el["page_w"] else 1.0
    sy = page_rect.height / el["page_h"] if el["page_h"] else 1.0
    box = fitz.Rect(
        page_rect.x0 + el["x"] * sx,
        page_rect.y0 + el["y"] * sy,
        page_rect.x0 + (el["x"] + el["w"]) * sx,
        page_rect.y0 + (el["y"] + el["h"]) * sy,
    )
    clipped = box & page_rect
    if clipped.is_empty or clipped.width < 0.5 or clipped.height < 0.5:
        raise PDFMergeError(f"{label} lies outside page {el['page']}.")
    return clipped


def _draw_whiteout(fitz, page, box, fill):
    shape = page.new_shape()
    shape.draw_rect(box * page.derotation_matrix)
    shape.finish(color=None, fill=fill, width=0)
    shape.commit()


def _redact_boxes(fitz, page, boxes):
    derot = page.derotation_matrix
    for box, fill in boxes:
        page.add_redact_annot(box * derot, fill=fill)
    page.apply_redactions(images=getattr(fitz, "PDF_REDACT_IMAGE_PIXELS", 2))


def _textbox(page, rect, text, font_size, kwargs):
    try:
        return page.insert_textbox(
            rect, text, fontsize=font_size, lineheight=EDIT_LINE_HEIGHT, **kwargs
        )
    except TypeError:
        return page.insert_textbox(rect, text, fontsize=font_size, **kwargs)


_PYMUPDF_FONTS = {
    ("sans", False, False): "helv",
    ("sans", True, False): "hebo",
    ("sans", False, True): "heit",
    ("sans", True, True): "hebi",
    ("serif", False, False): "tiro",
    ("serif", True, False): "tibo",
    ("serif", False, True): "tiit",
    ("serif", True, True): "tibi",
    ("mono", False, False): "cour",
    ("mono", True, False): "cobo",
    ("mono", False, True): "coit",
    ("mono", True, True): "cobi",
}


def _insert_text(fitz, page, box, el, fontname, fontfile, label):
    align = {
        "left": fitz.TEXT_ALIGN_LEFT,
        "center": fitz.TEXT_ALIGN_CENTER,
        "right": fitz.TEXT_ALIGN_RIGHT,
    }[el["align"]]

    kwargs = dict(
        fontname=fontname,
        fontfile=fontfile,
        color=el["color"],
        align=align,
        rotate=page.rotation,
    )
    if fontfile is None:
        key = (el.get("font_family", "sans"), bool(el.get("bold")), bool(el.get("italic")))
        kwargs["fontname"] = _PYMUPDF_FONTS.get(key, "helv")
    elif el["bold"]:
        kwargs["render_mode"] = 2
        kwargs["border_width"] = max(0.2, el["font_size"] * 0.025)

    page_rect = page.rect
    derot = page.derotation_matrix
    font_size = el["font_size"]

    for _ in range(12):
        rc = _textbox(page, box * derot, el["text"], font_size, kwargs)
        if rc >= 0:
            return
        grown = fitz.Rect(box.x0, box.y0, box.x1, min(page_rect.y1, box.y1 - rc + 1))
        if grown.height > box.height + 0.5:
            box = grown
            continue
        font_size *= 0.9
        if font_size < EDIT_MIN_FONT_SIZE:
            break

    raise PDFMergeError(
        f"{label}: the text does not fit on page {el['page']}. "
        "Enlarge the box or shorten the text."
    )


def _process_edit_pdf(request, input_path, work_dir):
    elements, redact = _parse_edit_payload(request.POST.get("edits"))
    _check_pdf_readable(input_path, EDIT_MAX_PAGES)

    fitz = _import_fitz()
    fontname, fontfile = _edit_font()

    try:
        doc = fitz.open(input_path)
    except Exception as exc:
        logger.warning("PyMuPDF could not open PDF: %s", exc)
        raise PDFMergeError("The PDF is corrupted or unreadable.")

    try:
        if doc.needs_pass:
            raise PDFMergeError("The PDF is password protected. Unlock it first.")
        page_count = doc.page_count

        by_page = {}
        for index, el in enumerate(elements, start=1):
            if el["page"] > page_count:
                raise PDFMergeError(
                    f"Element {index} targets page {el['page']}, "
                    f"but the PDF has {page_count} page(s)."
                )
            el["label"] = f"Element {index}"
            by_page.setdefault(el["page"], []).append(el)

        for page_no, items in sorted(by_page.items()):
            page = doc[page_no - 1]
            page_rect = page.rect
            whiteouts = [e for e in items if e["type"] == "whiteout"]
            texts = [e for e in items if e["type"] == "text"]

            if whiteouts:
                boxes = [
                    (_element_rect(fitz, e, page_rect, e["label"]), e["fill"]) for e in whiteouts
                ]
                if redact:
                    _redact_boxes(fitz, page, boxes)
                else:
                    for box, fill in boxes:
                        _draw_whiteout(fitz, page, box, fill)

            for e in texts:
                box = _element_rect(fitz, e, page_rect, e["label"])
                _insert_text(fitz, page, box, e, fontname, fontfile, e["label"])

        output_path = os.path.join(work_dir, "output.pdf")
        doc.save(output_path, garbage=3, deflate=True)

    except PDFMergeError:
        raise
    except Exception as exc:
        logger.warning("Edit PDF failed: %s", exc)
        raise PDFMergeError("The PDF could not be edited. It may be corrupted.")
    finally:
        doc.close()

    return output_path, "_edited.pdf", "application/pdf"


@csrf_exempt
@require_POST
@clerk_authenticated
def edit_pdf(request):
    return _run_single_pdf_tool(
        request, prefix="edit_", max_size=EDIT_MAX_FILE_SIZE, processor=_process_edit_pdf
    )


# ======================================================================
# NEW TOOL 15: OCR PDF (Tesseract OCR -> Searchable PDF or Extracted Text)
# ======================================================================
_OCR_LANG_WHITELIST = {"eng", "urd", "ara", "hin", "fra", "deu", "spa", "eng+urd", "eng+ara"}


def _process_ocr_pdf(request, input_path, work_dir):
    """
    POST /ocr-pdf/
    Form-data:
      file = <pdf>
      lang = eng | urd | ara | hin | fra | deu | spa | eng+urd (default: eng)
      output_format = pdf | txt (default: pdf -> searchable PDF with invisible text layer)
      pages = optional page range e.g. "1-10"
    """
    lang = (request.POST.get("lang") or "eng").strip().lower()
    # Map 2-letter codes if sent from frontend
    lang_map = {"en": "eng", "ur": "urd", "ar": "ara", "hi": "hin", "fr": "fra", "de": "deu", "es": "spa"}
    lang = lang_map.get(lang, lang)
    if lang not in _OCR_LANG_WHITELIST:
        lang = "eng"

    out_fmt = (request.POST.get("output_format") or "pdf").strip().lower()
    raw_pages = (request.POST.get("pages") or "").strip()

    fitz = _import_fitz()
    tesseract_bin = shutil.which("tesseract")
    if not tesseract_bin:
        raise PDFMergeError(
            "Tesseract OCR is not installed on the server. Run: sudo apt install tesseract-ocr",
            status=500,
        )

    try:
        doc = fitz.open(input_path)
    except Exception:
        raise PDFMergeError("The PDF is corrupted or unreadable.")

    try:
        if doc.needs_pass:
            raise PDFMergeError("The PDF is password protected. Unlock it first.")
        total_pages = doc.page_count
        if total_pages < 1:
            raise PDFMergeError("The PDF contains no pages.")
        if total_pages > OCR_MAX_PAGES and not raw_pages:
            raise PDFMergeError(
                f"PDF has {total_pages} pages; OCR maximum is {OCR_MAX_PAGES} pages at once. Specify a page range.",
                status=413,
            )

        target_indexes = (
            _expand_page_ranges(_parse_page_ranges(raw_pages), total_pages)
            if raw_pages else list(range(total_pages))
        )
        if len(target_indexes) > OCR_MAX_PAGES:
            raise PDFMergeError(f"Please select at most {OCR_MAX_PAGES} pages for OCR.", status=413)

        from PIL import Image
        try:
            import pytesseract
            has_pytesseract = True
        except ImportError:
            has_pytesseract = False

        mat = fitz.Matrix(200 / 72.0, 200 / 72.0)  # 200 DPI for fast, accurate OCR

        # Mode A: Plain text output (.txt)
        if out_fmt == "txt":
            extracted_pages = []
            for idx in target_indexes:
                pix = doc[idx].get_pixmap(matrix=mat, alpha=False)
                img_path = os.path.join(work_dir, f"ocr_p{idx}.png")
                pix.save(img_path)
                if has_pytesseract:
                    try:
                        txt = pytesseract.image_to_string(Image.open(img_path), lang=lang, timeout=60)
                    except Exception:
                        txt = pytesseract.image_to_string(Image.open(img_path), lang="eng", timeout=60)
                else:
                    res = subprocess.run(
                        [tesseract_bin, img_path, "stdout", "-l", lang],
                        stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60, check=False,
                    )
                    if res.returncode != 0 and lang != "eng":
                        res = subprocess.run(
                            [tesseract_bin, img_path, "stdout", "-l", "eng"],
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60, check=False,
                        )
                    txt = res.stdout.decode("utf-8", errors="replace")
                os.remove(img_path)
                extracted_pages.append(f"--- Page {idx + 1} ---\n{txt.strip()}")

            output_txt = os.path.join(work_dir, "output.txt")
            with open(output_txt, "w", encoding="utf-8") as f:
                f.write("\n\n".join(extracted_pages))
            return output_txt, "_ocr.txt", "text/plain; charset=utf-8"

        # Mode B: Searchable PDF output (.pdf)
        out_doc = fitz.open()
        for idx in target_indexes:
            pix = doc[idx].get_pixmap(matrix=mat, alpha=False)
            # 1. Try PyMuPDF native pdfocr_tobytes
            if hasattr(pix, "pdfocr_tobytes"):
                try:
                    pdf_bytes = pix.pdfocr_tobytes(language=lang)
                except Exception:
                    pdf_bytes = pix.pdfocr_tobytes(language="eng")
                ocr_page_doc = fitz.open("pdf", pdf_bytes)
                out_doc.insert_pdf(ocr_page_doc)
                ocr_page_doc.close()
            else:
                # 2. Fallback to pytesseract or tesseract CLI PDF output
                img_path = os.path.join(work_dir, f"ocr_p{idx}.png")
                pix.save(img_path)
                if has_pytesseract:
                    try:
                        pdf_bytes = pytesseract.image_to_pdf_or_hocr(img_path, lang=lang, extension="pdf", timeout=60)
                    except Exception:
                        pdf_bytes = pytesseract.image_to_pdf_or_hocr(img_path, lang="eng", extension="pdf", timeout=60)
                    ocr_page_doc = fitz.open("pdf", pdf_bytes)
                    out_doc.insert_pdf(ocr_page_doc)
                    ocr_page_doc.close()
                else:
                    base_out = os.path.join(work_dir, f"ocr_out_{idx}")
                    res = subprocess.run(
                        [tesseract_bin, img_path, base_out, "-l", lang, "pdf"],
                        stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60, check=False,
                    )
                    if res.returncode != 0 and lang != "eng":
                        subprocess.run(
                            [tesseract_bin, img_path, base_out, "-l", "eng", "pdf"],
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60, check=False,
                        )
                    page_pdf_path = f"{base_out}.pdf"
                    if os.path.exists(page_pdf_path):
                        ocr_page_doc = fitz.open(page_pdf_path)
                        out_doc.insert_pdf(ocr_page_doc)
                        ocr_page_doc.close()
                        os.remove(page_pdf_path)
                if os.path.exists(img_path):
                    os.remove(img_path)

        output_path = os.path.join(work_dir, "output.pdf")
        out_doc.save(output_path, garbage=3, deflate=True)
        out_doc.close()
        return output_path, "_ocr.pdf", "application/pdf"
    finally:
        doc.close()


@csrf_exempt
@require_POST
@clerk_authenticated
def ocr_pdf(request):
    return _run_single_pdf_tool(
        request, prefix="ocr_", max_size=OCR_MAX_FILE_SIZE, processor=_process_ocr_pdf
    )


# ======================================================================
# NEW TOOL 16: WATERMARK PDF (Image Upload + Adjustable Transparency / Text)
# ======================================================================
def _prepare_transparent_watermark_png(img_path, out_png_path, opacity):
    """Apply user-selected opacity (0.05 - 1.0) to the uploaded watermark/background image."""
    from PIL import Image, ImageOps

    with Image.open(img_path) as im:
        w, h = im.size
        if w * h > IMAGE_MAX_PIXELS:
            raise PDFMergeError("Watermark image resolution is too large.", status=413)
        try:
            im = ImageOps.exif_transpose(im) or im
        except Exception:
            pass
        rgba = im.convert("RGBA")
        if w > 2400 or h > 2400:
            rgba.thumbnail((2400, 2400), Image.Resampling.LANCZOS)

        alpha = rgba.getchannel("A")
        alpha = alpha.point(lambda p: int(round(p * opacity)))
        rgba.putalpha(alpha)
        rgba.save(out_png_path, "PNG")
        return rgba.size


def _compute_placement_rect(fitz, page_rect, item_w, item_h, position, margin=28):
    pw, ph = page_rect.width, page_rect.height
    x0, y0 = page_rect.x0, page_rect.y0

    if position == "top-left":
        x, y = x0 + margin, y0 + margin
    elif position == "top-right":
        x, y = x0 + pw - item_w - margin, y0 + margin
    elif position == "bottom-left":
        x, y = x0 + margin, y0 + ph - item_h - margin
    elif position == "bottom-right":
        x, y = x0 + pw - item_w - margin, y0 + ph - item_h - margin
    elif position == "top-center":
        x, y = x0 + (pw - item_w) / 2, y0 + margin
    elif position == "bottom-center":
        x, y = x0 + (pw - item_w) / 2, y0 + ph - item_h - margin
    else:  # center
        x, y = x0 + (pw - item_w) / 2, y0 + (ph - item_h) / 2

    return fitz.Rect(x, y, x + item_w, y + item_h)


def _process_watermark_pdf(request, input_path, work_dir):
    """
    POST /watermark-pdf/
    Form-data:
      file = <pdf>
      watermark_image = <image file> (optional if text is provided)
      text = <string> (optional fallback if no image is uploaded)
      opacity = 0.05..1.0 OR 5..100 (default 0.30 / 30%)
      scale = 0.1..1.0 OR 10..100 (% of page width, default 50%)
      position = center | stretch | tile | bottom-right | bottom-left | top-right | top-left (default center)
      layer = over | under (default over)
      pages = optional page range e.g. "1,3,5-7"
    """
    wm_upload = request.FILES.get("watermark_image") or request.FILES.get("image")
    wm_text = (request.POST.get("text") or "").strip()

    if wm_upload is None and not wm_text:
        raise PDFMergeError("Please upload a watermark/background image or enter watermark text.")

    # Parse opacity (supports both 0.35 and 35)
    raw_opacity = (request.POST.get("opacity") or "0.35").strip()
    try:
        opacity = float(raw_opacity)
        if opacity > 1.0:
            opacity = opacity / 100.0
        opacity = max(0.05, min(1.0, opacity))
    except ValueError:
        raise PDFMergeError("'opacity' must be a number between 5 and 100.")

    # Parse scale (supports 0.5 or 50)
    raw_scale = (request.POST.get("scale") or "50").strip()
    try:
        scale_pct = float(raw_scale)
        if scale_pct > 1.0:
            scale_pct = scale_pct / 100.0
        scale_pct = max(0.10, min(1.0, scale_pct))
    except ValueError:
        scale_pct = 0.50

    position = (request.POST.get("position") or "center").strip().lower()
    layer = (request.POST.get("layer") or "over").strip().lower()
    overlay = layer != "under"
    raw_pages = (request.POST.get("pages") or "").strip()

    fitz = _import_fitz()
    _check_pdf_readable(input_path, SPLIT_MAX_PAGES)

    wm_png_path = None
    wm_size = (200, 200)
    if wm_upload is not None:
        raw_img_path = _save_typed_upload(
            wm_upload, work_dir, max_size=IMAGE_MAX_FILE_SIZE,
            allowed_exts=IMAGE_EXTENSIONS, magic_check=_is_image_header, label="watermark image", index=1,
        )
        wm_png_path = os.path.join(work_dir, "watermark_rgba.png")
        wm_size = _prepare_transparent_watermark_png(raw_img_path, wm_png_path, opacity)

    doc = fitz.open(input_path)
    try:
        total_pages = doc.page_count
        target_indexes = (
            _expand_page_ranges(_parse_page_ranges(raw_pages), total_pages)
            if raw_pages else list(range(total_pages))
        )

        for idx in target_indexes:
            page = doc[idx]
            rect = page.rect

            if wm_png_path:
                img_w, img_h = wm_size
                aspect = img_h / max(1.0, float(img_w))

                if position in ("stretch", "full", "background"):
                    page.insert_image(rect, filename=wm_png_path, keep_proportion=False, overlay=overlay)
                elif position == "tile":
                    target_w = max(80.0, rect.width * min(0.35, scale_pct))
                    target_h = target_w * aspect
                    step_x = target_w * 1.35
                    step_y = target_h * 1.35
                    y = rect.y0 + 20
                    while y < rect.y1:
                        x = rect.x0 + 20
                        while x < rect.x1:
                            cell = fitz.Rect(x, y, x + target_w, y + target_h)
                            page.insert_image(cell, filename=wm_png_path, keep_proportion=True, overlay=overlay)
                            x += step_x
                        y += step_y
                else:
                    target_w = rect.width * scale_pct
                    target_h = target_w * aspect
                    if target_h > rect.height * scale_pct:
                        target_h = rect.height * scale_pct
                        target_w = target_h / aspect
                    target_rect = _compute_placement_rect(fitz, rect, target_w, target_h, position)
                    page.insert_image(target_rect, filename=wm_png_path, keep_proportion=True, overlay=overlay)
            else:
                # Text watermark fallback rendered onto a transparent PNG so opacity & rotation work identically
                from PIL import Image, ImageDraw, ImageFont
                txt_img = Image.new("RGBA", (900, 300), (255, 255, 255, 0))
                draw = ImageDraw.Draw(txt_img)
                try:
                    fnt = ImageFont.truetype(PDF_FONT_PATH, 64) if PDF_FONT_PATH else ImageFont.load_default(size=64)
                except Exception:
                    fnt = ImageFont.load_default()
                alpha_int = int(round(255 * opacity))
                draw.text((40, 110), wm_text[:60], fill=(120, 120, 120, alpha_int), font=fnt)
                rot_img = txt_img.rotate(30, expand=True, resample=Image.Resampling.BICUBIC)
                txt_png = os.path.join(work_dir, "wm_text.png")
                rot_img.save(txt_png, "PNG")
                tw = rect.width * scale_pct
                th = tw * (rot_img.height / max(1.0, float(rot_img.width)))
                target_rect = _compute_placement_rect(fitz, rect, tw, th, position)
                page.insert_image(target_rect, filename=txt_png, keep_proportion=True, overlay=overlay)

        output_path = os.path.join(work_dir, "output.pdf")
        doc.save(output_path, garbage=3, deflate=True)
    finally:
        doc.close()

    return output_path, "_watermarked.pdf", "application/pdf"


@csrf_exempt
@require_POST
@clerk_authenticated
def watermark_pdf(request):
    return _run_single_pdf_tool(
        request, prefix="watermark_", max_size=WATERMARK_MAX_FILE_SIZE, processor=_process_watermark_pdf
    )


# ======================================================================
# NEW TOOL 17: SIGN PDF (Drawn Base64 Signature or Uploaded Image)
# ======================================================================
def _process_sign_pdf(request, input_path, work_dir):
    """
    POST /sign-pdf/
    Form-data:
      file = <pdf>
      signature_data = data:image/png;base64,... (drawn on canvas)
      OR signature_image = <image file>
      position = bottom-right | bottom-left | bottom-center | center | top-right | top-left (default bottom-right)
      scale = size % of page width (10..60, default 22)
      pages = "last" | "all" | "first" | "1,3,5-7" (default "last")
    """
    sig_upload = request.FILES.get("signature_image")
    sig_data = (request.POST.get("signature_data") or "").strip()

    if sig_upload is None and not sig_data:
        raise PDFMergeError("Please draw your signature or upload a signature image.")

    sig_png_path = os.path.join(work_dir, "signature.png")

    if sig_upload is not None:
        raw_sig = _save_typed_upload(
            sig_upload, work_dir, max_size=IMAGE_MAX_FILE_SIZE,
            allowed_exts=IMAGE_EXTENSIONS, magic_check=_is_image_header, label="signature image", index=1,
        )
        _prepare_transparent_watermark_png(raw_sig, sig_png_path, opacity=1.0)
    else:
        if "," in sig_data:
            sig_data = sig_data.split(",", 1)[1]
        try:
            raw_bytes = base64.b64decode(sig_data, validate=False)
        except Exception:
            raise PDFMergeError("Invalid signature data.")
        if len(raw_bytes) < 20 or len(raw_bytes) > 10 * 1024 * 1024:
            raise PDFMergeError("Signature image size is invalid.")
        with open(sig_png_path, "wb") as f:
            f.write(raw_bytes)

    from PIL import Image
    try:
        with Image.open(sig_png_path) as im:
            sig_w, sig_h = im.size
    except Exception:
        raise PDFMergeError("Could not read the signature image.")

    position = (request.POST.get("position") or "bottom-right").strip().lower()
    raw_pages = (request.POST.get("pages") or "last").strip().lower()
    try:
        scale_pct = float(request.POST.get("scale") or "22")
        if scale_pct > 1.0:
            scale_pct /= 100.0
        scale_pct = max(0.08, min(0.75, scale_pct))
    except ValueError:
        scale_pct = 0.22

    fitz = _import_fitz()
    _check_pdf_readable(input_path, SPLIT_MAX_PAGES)

    doc = fitz.open(input_path)
    try:
        total_pages = doc.page_count
        if raw_pages in ("", "last"):
            target_indexes = [total_pages - 1]
        elif raw_pages == "first":
            target_indexes = [0]
        elif raw_pages == "all":
            target_indexes = list(range(total_pages))
        else:
            target_indexes = _expand_page_ranges(_parse_page_ranges(raw_pages), total_pages)

        aspect = sig_h / max(1.0, float(sig_w))
        for idx in target_indexes:
            page = doc[idx]
            rect = page.rect
            target_w = rect.width * scale_pct
            target_h = target_w * aspect
            target_rect = _compute_placement_rect(fitz, rect, target_w, target_h, position, margin=36)
            page.insert_image(target_rect, filename=sig_png_path, keep_proportion=True, overlay=True)

        output_path = os.path.join(work_dir, "output.pdf")
        doc.save(output_path, garbage=3, deflate=True)
    finally:
        doc.close()

    return output_path, "_signed.pdf", "application/pdf"


@csrf_exempt
@require_POST
@clerk_authenticated
def sign_pdf(request):
    return _run_single_pdf_tool(
        request, prefix="sign_", max_size=SIGN_MAX_FILE_SIZE, processor=_process_sign_pdf
    )


# ======================================================================
# NEW TOOL 18: POWERPOINT (.PPTX / .PPT) TO PDF
# ======================================================================
def _process_powerpoint_to_pdf(request, input_path, work_dir):
    output_path = os.path.join(work_dir, "output.pdf")

    # 1. Try LibreOffice headless first for pixel-perfect slide conversion
    for binary in ("libreoffice", "soffice"):
        if shutil.which(binary):
            try:
                subprocess.run(
                    [binary, "--headless", "--convert-to", "pdf", "--outdir", work_dir, input_path],
                    cwd=work_dir, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120, check=False,
                )
                generated = os.path.join(work_dir, Path(input_path).stem + ".pdf")
                if os.path.exists(generated) and os.path.getsize(generated) > 0:
                    os.replace(generated, output_path)
                    return output_path, ".pdf", "application/pdf"
            except Exception as exc:
                logger.warning("LibreOffice PPTX conversion failed, falling back to python-pptx: %s", exc)

    # 2. Pure-Python Fallback using python-pptx + ReportLab (16:9 Landscape slides)
    try:
        from pptx import Presentation
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.lib.units import mm
        from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    except ImportError:
        raise PDFMergeError(
            "Server is missing PowerPoint conversion tools. Install LibreOffice or python-pptx.",
            status=500,
        )

    try:
        prs = Presentation(input_path)
    except Exception:
        raise PDFMergeError("The PowerPoint file is corrupted or unreadable (use .pptx format).")

    if len(prs.slides) == 0:
        raise PDFMergeError("The presentation contains no slides.", status=422)
    if len(prs.slides) > PPTX_MAX_SLIDES:
        raise PDFMergeError(f"Presentation has {len(prs.slides)} slides; maximum is {PPTX_MAX_SLIDES}.", status=413)

    font, font_bold = _register_pdf_font()
    styles = getSampleStyleSheet()
    slide_badge_style = ParagraphStyle(
        "SlideBadge", parent=styles["Normal"], fontName=font_bold, fontSize=9, textColor=colors.HexColor("#7C3AED"), spaceAfter=4
    )
    title_style = ParagraphStyle(
        "SlideTitle", parent=styles["Heading1"], fontName=font_bold, fontSize=22, leading=26, spaceAfter=12
    )
    bullet_style = ParagraphStyle(
        "SlideBullet", parent=styles["Normal"], fontName=font, fontSize=13, leading=19, leftIndent=14, spaceAfter=8
    )

    story = []
    for idx, slide in enumerate(prs.slides, start=1):
        if idx > 1:
            story.append(PageBreak())
        story.append(Paragraph(f"Slide {idx} of {len(prs.slides)}", slide_badge_style))

        title_text = ""
        if slide.shapes.title and slide.shapes.title.text:
            title_text = slide.shapes.title.text.strip()
            if title_text:
                story.append(Paragraph(_xml_escape(title_text), title_style))

        wrote_body = False
        for shape in slide.shapes:
            if shape == slide.shapes.title:
                continue
            if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    txt = para.text.strip()
                    if txt:
                        story.append(Paragraph(f"• {_xml_escape(txt)}", bullet_style))
                        wrote_body = True
            elif getattr(shape, "has_table", False) and shape.has_table:
                t_rows = []
                for row in shape.table.rows:
                    t_rows.append([_xlsx_cell_text(c.text) for c in row.cells])
                if t_rows:
                    tbl = Table(t_rows, repeatRows=1)
                    tbl.setStyle(TableStyle([
                        ("FONTNAME", (0, 0), (-1, -1), font),
                        ("FONTSIZE", (0, 0), (-1, -1), 10),
                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EDE9FE")),
                        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                        ("PADDING", (0, 0), (-1, -1), 5),
                    ]))
                    story.append(Spacer(1, 4 * mm))
                    story.append(tbl)
                    wrote_body = True

        if not title_text and not wrote_body:
            story.append(Paragraph("<i>(Visual / Image Slide)</i>", bullet_style))

    doc = SimpleDocTemplate(
        output_path, pagesize=landscape(A4),
        leftMargin=20 * mm, rightMargin=20 * mm, topMargin=18 * mm, bottomMargin=18 * mm,
        title=Path(input_path).stem, author="SnapPDF",
    )
    doc.build(story)
    return output_path, ".pdf", "application/pdf"


@csrf_exempt
@require_POST
@clerk_authenticated
def powerpoint_to_pdf(request):
    return _run_single_file_tool(
        request, prefix="pptx2pdf_", max_size=PPTX_MAX_FILE_SIZE,
        allowed_exts=PPTX_EXTENSIONS, magic_check=_is_office_header, label="PowerPoint (.pptx)",
        processor=_process_powerpoint_to_pdf,
    )


# ======================================================================
# NEW TOOL 19: ADD PAGE NUMBERS ON EVERY PAGE
# ======================================================================
def _process_add_page_numbers(request, input_path, work_dir):
    """
    POST /add-page-numbers/
    Form-data:
      file = <pdf>
      position = bottom-center | bottom-right | bottom-left | top-center | top-right | top-left (default bottom-center)
      format = "page_of_total" ("Page 1 of 10") | "slash" ("1 / 10") | "simple" ("1") | "dash" ("- 1 -")
      start_number = integer starting number (default 1)
      font_size = 8..36 (default 11)
      color = #RRGGBB (default #000000)
    """
    position = (request.POST.get("position") or "bottom-center").strip().lower()
    fmt = (request.POST.get("format") or "page_of_total").strip().lower()
    color = _hex_to_rgb(request.POST.get("color") or "#333333", "Color")

    try:
        start_num = max(1, int(request.POST.get("start_number") or 1))
    except ValueError:
        start_num = 1

    try:
        font_size = max(7.0, min(36.0, float(request.POST.get("font_size") or 11.0)))
    except ValueError:
        font_size = 11.0

    fitz = _import_fitz()
    _check_pdf_readable(input_path, SPLIT_MAX_PAGES)

    doc = fitz.open(input_path)
    try:
        total = doc.page_count
        for idx, page in enumerate(doc):
            num = start_num + idx
            total_disp = total + start_num - 1
            if fmt == "simple":
                label = f"{num}"
            elif fmt == "slash":
                label = f"{num} / {total_disp}"
            elif fmt == "dash":
                label = f"- {num} -"
            else:
                label = f"Page {num} of {total_disp}"

            rect = page.rect
            box_w = 220
            box_h = font_size * 2.2
            margin_x = 32
            margin_y = 22

            if "top" in position:
                y0 = rect.y0 + margin_y
            else:
                y0 = rect.y1 - margin_y - box_h

            if "left" in position:
                x0 = rect.x0 + margin_x
                align = fitz.TEXT_ALIGN_LEFT
            elif "right" in position:
                x0 = rect.x1 - margin_x - box_w
                align = fitz.TEXT_ALIGN_RIGHT
            else:
                x0 = rect.x0 + (rect.width - box_w) / 2
                align = fitz.TEXT_ALIGN_CENTER

            target_box = fitz.Rect(x0, y0, x0 + box_w, y0 + box_h) * page.derotation_matrix
            page.insert_textbox(
                target_box,
                label,
                fontsize=font_size,
                fontname="helv",
                color=color,
                align=align,
                rotate=page.rotation,
            )

        output_path = os.path.join(work_dir, "output.pdf")
        doc.save(output_path, garbage=3, deflate=True)
    finally:
        doc.close()

    return output_path, "_numbered.pdf", "application/pdf"


@csrf_exempt
@require_POST
@clerk_authenticated
def add_page_numbers(request):
    return _run_single_pdf_tool(
        request, prefix="pagenum_", max_size=PAGENUM_MAX_FILE_SIZE, processor=_process_add_page_numbers
    )