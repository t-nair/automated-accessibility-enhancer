"""Non-scoring PDF accessibility remediation prototype."""

from .inspect_pdf import inspect_pdf
from .models import DocumentInspection, PageInspection, PageRoute

__all__ = ["DocumentInspection", "PageInspection", "PageRoute", "inspect_pdf"]
