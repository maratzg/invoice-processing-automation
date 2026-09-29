from __future__ import annotations


class InvoiceAutomationError(Exception):
    """Base class for expected, user-actionable automation errors."""


class ConfigurationError(InvoiceAutomationError):
    """Raised when config or filesystem paths are unsafe or invalid."""


class PdfExtractionError(InvoiceAutomationError):
    """Raised when a PDF cannot be read as usable text."""


class MappingError(InvoiceAutomationError):
    """Raised when mapping files are missing, malformed, or ambiguous."""


class TemplateError(InvoiceAutomationError):
    """Raised when the official MONEO template is missing or incompatible."""


class AccountingRuleError(InvoiceAutomationError):
    """Raised when a validated invoice still cannot be converted safely."""

