"""Errors a caller may want to handle separately. The CLI turns the usage errors
into a usage message before any conversion starts. All of them subclass ValueError."""


class DocxToPdfError(ValueError):
    """Base class for this package's errors."""


class UnsupportedFileType(DocxToPdfError):
    """The file's suffix is not a Word document type this tool converts."""


class TargetConflict(DocxToPdfError):
    """A PDF would overwrite an existing file, or two inputs would write the same PDF."""


class ConversionFailed(DocxToPdfError):
    """This one document could not be converted (corrupt, encrypted, unreadable)."""
