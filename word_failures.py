from enum import Enum, auto


class WordFailureType(Enum):
    FILE_NOT_FOUND = auto()
    FILE_LOCKED = auto()
    PROTECTED_VIEW = auto()
    CORRUPTED = auto()
    MACRO_FAILURE = auto()
    COM_FAILURE = auto()
    PERMISSION_DENIED = auto()
    INVALID_FORMAT = auto()
    TIMEOUT = auto()
    UNKNOWN = auto()

    def __str__(self):
        return self.name.replace('_', ' ').title()

    @property
    def is_recoverable(self):
        recoverable_types = {
            WordFailureType.FILE_LOCKED,
            WordFailureType.PROTECTED_VIEW,
            WordFailureType.COM_FAILURE,
            WordFailureType.TIMEOUT
        }
        return self in recoverable_types