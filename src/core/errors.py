class PptxAgentException(Exception):
    """Базовый класс для ошибок генерации презентации."""
    pass

class PptxSyntaxError(PptxAgentException):
    """Выбрасывается, если ИИ сгенерировал невалидный синтаксис (например, кривые координаты)."""
    pass

class PptxLogicError(PptxAgentException):
    """Выбрасывается при логических ошибках (попытка обратиться к несуществующему ID и т.д.)."""
    pass