class Constants:
    """
    Класс, содержащий константы, используемые в Action-классах и связанных функциях.

    Атрибуты:
    - PATTERNS_PRESET_INDEX (int): Индекс шаблона в Photoshop.
    - PS_PIXELS (int): Единица измерения в Photoshop - пиксели.
    - PS_PLACE_AFTER (int): Место вставки слоя после другого в Photoshop.
    - OPTIONS_PNG_FORMAT (int): Формат PNG для параметров экспорта.
    - EXPORT_AS_PNG_FORMAT (int): Формат PNG для экспорта.
    - SAVE_CHANGES (int): Опция сохранения изменений в Photoshop.
    - DO_NOT_SAVE_CHANGES (int): Опция отказа от сохранения изменений в Photoshop.
    - PROMPT_TO_SAVE_CHANGES (int): Опция запроса пользователя о сохранении изменений в Photoshop.
    """

    PATTERNS_PRESET_INDEX = 4
    PS_PIXELS = 1
    PS_PLACE_AFTER = 4

    OPTIONS_PNG_FORMAT = 13
    EXPORT_AS_PNG_FORMAT = 2

    SAVE_CHANGES = 1
    DO_NOT_SAVE_CHANGES = 2
    PROMPT_TO_SAVE_CHANGES = 3
