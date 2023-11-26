import abc
from abc import ABC
from win32com.client import Dispatch

from src import Constants


class BasePhotoshopAction(ABC):
    """
    Абстрактный базовый класс для сценариев в Photoshop.

    Атрибуты:
    - app: Объект приложения Photoshop.
    - doc: Текущий активный документ в Photoshop.

    Абстрактные методы:
    - execute(): Метод, который должен быть реализован в подклассах для выполнения сценария.
    """

    def __init__(self, app):
        """
        Инициализирует экземпляр класса BasePhotoshopAction.

        Параметры:
        - app: Объект приложения Photoshop.
        """
        self.app = app
        self.doc = self.app.ActiveDocument

    @abc.abstractmethod
    def execute(self):
        """
        Абстрактный метод для выполнения сценария в Photoshop.
        """
        raise NotImplementedError("Can't execute base class action")

    @staticmethod
    def rgb_color(red, green, blue):
        """
        Создает объект SolidColor с заданными значениями RGB.

        Параметры:
        - red (int): Красный компонент.
        - green (int): Зеленый компонент.
        - blue (int): Синий компонент.

        Возвращает:
        - Объект SolidColor.
        """
        color = Dispatch("Photoshop.SolidColor")
        color.RGB.Red = red
        color.RGB.Green = green
        color.RGB.Blue = blue
        return color

    def s(self, string):
        """
        Возвращает тип данных (TypeID) для строки.

        Параметры:
        - string (str): Входная строка.

        Возвращает:
        - Тип данных (TypeID) для строки.
        """
        return self.app.StringIDToTypeID(string)

    def c(self, char):
        """
        Возвращает тип данных (TypeID) для символьной строки.

        Параметры:
        - char (str): Входной символьную строку.

        Возвращает:
        - Тип данных (TypeID) для символьной строки.
        """
        return self.app.CharIDToTypeID(char)

    def apply_layer_mask(self):
        """
        Применяет маску слоя в Photoshop.
        """
        descriptor = Dispatch("Photoshop.ActionDescriptor")
        reference = Dispatch("Photoshop.ActionReference")
        descriptor.PutClass(self.s('new'), self.s('channel'))
        reference.PutEnumerated(self.s('channel'), self.s('channel'), self.s('mask'))
        descriptor.PutReference(self.s('at'), reference)
        descriptor.PutEnumerated(self.s('using'), self.s('userMaskEnabled'), self.s('revealSelection'))
        self.app.ExecuteAction(self.s('make'), descriptor)

    def convert_current_layer_to_smart_object(self):
        """
        Конвертирует текущий слой в умный объект в Photoshop.
        """
        descriptor = Dispatch("Photoshop.ActionDescriptor")
        self.app.ExecuteAction(self.s('newPlacedLayer'), descriptor)

    def apply_mosaic(self, cell_size):
        """
        Применяет эффект мозаики с заданным размером ячейки в Photoshop.

        Параметры:
        - cell_size (int): Размер ячейки мозаики.
        """
        desc = Dispatch("Photoshop.ActionDescriptor")
        desc.PutUnitDouble(self.c('ClSz'), self.c('#Pxl'), cell_size)
        self.app.ExecuteAction(self.c('Msc '), desc)

    def select_area(self, offset):
        """
        Выделяет область на изображении с заданным отступом в Photoshop.

        Параметры:
        - offset (int): Отступ от границ изображения.
        """
        descriptor = Dispatch("Photoshop.ActionDescriptor")
        reference = Dispatch("Photoshop.ActionReference")
        reference.PutProperty(self.c("Chnl"), self.c("fsel"))
        descriptor.PutReference(self.c("null"), reference)
        rect_descriptor = Dispatch("Photoshop.ActionDescriptor")
        rect_descriptor.PutUnitDouble(self.c("Top "), self.c("#Pxl"), offset)
        rect_descriptor.PutUnitDouble(self.c("Left"), self.c("#Pxl"), offset)
        rect_descriptor.PutUnitDouble(self.c("Btom"), self.c("#Pxl"), self.doc.Height - offset)
        rect_descriptor.PutUnitDouble(self.c("Rght"), self.c("#Pxl"), self.doc.Width - offset)
        descriptor.PutObject(self.c("T   "), self.c("Rctn"), rect_descriptor)
        self.app.ExecuteAction(self.c("setd"), descriptor)

    def select_layer(self):
        """
        Выделяет текущий слой в Photoshop.
        """
        self.select_area(offset=0)

    def circle_selection(self, center_x, center_y, radius):
        """
        Выделяет круговую область на изображении в Photoshop.

        Параметры:
        - center_x (int): Координата X центра круга.
        - center_y (int): Координата Y центра круга.
        - radius (int): Радиус круга.
        """
        descriptor = Dispatch("Photoshop.ActionDescriptor")
        reference = Dispatch("Photoshop.ActionReference")
        reference.PutProperty(self.c('Chnl'), self.c('fsel'))
        descriptor.PutReference(self.c('null'), reference)
        ellipse_descriptor = Dispatch("Photoshop.ActionDescriptor")
        ellipse_descriptor.PutUnitDouble(self.c('Top '), self.c('#Pxl'), center_y - radius)
        ellipse_descriptor.PutUnitDouble(self.c('Left'), self.c('#Pxl'), center_x - radius)
        ellipse_descriptor.PutUnitDouble(self.c('Btom'), self.c('#Pxl'), center_y + radius)
        ellipse_descriptor.PutUnitDouble(self.c('Rght'), self.c('#Pxl'), center_x + radius)
        descriptor.PutObject(self.c('T   '), self.c('Elps'), ellipse_descriptor)
        descriptor.PutBoolean(self.c('AntA'), True)
        self.app.ExecuteAction(self.c('setd'), descriptor)

    def is_pattern_exist(self, pattern_name) -> bool:
        """
        Проверяет существование узора с заданным именем в Photoshop.

        Параметры:
        - pattern_name (str): Имя узора.

        Возвращает:
        - True, если узор существует, иначе False.
        """
        reference = Dispatch("Photoshop.ActionReference")
        reference.PutProperty(self.s('property'), self.s('presetManager'))
        reference.PutEnumerated(self.s('application'), self.s('ordinal'), self.s('targetEnum'))
        descriptor = self.app.ExecuteActionGet(reference)
        presets_list = descriptor.GetList(self.app.StringIDToTypeID('presetManager'))
        name_list = presets_list.GetObjectValue(
            Constants.PATTERNS_PRESET_INDEX
        ).GetList(self.app.StringIDToTypeID('name'))

        for i in range(name_list.Count):
            if name_list.GetString(i) == pattern_name:
                return True
        return False

    def define_pattern(self, pattern_name):
        """
        Определяет новый узор с заданным именем в Photoshop.

        Параметры:
        - pattern_name (str): Имя нового узора.
        """
        descriptor = Dispatch("Photoshop.ActionDescriptor")
        descriptor.PutString(self.c('Nm  '), pattern_name)
        self.app.ExecuteAction(self.c('DfnP'), descriptor)

    def apply_pattern(self, pattern_name: str, opacity: float = 100.000000):
        """
        Применяет узор к слою в Photoshop.

        Параметры:
        - pattern_name (str): Имя узора.
        - opacity (float): Прозрачность узора. По умолчанию 100.
        """
        descriptor = Dispatch("Photoshop.ActionDescriptor")
        descriptor2 = Dispatch("Photoshop.ActionDescriptor")
        descriptor.PutEnumerated(self.s('using'), self.s('fillContents'), self.s('pattern'))
        descriptor2.PutString(self.s('name'), pattern_name)
        descriptor.PutObject(self.s('pattern'), self.s('pattern'), descriptor2)
        descriptor.PutUnitDouble(self.s('opacity'), self.s('percentUnit'), opacity)
        descriptor.PutEnumerated(self.s('mode'), self.s('blendMode'), self.s('normal'))
        self.app.ExecuteAction(self.s('fill'), descriptor)

    def apply_vibrance(self, vibrance: int, saturation: int):
        """
        Применяет эффект "сочность" к изображению в Photoshop.

        Параметры:
        - vibrance (int): Значение эффекта "сочность".
        - saturation (int): Начальная насыщенность.
        """
        self.doc.activeChannels = [self.doc.channels[i] for i in range(0, 4)]
        descriptor = Dispatch("Photoshop.ActionDescriptor")
        reference = Dispatch("Photoshop.ActionReference")
        reference.PutClass(self.c('AdjL'))
        descriptor.PutReference(self.c('null'), reference)
        descriptor_using = Dispatch("Photoshop.ActionDescriptor")
        descriptor_using.PutClass(self.c('Type'), self.s('vibrance'))
        descriptor.PutObject(self.c('Usng'), self.c('AdjL'), descriptor_using)
        self.app.ExecuteAction(self.c('Mk  '), descriptor)
        self.set_vibrance_parameters(vibrance=vibrance, saturation=saturation)

    def set_vibrance_parameters(self, vibrance: int, saturation: int):
        """
        Устанавливает параметры эффекта "сочность" в Photoshop.

        Параметры:
        - vibrance (int): Значение эффекта "сочность".
        - saturation (int): Начальная насыщенность.
        """
        descriptor = Dispatch("Photoshop.ActionDescriptor")
        reference = Dispatch("Photoshop.ActionReference")
        reference.PutEnumerated(self.c('AdjL'), self.c('Ordn'), self.c('Trgt'))
        descriptor.PutReference(self.c('null'), reference)
        descriptor2 = Dispatch("Photoshop.ActionDescriptor")
        descriptor2.PutInteger(self.s('vibrance'), vibrance)
        descriptor2.PutInteger(self.c('Strt'), saturation)
        descriptor.PutObject(self.c('T   '), self.s('vibrance'), descriptor2)
        self.app.ExecuteAction(self.c('setd'), descriptor)

    def export_as_png(self, save_path: str):
        """
        Экспортирует изображение в формате PNG в Photoshop.

        Параметры:
        - save_path (str): Путь для сохранения файла.
        """
        options = Dispatch('Photoshop.ExportOptionsSaveForWeb')
        options.Format = Constants.OPTIONS_PNG_FORMAT
        self.doc.Export(
            ExportIn=f'{save_path}.png',
            ExportAs=Constants.EXPORT_AS_PNG_FORMAT,
            Options=options
        )

    def close_active_document(self, option: int = Constants.PROMPT_TO_SAVE_CHANGES):
        """
        Закрывает текущий активный документ в Photoshop с опцией сохранения изменений.

        Параметры:
        - option (int): Опция сохранения изменений. По умолчанию - закрытие с вопросом о сохранении.
        """
        self.app.ActiveDocument.Close(option)
