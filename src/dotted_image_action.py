from src import BasePhotoshopAction, Constants


class DottedImageAction(BasePhotoshopAction):
    """
    Класс для создания точечного эффекта на изображении в Photoshop.

    Атрибуты:
    - PHOTO_LAYER_NAME (str): Название слоя с исходным изображением.
    - COLOR_FILL_LAYER_NAME (str): Название слоя для заливки цветом.
    - PATTERN_NAME_PREFIX (str): Префикс для именования узора.

    Методы:
    - create_dotted_pattern(self, pattern_name: str): Создание точечного узора.
    - get_dotted_pattern_name(self) -> str: Получение имени точечного узора.
    - execute(self): Выполнение сценария точечного эффекта.
    """

    PHOTO_LAYER_NAME = 'Photo'
    COLOR_FILL_LAYER_NAME = 'Color fill'
    PATTERN_NAME_PREFIX = 'Circle'

    def __init__(self, app, cell_size, vibrance, saturation):
        """
        Инициализирует экземпляр класса DottedImageAction.

        Параметры:
        - app: Объект приложения Photoshop.
        - cell_size (int): Размер ячейки для создания точечного эффекта.
        - vibrance (int): Значение эффекта "сочности".
        - saturation (int): Начальная насыщенность.
        """
        super().__init__(app)
        self.cell_size = cell_size
        self.vibrance = vibrance
        self.saturation = saturation
        self.app.Preferences.RulerUnits = Constants.PS_PIXELS
        self.pattern_name = f'{self.PATTERN_NAME_PREFIX}{self.cell_size}'

    def create_dotted_pattern(self, pattern_name: str):
        """
        Создает точечный узор.

        Параметры:
        - pattern_name (str): Имя для нового узора.
        """
        document = self.app.Documents.Add(Width=int(self.cell_size), Height=int(self.cell_size))
        document.ActiveLayer.IsBackgroundLayer = False
        self.app.ActiveDocument = document

        self.circle_selection(center_x=10, center_y=10, radius=10)
        document.Selection.Fill(FillType=self.rgb_color(red=0, green=0, blue=0))
        document.Selection.Deselect()
        document.ActiveLayer.Invert()

        self.define_pattern(pattern_name=pattern_name)
        self.close_active_document(Constants.DO_NOT_SAVE_CHANGES)

    def get_dotted_pattern_name(self) -> str:
        """
        Получает имя точечного узора, создавая его при необходимости.

        Возвращает:
        - Имя точечного узора.
        """
        if not self.is_pattern_exist(pattern_name=self.pattern_name):
            self.create_dotted_pattern(pattern_name=self.pattern_name)

        return self.pattern_name

    def execute(self):
        """
        Переопределяет метод базового класса,
        выполняет сценарий точечного эффекта на изображении.
        """

        self.convert_image_to_rgb_mode()
        self.convert_current_layer_to_smart_object()

        photo_layer = self.doc.ActiveLayer
        photo_layer.Name = self.PHOTO_LAYER_NAME

        color_fill_layer = self.doc.ArtLayers.Add()
        color_fill_layer.Name = self.COLOR_FILL_LAYER_NAME

        self.select_layer()
        self.doc.Selection.Fill(FillType=self.rgb_color(red=0, green=0, blue=0))
        self.doc.Selection.Deselect()

        color_fill_layer.Move(RelativeObject=photo_layer, InsertionLocation=Constants.PS_PLACE_AFTER)
        self.doc.ActiveLayer = photo_layer

        self.apply_mosaic(cell_size=self.cell_size)
        self.select_layer()
        self.apply_layer_mask()

        pattern_name = self.get_dotted_pattern_name()
        self.apply_pattern(pattern_name=pattern_name)

        self.apply_vibrance(vibrance=self.vibrance, saturation=self.saturation)
