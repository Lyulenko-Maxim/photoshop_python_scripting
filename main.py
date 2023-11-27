from pathlib import Path
import win32con
import win32gui
import win32ui
from win32com.client import Dispatch

from src import DottedImageAction


class AppProcessor:
    """
    Класс для выполнения действий Photoshop.

    Атрибуты:
    - CELL_SIZE (int): Размер ячейки для создания точечного эффекта.
    - VIBRANCE_VALUE (int): Значение эффекта "сочности".
    - SATURATION_VALUE (int): Начальная насыщенность.
    - EXPORT_FILENAME (str): Имя файла для экспорта.
    - FILE_TYPE (str): Типы файлов для открытия/сохранения.
    - API_FLAG: Флаги для настроек диалогового окна файлов.

    Методы:
    - show_message(message, title, style): Отображение диалогового окна с сообщением.
    - open_file_dialog(mode: int) -> str | None: Открытие диалогового окна для выбора файла/директории.
    - main(): Основной метод для выполнения действий Photoshop.
    """
    CELL_SIZE = 20
    VIBRANCE_VALUE = 0
    SATURATION_VALUE = 20
    EXPORT_FILENAME = 'after'
    FILE_TYPE = 'Image Files (*.bmp;*.jpg;*.png;*.psd)|*.bmp;*.jpg;*.png;*.psd|All Files (*.*)|*.*|'
    API_FLAG = win32con.OFN_OVERWRITEPROMPT | win32con.OFN_FILEMUSTEXIST

    @staticmethod
    def show_message(message, title, style):
        """
        Отображает диалоговое окно с сообщением.

        Параметры:
        - message (str): Текст сообщения.
        - title (str): Заголовок окна.
        - style: Стиль окна.

        Возвращает:
        - Результат отображения окна.
        """
        return win32gui.MessageBox(0, message, title, style)

    @classmethod
    def open_file_dialog(cls, mode: int) -> str | None:
        """
        Открывает диалоговое окно для выбора файла.

        Параметры:
        - mode (int): Режим диалогового окна (1 - открытие, 0 - сохранение).

        Возвращает:
        - Путь к выбранному файлу или None в случае отмены.
        """
        dialog = win32ui.CreateFileDialog(mode, None, None, cls.API_FLAG, cls.FILE_TYPE)
        dialog.SetOFNInitialDir(str(Path.cwd()))

        if not dialog.DoModal() == 1:
            cls.show_message(
                message='Отменено пользователем.',
                title='Уведомление',
                style=win32con.MB_OK | win32con.MB_ICONINFORMATION
            )
            exit()

        return dialog.GetPathName()

    @classmethod
    def main(cls):
        """
        Основной метод для выполнения действий Photoshop.
        """
        app = Dispatch("Photoshop.Application")
        file_path = cls.open_file_dialog(mode=1)
        app.Open(file_path)

        action = DottedImageAction(
            app=app,
            cell_size=cls.CELL_SIZE,
            vibrance=cls.VIBRANCE_VALUE,
            saturation=cls.SATURATION_VALUE
        )
        action.execute()

        save_path = cls.open_file_dialog(mode=0)
        action.export_as_png(save_path=save_path)
        action.close_active_document()


if __name__ == '__main__':
    AppProcessor.main()
