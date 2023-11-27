# Сценарии Adobe Photoshop, написанные на Python-библиотеке win32com. 
Это приложение на языке Python предоставляет возможности по обработке изображений с использованием автоматизации Adobe
Photoshop. Он включает в себя набор классов для создания конкретных эффектов изображений и основной класс для
взаимодействия с объектной моделью Photoshop.

## Доступный функционал
- **BasePhotoshopAction**: Абстрактный базовый класс для добавления новых сценариев Photoshop.


- **DottedImageAction**: Класс для создания точечного эффекта на изображениях.

## Изображение до обработки
<img src="https://github.com/Lyulenko-Maxim/photoshop_python_scripting/blob/master/examples/before.jpg" alt="Изображение до обработки" width="512">

## Изображение после обработки
<img src="https://github.com/Lyulenko-Maxim/photoshop_python_scripting/blob/master/examples/after.png" alt="Изображение после обработки" title="После" width="512"/>

## Требования
- Python 3.11
- Пакетный менеджер pip
- Операционная система Windows
- Установленный Adobe Photoshop в системе

## Установка
1. Склонируйте репозиторий или [скачайте zip-архив](https://github.com/Lyulenko-Maxim/photoshop_python_scripting/archive/refs/heads/master.zip):

   ```bash
   git clone https://github.com/Lyulenko-Maxim/photoshop_python_scripting.git
   
3. Создайте виртуальное окружение:

   ```bash
   python -m venv venv

4. Активируйте виртуальное окружение:

   ```bash
   .\venv\Scripts\activate

5. Установите необходимые зависимости Python:

   ```bash
   pip install -r requirements.txt

## Использование
1. Запустите скрипт main.py.

   ```bash
   py.exe .\main.py

2. Приложение предложит выбрать файл изображения для обработки.


3. Выбранное изображение откроется в Adobe Photoshop.


4. Приложение применит указанный эффект изображения (точечный узор) и предложит сохранить обработанное изображение.


5. Выберите расположение и имя файла для обработанного изображения.


6. Приложение экспортирует изображение в формате PNG с примененным эффектом.


7. Приложение предложит сохранить документ в формате psd с примененным эффектом.


8. Выберите расположение и имя файла для документа.

## Документация
[Python win32com Adobe Photoshop API](https://github.com/lohriialo/photoshop-scripting-python)


## Разработчики
[**Maxim Lyulenko**](https://github.com/Lyulenko-Maxim)

## Лицензия
Этот проект лицензирован по лицензии [**MIT**](https://github.com/Lyulenko-Maxim/photoshop_python_scripting/blob/main/LICENSE).
