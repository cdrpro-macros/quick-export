# Сборка и установка

## Сборка макроса

Перед тем как использовать макрос в CorelDRAW, его нужно собрать:

1. Скачайте и установите [Visual Studio Community](https://www.visualstudio.com/free-developer-offers/);
1. Откройте Visual Studio (выбирите _Visual C#_ environment при первом запуске);
1. Измените _Solution Configurations_ на _Release_ на стандартной панели (Standard bar);
1. Соберите проект (меню _Buil > Build QuickExport_ / _Shift+F6_).

## Ручная установка

Необходимые файлы можно скопировать вручную, без необходимости создания установщика:

1. Скопируйте макрос _.\bin\Release\QuickExport.dll_ в _C:\Program Files\Corel\CorelDRAW Graphics Suite X7\Programs64\Addons\QuickExport_ (вам нужно будет создать эту папку);
1. Скопируйте установщик _.\installer\QuickExportInstaller.gms_ в _C:\Users\\%your_user%\AppData\Roaming\Corel\CorelDRAW Graphics Suite X7\Draw\GMS_.

## Создание установщика

Или можно создать установщик для более простой установки:

1. Скачайте и установите [Nullsoft Scriptable Install System](http://nsis.sourceforge.net/Main_Page);
1. Скопируйте _.\bin\Release\QuickExport.dll_ в папку _.\installer_;
1. Через контекстное меню файла _*.nsi_ выбирите команду _Compile NSIS Script_.

После этого, вы сможете установить макрос с помощью созданного _*.exe_ файла.

## CorelDRAW

1. Запустите CorelDRAW;
1. Нажмите _OK_ в появившемся окне с вопросом об установке макроса. Если окно не появилось, откройте Менеджер макросов (_Macro Manager_) или запустите _QuickExportInstaller_ вручную.

После этих действий, кнопка макроса появится на Стандартной пенели (_Standard Bar_).
