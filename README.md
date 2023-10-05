# Yandex_Maps_Parser
### Парсер сервиса Yandex Maps https://yandex.ru/maps/

Основой работы парсера является библиотека "Selenium" 

Полученные данные автоматически записываются в файл .xlsx для более удобного восприятия

## RUN 
```commandline
git clone https://github.com/GCTMLP/ya_maps_parser.git
cd YandexMapsParser
pip3 install -r requirements.txt
python3 ya_maps_parse.py -p Город -c Красота -f путь_до_файла
```

## Список собираемой информации об организации

- Название организации
- Адрес
- Телефоны
- Сайт организации
- Страницы организации в соц. сетях
- Рейтинг

Список собираемой информации можно расширить, добавив в метод "get_additional_data" свой код по образу и подобию
