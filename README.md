# Установка

## Шаг 1
1. [Скачиваем python, если у вас его нет](https://www.python.org/downloads/).
![123](https://sun9-north.userapi.com/sun9-86/s/v1/ig2/SgZfKWxZBBDYKDX8yNqRfqf1s2LCvt4huQ_czfzoXn4qGeQz7TkqQ7IcaQpWrSLCEVYRzKNDF-1Jv-H8jEbfFWtc.jpg?size=2136x1258&quality=96&type=album)
2. Запускаем инсталятор и ставим все галочки как на скрине ![123](https://sun9-east.userapi.com/sun9-30/s/v1/ig2/qP3n57m55cTHXOKkyfyK9CG2xUoH8BP6PBWyzYkaBVt4NcnGcDKJarJw62djetN6cdb_gEzrqKf-r_IecB-29vgu.jpg?size=993x609&quality=96&type=album)
3. Открываем командну строку (заходим в поиск и вводим 'командная строка', либо используем миллион других способов, которые легко загуглить) ![1234](https://sun1.userapi.com/sun1-17/s/v1/ig2/gQ7LFNTw8dG0sjXhkBVPpNIsjWaWSImivCem-JW40SqFfcr3eRXZNt8RwnajR1rWAaIUEDpn_HAtrsr2d0mc5kv5.jpg?size=2158x1286&quality=96&type=album)
4. Вводим в командную каждую строку из тех, что можно увидеть ниже (ВАЖНО: Вводим построчно и после каждого ввода нажимаем Enter и ожидаем загрузки пакета. Затем повторяем действие еще 8 раз)
```
pip3 install requests
pip3 install python-docx
pip3 install aspose.words 
pip3 install bs4
pip3 install docx
pip3 install py7zr
pip3 install pyunpack
pip3 install pandas
pip3 install openpyxl
```
5. После того как у нас все установилось можно переходить к шагу 2

## Шаг 2
1. Копируем файлы в любую папку (наличие файла *urls.txt* обязательно, как и его название)![123](https://sun9-west.userapi.com/sun9-54/s/v1/ig2/wtlPpfeY6-bT7Gq8o5CWwlUTOOQ-2MpjUF0uRnkj_FutAJiEEnov2c91vP8wc9mTP2EIFh1DUVFT5rTONtkhtRvJ.jpg?size=2160x1440&quality=96&type=album)
2. Открываем файл и ждем некоторое количество времени. Зависит от количества ссылок, которые нужно обработать, а также от степени нагруженности на госзакупки ![123](https://sun9-east.userapi.com/sun9-73/s/v1/ig2/SNOSstP0wTbq8xhxHxNa0ydDSKilfH9ftWaWY1Ph4-OVYj7I6dlvQvEHesDb4z_7u97_2ZlbmJU7Ivj5NqvzP9za.jpg?size=2160x1440&quality=96&type=album)
3. После запуска появится файл *logs.txt*, в котором указаны некорректные ссылки(ссылки другого формата или отсутствие документов для загрузки), а также папка *content*, в которую скачиваются документы для дальнейшей обработки. Папку не нужно трогать и обижать потому что может что-то пойти не так. ![123](https://sun9-north.userapi.com/sun9-84/s/v1/ig2/HdCoKcqmHjxqyGGCXxnmrLARFFxu4dOMOiLjH59B35E-DU8R_CfKEWFE7Cf8FM1PR6XGSA6MsnPLOUKwZeyFA4QK.jpg?size=2160x1440&quality=96&type=album)
4. После завершения обработки *content* удаляется вместе со всеми скачанными файлами и появляется файл *data.xlsx*, в котором находится итоговая информация ![123](https://sun9-east.userapi.com/sun9-32/s/v1/ig2/FqO1K8v8DeC1jPVltSkfI-6foNe2s-Xy84_NVE7lkmCQG6E__6P2JrhlSnhCJzXsGLTj24ZFhd_MPnvLBzImAd74.jpg?size=2160x1440&quality=96&type=album)


## To do

1. Обработка сканов фотографий
