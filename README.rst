Пакетная рассылка почты
========================

Для проведения конкретной рассылки необходимы два файла: шаблон письма и список с адресами и текстами подстановки.

Имя шаблона должно заканчиваться на «text.html». Пример шаблона::

    <h2>{Как_обращаться}, добрый день!</h2>

    <p>
    Это — простое письмо без излишеств.
    Вот только здесь {константа}.
    </p>

    <p>
    С уважением,
    <br/>С.А. Шашков
    </p>


Шаблон, к сожалению, должен быть в формате html.
Однако во всех простых случаях формат совершенно банальный, но позволяет вставлять в письмо некоторое оформление,
таблицы, ссылки и т.д.
Так, абзац текста должен быть окружён тегами ``<p> … </p>``.
Внутри абзаца можно начать новую строку при помощи тега ``<br/>``.
Заголовок письма можно добавить командами ``<h3>…</h3>``.
Ссылки добавляются в виде ``<a href="http://address.ru/file.pdf">Описание ссылки</a>``.
Бояться не нужно, будет предпросмотр.

Список рассылки должен лежать в excel-файле с именем, заканчивающимся на «list.xlsx». Пример:


Установка
-----

Для установки пакета необходимо выполнить в терминале

``pip install git+https://github.com/ShashkovS/batch_email_sender``

В первой строке должны быть заголовки столбцов. Первый столбец должен называться «ok». В него рассыльщик проставить ok по тем письмам, которые удалось отправить. Один из столбцов должен называться «email», в нём адресат письма. Также должен быть столбец «subject» — тема конкретного письма. Кроме того, для каждого текста-подстановки вида {Как_обращаться} из шаблона тоже должен быть столбец. Его содержимое и будет подставляться. Кроме этого в таблице могут быть любые данные, они не будут использоваться.
Примеры файлов выше есть вместе с рассыльщиком.

License
-------

This is free and unencumbered software released into the public domain.

Anyone is free to copy, modify, publish, use, compile, sell, or
distribute this software, either in source code form or as a compiled
binary, for any purpose, commercial or non-commercial, and by any
means.
