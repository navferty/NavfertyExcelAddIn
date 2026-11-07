{
  "title": "Валидация XML по XSD",
  "description": "Валидация XML-файла по XSD-схеме. Подробный отчёт об ошибках будет создан в новой книге.",
  "image": "https://navferty.net/images/ValidateXml1.png"
}

[Главная](/ru/) › [Команды](/ru/#commands) › Валидация XML по XSD

# <img alt="" src="/images/icons/validateXml.png" class="icons" width="64" height="64" />Валидация XML по XSD

Валидация XML-файла по XSD-схеме. Подробный отчёт об ошибках будет создан в новой книге.

<div class="shots">

<img src="/images/ValidateXml1.png" loading="lazy" width="883" height="257" alt="Отчёт проверки на листе: тип, элемент, значение и сообщение" />

</div>

## Посмотреть пример отчёта об ошибках

<div class="tbl">

| Серьёзность | Элемент  | Сообщение                                                                                                                         |
|-------------|----------|-----------------------------------------------------------------------------------------------------------------------------------|
| Ошибка      | city     | Элемент 'shipto' содержит недопустимый дочерний элемент 'city'. Ожидаемый список элементов: 'address'.                            |
| Ошибка      | quantity | Элемент 'quantity' недопустим - значение '-5' недопустимо согласно типу данных 'http://www.w3.org/2001/XMLSchema:positiveInteger' |
| Ошибка      | price    | Элемент 'price' недопустим - значение 'asdasd' недопустимо согласно типу данных 'http://www.w3.org/2001/XMLSchema:decimal'        |

</div>

<div class="info-box">

**Статья по теме:** [Проверка XML по XSD-схеме в Excel](/ru/articles/validate-xml-against-xsd-in-excel/)

</div>

Надстройка бесплатная, с открытым кодом. [Инструкция по установке](/ru/install/).
