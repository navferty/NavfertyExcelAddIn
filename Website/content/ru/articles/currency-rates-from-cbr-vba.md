{
  "title": "Курсы валют с сайта ЦБ на VBA с автообновлением",
  "description": "Как загрузить курсы валют из XML API Банка России в Excel с помощью VBA и пользовательской функции, которая обновляется автоматически.",
  "image": "https://navferty.net/images/articles/currency-rates-from-cbr-vba/1590228471139931720.png"
}

[← Все статьи](/ru/articles/)

# Курсы валют с сайта ЦБ на VBA с автообновлением

Как загрузить курсы валют из XML API Банка России в Excel с помощью VBA и пользовательской функции, которая обновляется автоматически.

Частый запрос от пользователей Excel: макрос, который загружает курсы валют с сайта ЦБ РФ, а формулы при этом обновляются автоматически и подтягивают актуальный курс на текущую дату. Готовый вариант решения лежит на GitHub: <a href="https://github.com/navferty/CBR-VBA-Currencies" rel="noopener" target="_blank">https://github.com/navferty/CBR-VBA-Currencies</a>. Ниже разобрано, как он делается по шагам.

<img src="/images/articles/currency-rates-from-cbr-vba/1590231961122251017.webp" loading="lazy" width="650" height="427" alt="Страница ЦБ РФ с описанием API курсов валют" />

## API Центробанка

На сайте ЦБ РФ есть официальное API для загрузки динамики курсов (пригодится, если понадобится курс на определённую дату в прошлом). Это HTTP GET-запрос вида:

    http://www.cbr.ru/scripts/XML_dynamic.asp?date_req1=02/03/2001&date_req2=14/03/2001&VAL_NM_RQ=R01235

Здесь `R01235` — код валюты (доллар США). Список кодов валют для запроса тоже есть на сайте ЦБ.

API возвращает данные в формате XML. Для VBA это даже удобнее, чем JSON: поддержка XML есть из коробки, а для разбора JSON нужны сторонние библиотеки. Пример ответа:

<img src="/images/articles/currency-rates-from-cbr-vba/1590228257130434333.webp" loading="lazy" width="700" height="329" alt="Пример XML-ответа с курсами на каждую дату" />

В ответе есть курс на каждую дату интервала с указанием кода валюты, то есть ровно то, что нужно.

## Запрос из VBA

Теперь сделаем такой запрос из VBA. Сначала подключим к книге ссылку на MSXML 6.0 (если доступна другая версия, ничего страшного: могут слегка отличаться имена типов, логика та же). Для этого в редакторе VBA откройте Tools → References.

<img src="/images/articles/currency-rates-from-cbr-vba/1590228471139931720.png" loading="lazy" width="445" height="364" alt="Окно References с MSXML 6.0 и Microsoft Scripting Runtime" />

Заодно добавим ссылку на Microsoft Scripting Runtime: оттуда берём тип `Dictionary` для хранения соответствия кодов валют (USD → R01235).

Создадим простую функцию, которая отправляет запрос и загружает ответ в XML-документ:

    Private Function RequestGetXml() As DOMDocument60
        Dim XMLHTTP As New MSXML2.XMLHTTP60
        Dim requestUrl As String

        requestUrl = "http://www.cbr.ru/scripts/XML_dynamic.asp?date_req1=02/03/2001&date_req2=14/03/2001&VAL_NM_RQ=R01235"
        XMLHTTP.Open "GET", requestUrl, False
        XMLHTTP.send

        Dim resultXmlDocument As DOMDocument60
        Set resultXmlDocument = XMLHTTP.responseXML

        Stop ' look at the document in the Locals window

        Set RequestGetXml = resultXmlDocument
    End Function

Механизм XmlHttpRequest, в том числе с авторизацией, хорошо расписан здесь: <a href="https://codingislove.com/http-requests-excel-vba/" rel="noopener" target="_blank">https://codingislove.com/http-requests-excel-vba/</a>.

Команда `Stop` приостанавливает выполнение, и в окне Locals можно посмотреть на полученный XML-документ.

## Класс для хранения курса

Теперь у нас есть XML-документ, по которому можно пройти в цикле и достать курс на каждую дату. Но сначала добавим класс для хранения этих значений. Нужны три свойства: код валюты (строка), дата и само значение (для денежных сумм лучше использовать тип Decimal). Добавим class module с именем `CurrencyRecord`:

    Option Explicit

    Private prCurrencyDate As Date
    Private prCurrencyValue As Variant 'decimal
    Private prCurrencyCode As String

    Public Property Get CurrencyDate() As Date
        CurrencyDate = prCurrencyDate
    End Property
    Public Property Let CurrencyDate(ByVal vNewValue As Date)
        prCurrencyDate = vNewValue
    End Property

    Public Property Get CurrencyValue() As Variant
        CurrencyValue = prCurrencyValue
    End Property
    Public Property Let CurrencyValue(ByVal vNewValue As Variant)
        prCurrencyValue = CDec(vNewValue)
    End Property

    Public Property Get CurrencyCode() As String
        CurrencyCode = prCurrencyCode
    End Property
    Public Property Let CurrencyCode(ByVal vNewValue As String)
        prCurrencyCode = vNewValue
    End Property

Свойства можно добавлять командой Insert → Procedure..., она сама сгенерирует геттеры и сеттеры. Можно обойтись и публичными полями, но доступ через свойства позволяет управлять данными и считается хорошим тоном.

Функция `CDec` преобразует значение в десятичное число (тип Decimal). Для денежных сумм советую использовать именно его, а не Double: это позволяет избежать ошибок округления десятичных дробей.

<img src="/images/articles/currency-rates-from-cbr-vba/1590232854180743238.png" loading="lazy" width="307" height="177" alt="Преобразование значения функцией CDec" />

## Обход ответа

Вернёмся к задаче. Создадим коллекцию, которую будем заполнять записями типа `CurrencyRecord`:

    Private Function LoadUsdItems() As Collection

        Dim resultXmlDocument As DOMDocument60
        Dim recordItem As IXMLDOMElement
        Dim col As New Collection
        Dim resultItem As CurrencyRecord

        Set resultXmlDocument = RequestGetXml()

        For Each recordItem In resultXmlDocument.LastChild.ChildNodes
            Set resultItem = New CurrencyRecord
            With resultItem
                .CurrencyCode = recordItem.Attributes.getNamedItem("Id").NodeValue
                .CurrencyDate = CDate(recordItem.Attributes.getNamedItem("Date").NodeValue)
                .CurrencyValue = CDec(recordItem.ChildNodes.Item(1).nodeTypedValue)
            End With
            col.Add resultItem
        Next

        Set LoadUsdItems = col
    End Function

Мы обходим элементы, в которых хранятся курсы на каждую дату. Обращение к `resultXmlDocument.LastChild.ChildNodes` можно переделать на XPath, чтобы навигация по документу была нагляднее.

## Запрос по коду валюты

Убедившись, что запрос и разбор ответа работают, перепишем функцию так, чтобы она запрашивала курс по коду валюты. Добавим словарь соответствия кодов и сборку адреса запроса с учётом дат начала и окончания периода и кода валюты:

    ' returns a Collection of CurrencyRecord objects
    Private Function GetCurrency(currCode As String, startDate As Date, endDate As Date) As Collection

        Dim codeDict As Dictionary
        Set codeDict = New Dictionary
        codeDict.Add "USD", "R01235"
        codeDict.Add "GBP", "R01035"
        codeDict.Add "BYN", "R01090B"
        ' add other currencies here

        Dim currCbrCode As String
        currCbrCode = codeDict.Item(currCode)

        Dim requestUri As String
        requestUri = "http://www.cbr.ru/scripts/XML_dynamic.asp?" _
                    + "date_req1=" + Format(startDate, "dd/mm/yyyy") _
                    + "&date_req2=" + Format(endDate, "dd/mm/yyyy") _
                    + "&VAL_NM_RQ=" + currCbrCode
        ' Date format is important!

        Dim resp As DOMDocument60
        Set resp = RequestGetXml(requestUri)

        Dim col As New Collection
        Dim recordItem As IXMLDOMElement
        Dim resultItem As CurrencyRecord
        For Each recordItem In resp.LastChild.ChildNodes
            Set resultItem = New CurrencyRecord
            With resultItem
                .CurrencyCode = currCode
                .CurrencyDate = CDate(recordItem.Attributes.getNamedItem("Date").NodeValue)
                .CurrencyValue = CDec(recordItem.ChildNodes.Item(1).nodeTypedValue)
            End With
            col.Add resultItem
        Next

        Set GetCurrency = col
    End Function

Я также заменил `CurrencyCode` при создании записи в цикле: вместо `"R01235"` запись содержит общеизвестный код валюты `"USD"`.

`RequestGetXml` — вспомогательная функция, которая загружает XML-документ HTTP-запросом (теперь адрес передаётся аргументом):

    Private Function RequestGetXml(requestUrl As String) As DOMDocument60
        'about requests in VBA: https://codingislove.com/http-requests-excel-vba/
        Dim XMLHTTP As New MSXML2.XMLHTTP60

        XMLHTTP.Open "GET", requestUrl, False
        XMLHTTP.send
        Set RequestGetXml = XMLHTTP.responseXML
    End Function

Проверим, что вся функция работает как ожидается, добавив временную процедуру для теста:

    Private Sub TestCurrLoad()
        Dim col As Collection
        Set col = GetCurrency("USD", #1/1/2020#, #2/2/2020#)
        Stop
    End Sub

## Функция для листа

Теперь добавим публичную функцию, которую можно использовать в формулах на листе (UDF, user-defined function):

    Public Function GetTodayCurrency(currCode As String, volatileArg As Variant) As Variant 'decimal
        Dim currItem As CurrencyRecord
        Dim col As Collection
        Set col = GetCurrency(currCode, DateAdd("d", -1, Date), Now)

        If col.Count = 0 Then
            Exit Function
        End If

        Set currItem = col.Item(col.Count) ' VBA collections are numbered from 1
        GetTodayCurrency = currItem.CurrencyValue

        Debug.Print "Currency loaded, value is " + CStr(currItem.CurrencyValue)
    End Function

С помощью `DateAdd` определяем интервал дат для загрузки: от вчера до сегодня. Получив две записи, берём значение из последней (нумерация в коллекциях VBA начинается с 1).

Обратите внимание на второй аргумент, `volatileArg`. Его значение не используется, но он нужен, чтобы заставить Excel пересчитывать результат при изменении этого аргумента в формуле ячейки:

    =GetTodayCurrency("USD"; NOW())

Чтобы убедиться, что курс обновился, в функцию добавлен вызов `Debug.Print` с сообщением об успешной загрузке. После отладки его можно удалить, чтобы не засорять окно Immediate.

## Проверка

Проверим, что курс загружается при открытии книги:

<img src="/images/articles/currency-rates-from-cbr-vba/1590231043185812459.webp" loading="lazy" width="700" height="346" alt="Окно Immediate с сообщением об успешной загрузке курса" />

В окне Immediate видно сообщение об успешной загрузке сразу после открытия книги. Такое же сообщение появится при любом изменении на листе, потому что Excel пересчитывает формулу. Если в книге много ячеек с вызовом этой функции, лучше добавить кэш: например, статическую переменную со словарём (`Dictionary`), где ключ — сочетание кода валюты и даты, а значение — `CurrencyRecord`. Это избавит от множества одинаковых запросов к сайту ЦБ.

Если нужны доработки, их можно предложить в issues репозитория <a href="https://github.com/navferty/CBR-VBA-Currencies" rel="noopener" target="_blank">https://github.com/navferty/CBR-VBA-Currencies</a>.

Статья впервые опубликована на <a href="https://pikabu.ru/story/kursyi_valyut_s_sayta_tsb_na_vba_s_avtoobnovleniem_7469372" rel="noopener" target="_blank">Пикабу</a>.
