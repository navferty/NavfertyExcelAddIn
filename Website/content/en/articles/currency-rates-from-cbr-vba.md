{
  "title": "Currency rates from the Bank of Russia website in VBA",
  "description": "How to load exchange rates from the Bank of Russia (CBR) XML API into Excel with VBA and a user-defined function that refreshes automatically.",
  "image": "https://navferty.net/images/articles/currency-rates-from-cbr-vba/1590228471139931720.png"
}

[← All articles](/en/articles/)

# Currency rates from the Bank of Russia website in VBA

How to load exchange rates from the Bank of Russia (CBR) XML API into Excel with VBA and a user-defined function that refreshes automatically.

A common request from Excel users is a macro that loads exchange rates from the Bank of Russia (CBR) website, with formulas that refresh themselves and pick up the current rate for today. A ready-made solution is available on GitHub: <a href="https://github.com/navferty/CBR-VBA-Currencies" rel="noopener" target="_blank">https://github.com/navferty/CBR-VBA-Currencies</a>. Below is how it is built, step by step.

<img src="/images/articles/currency-rates-from-cbr-vba/1590231961122251017.webp" loading="lazy" width="650" height="427" alt="The Bank of Russia page describing the exchange rate API" />

## The Bank of Russia API

The CBR website has an official API that returns the dynamics of exchange rates (handy if you need a rate for a specific date in the past). It is an HTTP GET request like this:

    http://www.cbr.ru/scripts/XML_dynamic.asp?date_req1=02/03/2001&date_req2=14/03/2001&VAL_NM_RQ=R01235

Here `R01235` is the currency code (US dollar). The list of currency codes for the request is also published on the CBR website.

The API returns XML. For VBA that is actually easier than JSON, since XML is supported out of the box while parsing JSON needs third-party libraries. A sample response:

<img src="/images/articles/currency-rates-from-cbr-vba/1590228257130434333.webp" loading="lazy" width="700" height="329" alt="An example XML response with a rate for each date" />

The response contains a rate for every date in the interval together with the currency code, which is exactly what we need.

## Making the request from VBA

Now let us make this request from VBA. First, add a reference to MSXML 6.0 in the workbook (if another version is available that is fine: type names may differ slightly, but the logic is the same). In the VBA editor, open Tools → References.

<img src="/images/articles/currency-rates-from-cbr-vba/1590228471139931720.png" loading="lazy" width="445" height="364" alt="The References dialog with MSXML 6.0 and Microsoft Scripting Runtime" />

While there, also add a reference to Microsoft Scripting Runtime. It provides the `Dictionary` type, which we will use to map currency codes (USD → R01235).

Let us create a simple function that sends the request and loads the response into an XML document:

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

The XmlHttpRequest mechanism, including authorization examples, is described well here: <a href="https://codingislove.com/http-requests-excel-vba/" rel="noopener" target="_blank">https://codingislove.com/http-requests-excel-vba/</a>.

The `Stop` statement pauses execution so you can inspect the received XML document in the Locals window.

## A class to hold the rate

Now we have an XML document that we can walk through in a loop, extracting the rate for each date. But first let us add a class to store those values. We need three properties: the currency code (a string), the date and the value itself (for monetary amounts it is better to use the Decimal type). Add a class module named `CurrencyRecord`:

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

Properties can be added with Insert → Procedure..., which generates the getters and setters automatically. Public fields would also work, but access through properties lets you control the data and is considered good practice.

The `CDec` function converts a value to a decimal number (the Decimal type). I strongly recommend it over Double for monetary amounts, because it avoids rounding errors with decimal fractions.

<img src="/images/articles/currency-rates-from-cbr-vba/1590232854180743238.png" loading="lazy" width="307" height="177" alt="Converting a value with CDec" />

## Walking through the response

Back to the task. Let us create a collection and fill it with `CurrencyRecord` objects:

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

The loop goes through the elements that hold the rate for each date. The access to `resultXmlDocument.LastChild.ChildNodes` can be rewritten with XPath to make navigation through the document clearer.

## Requesting by currency code

Once the request and the response handling work, let us rewrite the function to request a rate by currency code. We add a dictionary of code mappings and build the request address from the start and end dates and the currency code:

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

I also changed `CurrencyCode` when creating a record in the loop: instead of `"R01235"` the record now holds the familiar currency code `"USD"`.

`RequestGetXml` is a helper function that loads an XML document with an HTTP request (the address is now passed as an argument):

    Private Function RequestGetXml(requestUrl As String) As DOMDocument60
        'about requests in VBA: https://codingislove.com/http-requests-excel-vba/
        Dim XMLHTTP As New MSXML2.XMLHTTP60

        XMLHTTP.Open "GET", requestUrl, False
        XMLHTTP.send
        Set RequestGetXml = XMLHTTP.responseXML
    End Function

Let us check that the whole function works as expected by adding a temporary test procedure:

    Private Sub TestCurrLoad()
        Dim col As Collection
        Set col = GetCurrency("USD", #1/1/2020#, #2/2/2020#)
        Stop
    End Sub

## A worksheet function

Now let us add a public function that can be used in worksheet formulas (a UDF, user-defined function):

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

The `DateAdd` function sets the date range to load: from yesterday to today. We get two records and take the value from the last one (VBA collections are numbered from 1).

Note the second argument, `volatileArg`. Its value is not used, but it is there to force Excel to recalculate the result when that argument changes in the cell formula:

    =GetTodayCurrency("USD", NOW())

To confirm that the rate was refreshed, the function also calls `Debug.Print` with a success message. After debugging, that call can be removed so it does not clutter the Immediate window.

## Checking it works

Let us check that the rate is loaded when the workbook is opened:

<img src="/images/articles/currency-rates-from-cbr-vba/1590231043185812459.webp" loading="lazy" width="700" height="346" alt="The Immediate window with the &quot;currency loaded&quot; message" />

The Immediate window shows the success message right after the workbook is opened. The same message will appear after any change on the sheet, because Excel recalculates the formula. If the workbook has many cells calling this function, it is better to add a cache: for example, a static variable holding a `Dictionary` where the key is a combination of the currency code and the date, and the value is the `CurrencyRecord`. That avoids a lot of identical requests to the CBR website.

Improvements can be suggested in the issues of the repository <a href="https://github.com/navferty/CBR-VBA-Currencies" rel="noopener" target="_blank">https://github.com/navferty/CBR-VBA-Currencies</a>.

Originally published in Russian on <a href="https://pikabu.ru/story/kursyi_valyut_s_sayta_tsb_na_vba_s_avtoobnovleniem_7469372" rel="noopener" target="_blank">Pikabu</a>.
