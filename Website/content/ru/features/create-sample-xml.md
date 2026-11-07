{
  "title": "Создать образец XML по XSD",
  "description": "Генерация образца XML-файла из XSD-схемы для визуализации структуры.",
  "image": "https://navferty.net/images/CreateSampleXml1.png"
}

[Главная](/ru/) › [Команды](/ru/#commands) › Создать образец XML по XSD

# <img alt="" src="/images/icons/createSampleXml.png" class="icons" width="64" height="64" />Создать образец XML по XSD

Генерация образца XML-файла из XSD-схемы для визуализации структуры.

<div class="shots">

<img src="/images/CreateSampleXml1.png" loading="lazy" width="497" height="441" alt="Созданный пример XML-файла в редакторе" />

</div>

## Посмотреть пример

**Входная XSD-схема:**

    <xsd:schema xmlns:xsd="http://www.w3.org/2001/XMLSchema"
               xmlns:tns="http://tempuri.org/PurchaseOrderSchema.xsd"
               targetNamespace="http://tempuri.org/PurchaseOrderSchema.xsd"
               elementFormDefault="qualified">
     <xsd:element name="PurchaseOrder" type="tns:PurchaseOrderType"/>
     <xsd:complexType name="PurchaseOrderType">
      <xsd:sequence>
       <xsd:element name="ShipTo" type="tns:USAddress" maxOccurs="2"/>
       <xsd:element name="BillTo" type="tns:USAddress"/>
      </xsd:sequence>
      <xsd:attribute name="OrderDate" type="xsd:date"/>
     </xsd:complexType>

     <xsd:complexType name="USAddress">
      <xsd:sequence>
       <xsd:element name="name"   type="xsd:string"/>
       <xsd:element name="street" type="xsd:string"/>
       <xsd:element name="city"   type="xsd:string"/>
       <xsd:element name="state"  type="xsd:string"/>
       <xsd:element name="zip"    type="xsd:integer"/>
      </xsd:sequence>
      <xsd:attribute name="country" type="xsd:NMTOKEN" fixed="US"/>
     </xsd:complexType>
    </xsd:schema>

**Сгенерированный XML:**

    <PurchaseOrder xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" OrderDate="1900-01-01" xmlns="http://tempuri.org/PurchaseOrderSchema.xsd">
      <ShipTo country="US">
        <name>name1</name>
        <street>street1</street>
        <city>city1</city>
        <state>state1</state>
        <zip>1</zip>
      </ShipTo>
      <ShipTo country="US">
        <name>name2</name>
        <street>street2</street>
        <city>city2</city>
        <state>state2</state>
        <zip>-79228162514264337593543950335</zip>
      </ShipTo>
      <BillTo country="US">
        <name>name1</name>
        <street>street1</street>
        <city>city1</city>
        <state>state1</state>
        <zip>1</zip>
      </BillTo>
    </PurchaseOrder>

<div class="info-box">

**Статья по теме:** [Проверка XML по XSD-схеме в Excel](/ru/articles/validate-xml-against-xsd-in-excel/)

</div>

Надстройка бесплатная, с открытым кодом. [Инструкция по установке](/ru/install/).
