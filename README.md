# WSO2 / Axis2 Excel Message Builder

## Description
The purpose of this project is to build a custom message builder in order to treat xls document. In a first version it will only convert the content of the document in csv and put it as a text in the payload

## Dependency
In order to read and transform xls files we are using [Apache POI](https://poi.apache.org/)

## Build
Build is performed via maven so simply run the following command to build the jar

```
mvn clean instal
```

## Deployment
To use this new builder you have to copy the library (poi-ooxml) and the generated jar to WSO2 component directory  *{WSO2_HOME}/repository/components/lib*
Then you need to setup axis to use this message builder for a give content type

```xml
<messageBuilder class="org.wso2.custom.ExcelMessageBuilder" contentType="application/vnd.ms-excel"/> 
```
