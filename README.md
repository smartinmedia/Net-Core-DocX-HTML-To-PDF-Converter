# Report-From-DocX-HTML-To-PDF-Converter - Create custom reports based on Word docx or HTML documents and convert to PDF with .NET CORE

**Still under development... don't clone yet...**

## Donate Now!
We are a company (Smart In Media GmbH & Co. KG, https://www.smartinmedia.com) and have to pay salaries to our developers. As we also believe in free software, we give away this under the MIT license and you can do whatever you want with it. Your are not bound to making continuous paments, etc. So, if you saved money and time by using this library, please donate to us via Paypal!

## Background/Why?
I was working on a medical project (https://www.easyradiology.net) and wanted that users receive a dynamically created PDF document via e-mail. Can't be too difficult, I thought. There'll be some library to do that. I discovered that in fact, there is PDF Sharp, an open source library. However, I found out that you have to "code" each line of the PDF yourself. What happens, if your sales department or user support wants to change just one line here or update a logo there in the document? Then, you have to go back to your code, change it, talk again to the sales dept or user support, if they are happy, etc until you can create a new release. This is super tedious. So I thought, would be great, if the sales department / user support can just create a simple Word docx or a HTML document with some placeholders and as a developer, you can just paste your dynamic content. 
Using Word docx - a Microsoft product - in .NET and adding stuff to it and then converting to PDF cannot be too difficult, right? Unfortunately, I was surprised, just how difficult that endeavour is if you don't want to pay huge amounts of money to commercial libraries. The cheapest started at U$300 and the most expensive were around U$5,000. So, I started deloping a solution myself. It is not perfect, but it seems to work...
Again, if you like it, please donate money or code!

## What's the functionality?

Report from DOCX / HTML to PDF Converter can parse the source document and introduce the dynamic content into predefined __placeholders__. Then it can perform the following conversions:

* DOCX to DOCX 
* DOCX to PDF
* DOCX to HTML
* HTML to HTML
* HTML to DOCX
* HTML to PDF

## What do you need to get started?
Don't get scared away that I use LibreOffice, it is easier than you may think!
1. LibreOffice - just get the PORTABLE EDITION as you don't screw up your webserver with an installation. The portable version just runs without any installation. We need LibreOffice for converting from DOCX or from HTML to PDF and DOCX, etc

2. Nuget: 
* Microsoft.NetCore.App
* Document.Format.OpenXml
* System.Drawing.Common

3. The OpenXml PowerTools (thanks to Eric White for this great work), which I already included into the project, the whole code!


```csharp
public static void test(){

}
```
Another test
