Open-XML-PowerTools
===================

The Open XML PowerTools provides guidance and example code for programming with Open XML
Documents (DOCX, XLSX, and PPTX).  It is based on, and extends the functionality
of the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK).

InsideTravel Technology have forked the repository to add additional functionality to the DocumentAssembler class, namely;

- Ability to insert Documents or other Document Templates during document assembly
- Ability to insert Images during document assembly

The original [Open XML Powertools] project can be found here: https://github.com/OfficeDev/Open-Xml-PowerTools.

Copyright (c) Microsoft Corporation 2012-2017
Portions Copyright (c) Eric White 2016-2017
Licensed under the Microsoft Public License.
See License.txt in the project root for license information.

Open-Xml-PowerTools Content
===========================

There is a lot of content about Open-Xml-PowerTools at the [Open-Xml-PowerTools Resource Center at OpenXmlDeveloper.org](http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx)

See:
- [DocumentBuilder Resource Center](http://openxmldeveloper.org/wiki/w/wiki/documentbuilder.aspx)
- [PresentationBuilder Resource Center](http://openxmldeveloper.org/wiki/w/wiki/presentationbuilder.aspx)
- [HtmlConverter Resource Center](http://openxmldeveloper.org/wiki/w/wiki/htmlconverter.aspx)
- [Introduction to DocumentAssembler](https://www.youtube.com/watch?v=9QqzCgfqA2Y)
- [Contributing to Open-Xml-PowerTools via GitHub](https://www.youtube.com/watch?v=Ii7z9L6Dkko)
- [Gitting, Building, and Installing Open-Xml-PowerTools](https://www.youtube.com/watch?v=60w-yPDSQD0)

Build Instructions
==================

Recently, we've updated the GitHub repo so that it pulls the Open-Xml-Sdk via NuGet.  The video at the following link shows how to clone and build the Open-Xml-PowerTools
when pulling the Open-Xml-Sdk via NuGet.  It uses Visual Studio 2017 Community Edition.

[Build Open-Xml-PowerTools](http://ericwhite.com/blog/2017/03/24/building-open-xml-powertools-when-pulling-the-open-xml-sdk-via-nuget/)

Procedures for enhancing Open-Xml-PowerTools
--------------------------------------------
There are a variety of things to do when adding a new CmdLet to Open-Xml-PowerTools:
- Write the new CmdLet.  Put it in the Cmdlets directory
- Modify Open-Xml-PowerTools.psm1
  - Call the new Cmdlet script to make the function available
  - Modify Export-ModuleMember function to export the Cmdlet and any aliases
- Update Readme.txt, describing the enhancement
- Add a new test to Test-OpenXmlPowerToolsCmdlets.ps1

Procedures for enhancing the core C# modules
- Modify the code
- Write xUnit tests
- Write an example if necessary
- Run xUnit tests on VS2015 Community Edition
- Run xUnit tests on VS2013 Update 4
