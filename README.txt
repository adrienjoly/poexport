POExport (Pocket Outlook Export)
 | Adrien Joly
 | Version 0.1
 
https://developer.berlios.de/projects/poexport/


___________
Description
===========

POExport aims to be a simple solution to extract information from Pocket Outlook databases (contacts, schedule, tasks...) into standard formats like CSV or XML. It's based on .net framework and Pocket Outlook Object Model (POOM) wrapper.

So far, POExport is a script-like program that exports all the pocket outlook contacts from a PocketPC to a CSV file, following the structure defined in a custom CSV template. The default CSV template follows gmail contacts format, making it possible to import your contacts from your PocketPC straight to your gmail account.

At this stage, the application is not advanced at all but it's a very convenient tool that can be used for one-way contacts synchronization (or backup) from a PocketPC.

_________
Licensing
=========

Because it's based on Microsoft's POOM Wrapper Sample, POExport is under "Microsoft Shared Source" license. Please read and agree with the terms of this license provided in the "EULA.rtf" file.

Basically, the program is free of use and the source code is shared in a limited manner. Please refer to the license.

____________
Installation
============

Prerequisites:
- Microsoft .net Compact Framework Runtime version 1
- Microsoft POOM Wrapper (PocketOutlook.dll, platform-specific library)

Procedure:
- copy the binary file "poexport.exe" to your PPC, in the same directory as "PocketOutlook.dll"
- copy the template file "csvtemplate.csv" to your PPC, in the root directory

(?) How to install the "POOM Wrapper" (PocketOutlook.dll):
You can find it on Microsoft website. Search on google or go directly to this URL to download it:
-> http://www.microsoft.com/downloads/details.aspx?FamilyId=80D3D611-CC81-4190-AAB4-B1EA57637BAC&displaylang=en
Then follow the instructions provided to install it.

_____
Usage
=====

Run "poexport.exe". After a few seconds of processing, you'll get a "contacts.csv" file in the root directory of your PocketPC.

If you get an error, follow the installation instructions above. Make sure the POOM wrapper library is compatible with your plateform.

Note that POExport has no GUI, it's a script-like program, so don't expect to see any message on the screen. The "\contacts.csv" file will be overwritten after every execution.

____________________________
Customizing the CSV template
============================

The template file defines the mapping between POOM field names and the column header labels that will be inserted in the exported CSV contacts file.

Its structure is very simple: each line defines a mapping using the syntax: "FieldName","My corresponding column header"
Here is an example: "Email1Address","E-mail address #1"
Only the mapped fields will appear as columns in the exported CSV file.
The default template defines mappings for every field, and is compatible with gmail contact import.

List of recognized field names:
"FileAs", "Title", "FirstName", "MiddleName", "LastName", "Email1Address", "Email2Address", "Email3Address", "WebPage", "MobileTelephoneNumber", "HomeTelephoneNumber", "Home2TelephoneNumber", "HomeFaxNumber", "HomeAddressStreet", "HomeAddressCity", "HomeAddressState", "HomeAddressPostalCode", "HomeAddressCountry", "JobTitle", "CompanyName", "Department", "OfficeLocation", "AssistantName", "PagerNumber", "BusinessTelephoneNumber", "Business2TelephoneNumber", "BusinessFaxNumber", "BusinessAddressStreet", "BusinessAddressCity", "BusinessAddressState", "BusinessAddressPostalCode", "BusinessAddressCountry", "CarTelephoneNumber", "RadioTelephoneNumber", "OtherAddressStreet", "OtherAddressCity", "OtherAddressState", "OtherAddressPostalCode", "OtherAddressCountry", "Suffix", "Spouse", "Children", "Categories", "Body"

___________
Source code
===========

The POExport application consists of a managed client using a native wrapper to the Pocket Outlook Object Model (POOM). Only the source code of the client is provided. The source code for the wrapper is distributed freely on Microsoft website, under the name "POOM Wrapper Sample".

The Visual Studio .NET project containing the source code does not include the compiled native wrapper dll (since this native wrapper is processor specific). Therefore, simply deploying the project "as is" will result in an error when the application is launched. In order to get the project to deploy and launch correctly you will have to compile the POOM Wrapper native source code (not included with POExport) for the intended processor (e.g. use eMbedded Visual Tools 3.0) and include the dll in the project so it deploys with the application. You can also change the default deployment directory from "\Windows" to "\Program Files\SDE POOM Sample". This can be done by right clicking on the Project file in the Solution Explorer of Visual Studio .NET and selecting "Properties". Choose "Device Extensions" under "Common Properties" and change the Output File Folder.

______________________
Contribution & Contact
======================

If you want to contribute to this project, contact me from the project website:
https://developer.berlios.de/projects/poexport/

___________________
Changelog / History
===================

2006/01/25 - V0.1
 - first public release
 - PO contacts are exported in a CSV file, following a CSV template