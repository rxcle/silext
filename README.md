# silext
Silverlight installer extractor

Extracts files from the Silverlight installer, allowing Silverlight to be used directly via COM and NPAPI without installing.


Usage: Silext <Silverlight_x64.exe> <target_path> [<options>]

Options: "s" Only extract 64-bit program files (otherwise extract everything)

Returns:  0 Success
         >0 Success with warning (e.g. no cleanup)
         <0 Fatal error 

Silext is Copyright (c) 2020 Rxcle. All rights reserved.

Individual redistribution or repackaging without explicit permission is not permitted.
This program is provided WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.


This is a tool FOR Microsoft Silverlight 5, but contains no direct references to or 
dependencies on Microsoft Silverlight 5 itself.

Microsoft Silverlight is Copyright (c) 2016 Microsoft Corporation. All rights reserved.


This application makes use of the following third-party open source components:

    [bit7z] Copyright (C) 2014 - 2019 Riccardo Ostani. (GNU GPL v2)
    [7-Zip] Copyright (C) 1999-2018 Igor Pavlov. (GNU LGPL)
