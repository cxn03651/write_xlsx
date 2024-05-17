---
layout: default
title: LINITATION
---
### <a name="limitation" class="anchor" href="#limitation"><span class="octicon octicon-link" /></a>LIMITATION

The following limits are imposed by Excel 2007+:

    Description                                Limit
    --------------------------------------     ------
    Maximum number of chars in a string        32,767
    Maximum number of columns                  16,384
    Maximum number of rows                     1,048,576
    Maximum chars in a sheet name              31
    Maximum chars in a header/footer           254

    Maximum characters in hyperlink url (1)    2079
    Maximum number of unique hyperlinks (2)    65,530

(1) Versions of Excel prior to Excel 2015 limited hyperlink links and anchor/locations to 255 characters each. Versions after that support urls up to 2079 characters. Excel::Writer::XLSX versions >= 1.0.2 support the new longer limit by default.

(2) Per worksheet. Excel allows a greater number of non-unique hyperlinks if they are contiguous and can be grouped into a single range. This isn't supported by Excel::Writer::XLSX.


[WriteXLSX]: index.html
