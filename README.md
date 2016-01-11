QuetzalXLSReader
================

QZXLSReader is a library that makes it very easy to extract information from an XLS file. It is a very simple library that is based off libxls and DHlibXLS. It may not be as complete as either, but it can get the job done in a very simple manner.

Note: I haven't tried it, but I believe it cannot open XLSX or XLSM files.

How to use
----------

There are two ways to use QZXLSReader:

1) The default way:

- Drag and drop the .framework.
- Add libiconv.dylib to your project.

2) The custom way:

- Change anything you want from the project.
- Compile this project.
- Fix any warnings that might come up (I had to do some dancing with the structs to make it work without exposing libxls).
- Add libiconv.dylib to your project.

Once you have the framework in your project and it's building correctly, you can start by calling:

    QZWorkbook *excelReader = [[QZWorkbook alloc] initWithContentsOfXLS:xlsFileURL];

This loads the workbook into memory. Now all you have to do is open any worksheet and start browsing the contents.

```
QZWorkSheet *firstWorkSheet = excelReader.workSheets.firstObject;
[firstWorkSheet open];

NSLog(@"%@", firstWorkSheet.rows.firstObject)
```

Once you're done, you can close either the worksheet you were working on or the whole workbook.

License
-------
TL; DR: I don't discriminate software, so feel free to use it anywhere as long as you credit Quetzal (or Fernando Olivares) appropriately.

Copyright © 2015, Fernando Olivares Lozano
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:
1. Redistributions of source code must retain the above copyright
   notice, this list of conditions and the following disclaimer.
2. Redistributions in binary form must reproduce the above copyright
   notice, this list of conditions and the following disclaimer in the
   documentation and/or other materials provided with the distribution.
3. All advertising materials mentioning features or use of this software
   must display the following acknowledgement:
   This product includes software developed by Fernando Olivares.
4. Neither the name of the <organization> nor the
   names of its contributors may be used to endorse or promote products
   derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY <COPYRIGHT HOLDER> ''AS IS'' AND ANY
EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

Non-Attribution License
---------

If you want a non-attribution license, get in touch with us! The email is at the end of this readme.


Conclusion
----------

If you find this software lacking, feel free to fork it and send us a pull request.

If you do make an app that uses it, let us know!

Quetzal is dedicated to making great iOS apps, so if you like our code, feel free to contact us! We do consulting work as a third party or we can build the app of your dreams!

FAQ
---

I'm running into an error in QZCell.h. How do I fix that?
- Short answer:
 
Change:

    - (id)initWithContent:(xlsCell *)cell;

to

    - (id)initWithContent:(struct xlsCell *)cell;

- Long answer: 

I'm not sure what's going on, since the compiler complains only with a newly built library (and it doesn't complain while building). Once built, it complains that it doesn't know what an xlsCell is, so you have to define it (again) as a struct.
