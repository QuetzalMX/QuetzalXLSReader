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

The license to use the framework is included in QZXLSReader.h; it's basically a BSD license with attribution.

We don't discriminate software, so feel free to use it anywhere as long as you credit Quetzal (or Fernando Olivares) appropriately.

Non-Attribution License
---------

If you want a non-attribution license, get in touch with us! The email is at the end of this readme.


Conclusion
----------

If you find this software lacking, feel free to fork it and send us a pull request.

If you do make an app that uses it, let us know!

Quetzal is dedicated to making great iOS apps, so if you like our code, feel free to contact us! We do consulting work as a third party or we can build the app of your dreams!

Contact: olivaresf@quetzal.li

FAQ
---

I'm running into an error in QZCell.h. How do I fix that?
- Short answer:
 
Change:

```- (id)initWithContent:(xlsCell *)cell;

to

```- (id)initWithContent:(struct xlsCell *)cell;

- Long answer: I'm not sure what's going on, since the compiler complains only with a newly built library (and it doesn't complain while building). Once built, it complains that it doesn't know what an xlsCell is, so you have to define it (again) as a struct.
