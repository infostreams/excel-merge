Excel Merge
===========

Merges two or more Excel files into one file, while keeping formatting, formulas, VBA code and 
conditional styling intact. This software works with Excel 2007 (.xlsx and .xlsm) files and can 
only generate Excel 2007 files as output. The older .xls format is unfortunately not supported, 
but you can work around that if necessary. 

This is a software library that is designed to be used as part of a larger piece of software. It 
cannot be used as standalone software by itself.

Installation
------------

**With composer**

    php composer.phar require infostreams/excel-merge

Use
---

Provided that you have already included the Composer autoloader (```vendor/autoload.php```) you can 
simply do 

    <?php
      $files = array("sheet_with_vba_code.xlsm", "generated_file.xlsx", "tmp/third_file.xlsx");
      
      $merged = new ExcelMerge\ExcelMerge($files);            
      $merged->download("my-filename.xlsm");
      
      // or
      
      $filename = $merged->save("my-directory/my-filename.xlsm");
    ?>


Raison d'être and use case
--------------------------
This library exists for one reason only: to work around the humongous memory requirements of the
otherwise excellent [PHPExcel](https://github.com/PHPOffice/PHPExcel) library. I had to export
the contents of a database as an Excel file with about 10 worksheets, some of them relatively 
large, and PHPExcel very quickly ran out of memory after producing about 2 or 3 of the required 
worksheets, even after increasing the PHP memory limit to 256 and then 512 Mb. I was not doing 
anything spectacular and am certainly 
[not the only one](http://stackoverflow.com/questions/4817651/phpexcel-runs-out-of-256-512-and-also-1024mb-of-ram) 
to have run into this issue.

At this point I could have chosen a different Excel library to generate the export, and 
[I did](https://github.com/MAXakaWIZARD/xls-writer), but these would not allow me to use VBA code
in my exported file, and would not recognize some of the Excel formulas I needed. PHPExcel would
allow me to do these things, but ran out of memory because it insists on keeping a complete mental
model of all the sheets in memory before it could produce an output file. That makes sense for 
PHPExcel but doesn't work for my use case.

Therefore, I decided to circumvent PHPExcel's memory limitations by using it to generate and then 
write all sheets as **individual Excel files**, and then write some code to merge these Excel
files into one.

How it works
------------
Instead of trying to keep a mental model of the whole Excel file in memory, this library simply 
operates directly on the XML files that are inside Excel2007 files. The library doesn't 
really understand these XML files, it just knows which files it needs to copy where and how to
modify the XML in order to add one sheet of one Excel file to the other. 

This means that the most memory it will ever use is directly related to how large your largest
worksheet is. 

Results
-------
I had to generate an Excel file with 11 relatively sizable worksheets (two or three sheets with 
about 2000 rows). PHPExcel took over 30 minutes and over 512 Mb of memory to generate this, after 
which I aborted the process. With this library, I can generate the same export in 28.2 seconds with 
a peak memory use of 67 Mb.

Support for 'native' Excel files
--------------------------------
I've tried merging files produced by Excel itself, but somehow it fails. I worked around it by
loading the file with PHPExcel and writing it as a new Excel2007 file, and then merging that 
instead. If you figure out why it fails: pull requests welcome.

Support for .xls files and Libre/OpenOffice Calc and Gnumeric
-------------------------------------------------------------
You can merge .xls files, or any of the import formats supported by PHPExcel, by reading the 
file with PHPExcel and writing it as a temporary Excel2007 file. You then merge the temporary 
Excel2007 file instead of the original file

Requirements
------------
This library uses DOMDocument and DOMXPath extensively. These are installed and available in PHP5 by 
default. If they aren't, check [here](http://php.net/manual/en/dom.setup.php).

Minimum PHP version is most likely v5.3.