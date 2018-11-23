# Convert HTML to DOCX document
# Author Amal Amrani 
# Author URI: [Amal Amrani Website link](http://amrani.es/ "amrani.es")

**convert_html_to_docx allows you to convert a HTML Document To Word .docx. The different elements supported and parsed by this library are the following:**


Bold <strong>
Cursive <em>
Paragraph <p>
Section  <div>
Lists nested with bullets or numbered <ul> / <ol>
Tables <table>
Links <a href >
Quotes <blockquote>
Images <img>

The images must be in the same location as the HTML document in order to be accessed and converted.

The images URL in the HTML document are converted first to Base64, before converting the document to .docx.
The generated documents support RTL orientation for texts that require writing from right to left.


To use this converter, you only need to include the following files that are already included in the .zip to download:


To save the .docx:

Stuk-jszip/dist/jszip.min.js 
Stuk-jszip/vendor/FileSaver.js

To create the different documents that integrate and define the .docx structure:

convert_html_to_docx.js

You can see an example here [example to convert html to docx link](http://amrani.es/proyectos/conversor-HTML-DOCX/convert-html-to-docx.html "convert html document to .docx")

