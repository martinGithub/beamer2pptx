# beamer2pptx
A python script to convert Beamer (latex) presentations to PowerPoint presentations

## Goal
the goal of this script is to help automatize the process of converting  a Beamer Presentation (Latex) into a microsoft powerpoint.
It parses the file given in the variables texfile using regular expressions and creates a powerpoint document with the same number of slides, with the slides titles, the images in each slides (but does not keep the size or layaout yet) and creates an image for each equation that are also added to the slides

## Usage

replace the right part in line texfile='example.tex' using the name of you latex file
and then run the script using python

## Limitations

It is very basic for now, it doesn't even recreate the bullet points. But is it still usefull as a first step to speed up manual conversion.


The main problem is that it uses regular expressions to parse the latex document. 
Parsing the latex robustly with regular expression seems difficult in case of nested brackets etc. 
We should instead use a actual parser like [plasTeX](http://plastex.sourceforge.net/) or maybe convert the document to xml using [LaTexXML](http://dlmf.nist.gov/LaTeXML/get.html ) for example and then use a xml python parser. Or we could first convert to html using the command htlatex or tex4ht and then parse the html but on my machine the conversion does not seem to work as it converts only the first slide.

the latex equations are converted into images and thus cannot be edited. It would be good to be able to keep them as editable equation in powerpoint using some latex plugin for powerpoint, or convert them into powerpoint equations.

## Alternatives 

You can compile the beamer into a pdf file and then convert the pdf to powerpoint using [pdf2pptx](https://github.com/ashafaei/pdf2pptx).
This works by converting the pdf into a set of images (one image for each slide) and then make a powerpoint from these images.
An obvious limitation of that approach is that you cannot reedit the slides in powerpoint and change the theme.


You can compile the beamer into a pdf file and then open the pdf in libreOffice (tested with version 4.2.8.2). The text might not be well preserved if you use some latin encoding for example (you can many strange characters), but the positions of the images are good. However the latex equations won't be well preserved. You might then be able to export into an ODF that you can import in powerpoint.


If you are interested in converting powerpoints files to beamer you can have a look at [pptx2beamer](https://github.com/IngoScholtes/pptx2beamer)
which uses the .Net Microsoft framework and requires Microsoft Visual Studio for the compilation.




