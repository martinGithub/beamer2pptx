# beamer2pptx
python script to convert beamer (latex) presentations to PowerPoint presentations

## Goal
the goal of this script is to help automatize the process of converting  a Beamer Presentation (Latex) into a microsoft powerpoint.
It parses the file given in the variables texfile using regular expressions and creates a powerpoint document with the same number of slides, with the slides titles , the images in each slides (but does not keep the size or layaout yet) and creates an image for each equation that are also added to the slides

## Usage

replace the right part in line texfile='example.tex' using the name of you latex file
and then run the script using python

## Limitations

It is very basic for now, it doesn't even recreate the bullet points. But is it still usefull as a first step to speed up manual conversion.
The main problem is that it uses regular expression to parse the latex document. 
Parsing the latex robustly with regular expression seems difficult in case of nested brackets etc we should use a actual parser or maybe convert the document to xml using LaTexXML and then use a xml parser ?  http://dlmf.nist.gov/LaTeXML/get.html 
we can alose parse the latex using  plasTeX 
We could first convert to html using the command htlatex or tex4ht and then parse the html but on my machine the conversion does not seem to work, only the first slide is converted into the html
plasTeX fails parsing when the latex uses \usepackage{mathtools}

## Alternatives 

you can convert compile the beamer into a pdf file and then convert the pdf to powerpoint using [pdf2pptx](https://github.com/ashafaei/pdf2pptx).
This works by converting the pdf into a set of images (one image for each slide) and then make a powerpoint from these images.
An obvious limitation of that approach is that you cannot reedit the slides in powerpoint and change the theme.
If you are interested in converting powerpoints to beamer you can have a look at [pptx2beamer](https://github.com/IngoScholtes/pptx2beamer)
which used the .Net framework and requires Visual Studio for the compilation.




