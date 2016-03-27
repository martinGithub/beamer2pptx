# beamer2pptx
python script to convert beamer (latex) presentations to PowerPoint presentations

# GOAL
 	the goal of this script is to help automatize the process of converting 
	  a Beamer Presentation (Latex) into a microsoft powerpoint
 	it parses the file given in the variables texfile using regular
	  expressions and creates a powerpoint document with the
 	same number of slides, with the slides titles , the images in each slides (but does not keep the size or layaout yet)
 	and creates an image for each equation that are also added to the slides

# USAGE
 	replace the right part in line texfile='presentation_test.tex' using the name of you latex file

# LIMITATIONS
 	parsing the latex robustly with regular expression seems difficult in case of nested brackets etc
 	we should use a actual parser or maybe convert the document to xml using LaTexXML ?  http://dlmf.nist.gov/LaTeXML/get.html 
 	could first convert to html using the commande htlatex and then parse the html# 
 	but on my machine i get errors when running htlatex (Can't find/open #	file `mathkerncmssi8.tfm')
 	i get the same error using tex4ht 
	  i did not get the time to pin down the problem 
 	plasTeX fails parsing when the latex uses 
 		-\usepackage{mathtools}
  		-french accents 
