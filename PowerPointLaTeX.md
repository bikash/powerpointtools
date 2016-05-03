# Introduction #

PowerPointLaTeX is a small add-in that allows you to embed LaTeX math code into your PowerPoint 2007 presentations in a user-friendly way.
It uses a web-service to render the LaTeX code (currently it uses the equation editor from http://www.codecogs.com/). Support for other LaTeX web-services is planned.

From version 0.7 on, MiKTeX is supported and also the recommended LaTeX compilation service.
It supports better aligning of inline formulas and flexible rendering resolution (using the preview package in LaTeX and dvi2png).

**NOTE:** You might have to install the preview package manually in MiKTeX. So if the MiKTeX service doesn't work, you might want to check if the preview package is installed. Actually it should be installed automatically if it is not found the first time when you use PowerPointLaTeX.

# Details #

PowerPointLaTeX adds a new ribbon to PowerPoint 2007:

![http://blog.blackhc.net/wp-content/uploads/2009/04/pptlatex1.jpg](http://blog.blackhc.net/wp-content/uploads/2009/04/pptlatex1.jpg)

There are two different types of equations: inline equations and formulas.
Inline equations can be added anywhere in-between normal text. Simply wrap your LaTeX math code in a pair of $$: "Some integral $$ \int x \, dx = {1 \over  2} x^2 + C$$".

<wiki:gadget url="http://hosting.gmodules.com/ig/gadgets/file/106581606564100174314/vimeo-video.xml" up\_videourl="http://vimeo.com/4442353" width="480" height="420" border="0"/>

Formulas are stand-alone objects. You can create them by using the Ribbon button "New Formula". A dialog will open up and allow you to enter a LaTeX formula. It will automatically create a preview using the normal webservice.

To edit a formula click on it and then on "Edit Formula" in the ribbon.

http://blog.blackhc.net/wp-content/uploads/2009/08/PPTLaTeX_eqeditor.JPG


PowerPointLaTeX caches the compiled formulas, so you don't need an active internet connection if you don't change anything (you don't need one either if you are using MiKTeX for rendering).

# MiKTeX Features #
It is possible to specify some LaTeX code that is included in every equation. You can set this "preamble" in the preferences dialog.

The MiKTeX service supports base-line aligning of equations using the preview package in LaTeX:

![http://powerpointtools.googlecode.com/svn/trunk/PowerPointLaTeX/Examples/BaselineText.jpg](http://powerpointtools.googlecode.com/svn/trunk/PowerPointLaTeX/Examples/BaselineText.jpg)

![http://powerpointtools.googlecode.com/svn/trunk/PowerPointLaTeX/Examples/BaselineResult.jpg](http://powerpointtools.googlecode.com/svn/trunk/PowerPointLaTeX/Examples/BaselineResult.jpg)

# Known Issues #

If you have huge LaTeX formulas PowerPoint's Auto-Fit feature might kick in. It's toxic to PowerPointLaTeX's inline formula embedding feature - the formulas will end up in the wrong place. Simply disable Auto-Fit for those text shapes and adapt the font size manually if necessary.

The code internally uses the clipboard to insert the formulas into the presentation and for some reason the original clipboard content is lost even though that should not happen.