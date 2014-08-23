pptshape
============================

Pptshape extracts shapes from ppt file and embeds to the sphinx document as png file.

pptshape adds new directive as follow.

::

   .. ppt-shape:: abc.png
      :pptfilename: testppt.pptx
      :shapename: shape-title

First line of the directive specifies filename of image file to be generated. `:pptfilename:` specifies 
name of PowerPoint presentaion. `:shapename:` specifies name of shape you specified(see Usage).

When you build sphinx project on Windows box with PowerPoint installed, pptshape opens ppt file 
and create png file if ppt file is newer than image files.

If the project is build on PC PowerPoint is not installed or non-Windows box, ppt-shape directive 
behave just like as standard image directive.

The ppt-shape directive is derived from standard image directive, so you can use directives such as 
`:height:` or `:alt:` image directive has.


Usage
--------------------

1. Install `pywin32 <http://sourceforge.net/projects/pywin32/>`_ package to your Windows box.

2. Create new presentation(.ppt) and draw shape.

3. Select the shape you wrote and display format tab. Specify shape of name the shape at 'title' field.

4. Save presentation file.

5. In `conf.py` file of your Sphinx project, add following configuration.

   .. code-block:: python

      extensions = ['pptshape.directive']

6. Add following lines in your .rst files.

   ::

      .. ppt-shape:: abc.png
         :pptfilename: testppt.pptx
         :shapename: shape-title


   `abc.png` is a name of png file to be created. `:pptfilename:` specifies name of PowerPoint presentaion. `:shapename:` specifies name of shape you specified at step 3.

7. Build sphinx project.


Special shape name
------------------

Shapename in the pptshape directive starts with `#` specifies
position in the PowerPoint presentation rather than title of the shape.

1. `#m.n` repesents `n` th shape in the `m` th slide.

2. `#n` represents `n` th shape through entire the ppt presentation.

Both `n` and `m` should be digit indexes pages and shapes from 1 and up.
For example, `#1` is for first shape of the document,
and `#2.1` is for first shape in the second slide.

This way is usefull for the document which cannot be modified to add
shape title, or for old power point (2007 or before) which does not
support shape title.


Requirements
============

* Python 2.7/3.3 or later

* pywin32 to generate png file.

Copyright 
=========================

Copyright (c) 2013, 2014 Atsuo Ishimoto

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
