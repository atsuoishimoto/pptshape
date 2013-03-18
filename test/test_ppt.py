import os, sys
from unittest import mock
import pptshape.ppt
import pptshape.directive

TESTDIR = os.path.split(__file__)[0]

def test_open():
    ppt = pptshape.ppt.PPTShape(os.path.join(TESTDIR, 'testppt.pptx'))
    shape = ppt.saveShape('shape-title', 
                os.path.join(TESTDIR, 'aaa.png'))
    ppt.quit()

def test_update():
    NAME_PNG = 'abc.png'
    path_png = os.path.join(TESTDIR, NAME_PNG)

    NAME_PPT = 'testppt.pptx'
    path_ppt = os.path.join(TESTDIR, NAME_PPT)

    directive = pptshape.directive.PPTShape(None, None, None, 
             None, None, None, None, None, None)
    if os.path.exists(path_png):
        os.unlink(path_png)

    directive.arguments = [NAME_PNG]
    directive.state = mock.Mock()
    directive.state.document.current_source = os.path.join(TESTDIR, 'test.rst')
    directive.options = {
        'pptfilename': NAME_PPT,
        'shapename': 'shape-title',
    }
    directive.run()
    assert os.path.exists(path_png)
    
    os.utime(path_png, (0, 0))
    directive.run()

    assert os.stat(path_png).st_mtime != 0

    # Make png file empty.
    open(path_png, 'w').close()

    directive.run()
    assert os.path.getsize(path_png) == 0
