import os, sys
try:
    from unittest import mock
except ImportError:
    import mock

import pptshape.directive
try:
    from pptshape import ppt
except ImportError:
    ppt = None

TESTDIR = os.path.split(__file__)[0]

def test_open():
    if not ppt:
        return

    p = ppt.PPTShape(os.path.join(TESTDIR, 'testppt.pptx'))
    shape = p.saveShape('assign_list', 
                os.path.join(TESTDIR, 'assign.png'))

    p.quit()

def test_update():
    if not ppt:
        return

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


def test_index():
    if not ppt:
        return

    NAME_PPT = 'testppt.pptx'
    path_ppt = os.path.join(TESTDIR, NAME_PPT)
    NAME_PNG = 'abc.png'
    path_png = os.path.join(TESTDIR, NAME_PNG)

    directive = pptshape.directive.PPTShape(None, None, None, 
             None, None, None, None, None, None)

    if os.path.exists(path_png):
        os.unlink(path_png)

    directive.arguments = [NAME_PNG]
    directive.state = mock.Mock()
    directive.state.document.current_source = os.path.join(TESTDIR, 'test.rst')
    directive.options = {
        'pptfilename': NAME_PPT,
        'shapename': '#1',
    }
    directive.run()
    assert os.path.exists(path_png)

def test_page():
    if not ppt:
        return

    NAME_PPT = 'testppt.pptx'
    path_ppt = os.path.join(TESTDIR, NAME_PPT)
    NAME_PNG = 'abc.png'
    path_png = os.path.join(TESTDIR, NAME_PNG)

    directive = pptshape.directive.PPTShape(None, None, None, 
             None, None, None, None, None, None)

    if os.path.exists(path_png):
        os.unlink(path_png)

    directive.arguments = [NAME_PNG]
    directive.state = mock.Mock()
    directive.state.document.current_source = os.path.join(TESTDIR, 'test.rst')
    directive.options = {
        'pptfilename': NAME_PPT,
        'shapename': '#2.1',
    }
    directive.run()
    assert os.path.exists(path_png)
    