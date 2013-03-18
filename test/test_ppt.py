import os, sys
import pptshape.ppt

TESTDIR = os.path.split(__file__)[0]

def test_open():
    ppt = pptshape.ppt.PPTShape(os.path.join(TESTDIR, 'testppt.pptx'))
    shape = ppt.saveShape('shape-title', 
                os.path.join(TESTDIR, 'aaa.png'))
    ppt.quit()


